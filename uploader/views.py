import os
import logging
import re
from django.shortcuts import render, redirect
from django.conf import settings
from django.contrib import messages
from django.utils import timezone
from django.http import HttpResponse
from wsgiref.util import FileWrapper
import mimetypes
from .utils import extract_info_from_zip, extract_info_from_word
from .models import UploadedFile, ExtractedInfo

# 配置日志
logger = logging.getLogger(__name__)

def upload_file(request):
    """文件上传页面（支持多文件上传）"""
    if request.method == 'POST':
        # 支持 name='files' 的多文件上传；同时兼容旧的 zip_file/word_file 单文件字段
        uploaded_files = request.FILES.getlist('files') or []
        # 兼容：如果客户端仍使用旧字段，则也加入
        if not uploaded_files:
            if request.FILES.get('zip_file'):
                uploaded_files = [request.FILES.get('zip_file')]
            elif request.FILES.get('word_file'):
                uploaded_files = [request.FILES.get('word_file')]

        if not uploaded_files:
            messages.error(request, '请上传至少一个ZIP或Word文件(.zip, .doc, .docx)！')
            return render(request, 'uploader/upload.html')

        all_results = []
        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)

        for file in uploaded_files:
            # 验证文件类型
            name_lower = file.name.lower()
            is_zip = name_lower.endswith('.zip')
            is_word = name_lower.endswith('.doc') or name_lower.endswith('.docx')

            if not (is_zip or is_word):
                logger.warning(f"跳过不支持的文件类型: {file.name}")
                all_results.append({
                    'file_name': file.name,
                    'extraction_status': '失败',
                    'error': '不支持的文件类型'
                })
                continue

            # 将文件保存到数据库
            try:
                # 创建上传文件记录
                uploaded_file = UploadedFile(
                    original_filename=file.name,
                    file_size=file.size,
                    file_type='zip' if is_zip else 'word'
                )
                
                # 保存文件到存储系统
                uploaded_file.file = file
                uploaded_file.save()
                
                logger.info(f"开始处理文件: {file.name}, ID: {uploaded_file.id}")
                
                # 使用Django存储的文件路径进行处理
                file_path = uploaded_file.file.path
                
                # 处理文件并提取信息
                if is_zip:
                    results = extract_info_from_zip(file_path, file.name)
                else:
                    results = extract_info_from_word(file_path, file.name)
                
                # 处理提取结果
                if results:
                    all_results.extend(results)
                    # 将提取的信息保存到数据库
                    for result in results:
                        # 处理DecimalField所需的数据转换
                        maintenance_fee = float(result.get('maintenance_fee', 0))
                        service_fee = float(result.get('service_fee', 0))
                        terminal_fee = float(result.get('terminal_fee', 0))
                        total_fees = float(result.get('total_fees', 0))
                        
                        # 转换其他价格字段
                        doc_maintenance_total = float(result.get('doc_maintenance_total', 0)) if result.get('doc_maintenance_total') else None
                        overall_total_price = float(result.get('overall_total_price', 0)) if result.get('overall_total_price') else None
                        total_price = float(result.get('total_price', 0)) if result.get('total_price') else None
                        other_fees = float(result.get('other_fees', 0)) if result.get('other_fees') else 0
                        
                        # 创建提取信息记录
                        extracted_info = ExtractedInfo(
                            uploaded_file=uploaded_file,
                            order_code=result.get('order_code'),
                            document_name=result.get('file_name', ''),
                            document_content=result.get('document_content'),
                            extraction_status=result.get('extraction_status', '成功'),
                            extraction_error=result.get('error'),
                            maintenance_fee=maintenance_fee,
                            service_fee=service_fee,
                            terminal_fee=terminal_fee,
                            other_fees=other_fees,
                            total_fees=total_fees,
                            doc_maintenance_total=doc_maintenance_total,
                            overall_total_price=overall_total_price,
                            total_price=total_price,
                            fiber_info=result.get('fiber_info'),
                            equipment_items=result.get('equipment_items'),
                            verification_passed=result.get('verification_passed', False),
                            verification_message=result.get('verification_message')
                        )
                        extracted_info.save()
                    
                    # 更新文档数量
                    uploaded_file.document_count = len(results)
                else:
                    all_results.append({
                        'file_name': file.name,
                        'order_code': '未知',
                        'extraction_status': '失败',
                        'error': '未提取到任何信息'
                    })
                    
                # 标记文件已处理
                uploaded_file.is_processed = True
                uploaded_file.processed_at = timezone.now()
                uploaded_file.save()
                
            except Exception as e:
                logger.error(f"处理文件 {file.name} 时出错: {str(e)}")
                all_results.append({
                    'file_name': file.name,
                    'order_code': '未知',
                    'extraction_status': '失败',
                    'error': str(e)
                })
                
                # 如果创建了上传文件记录，更新错误信息
                if 'uploaded_file' in locals():
                    uploaded_file.is_processed = True
                    uploaded_file.processed_at = timezone.now()
                    uploaded_file.processing_error = str(e)
                    uploaded_file.save()

        # 检查总体结果并设置消息
        if not all_results:
            messages.warning(request, '未能从上传的文件中提取到任何有效信息！')
        else:
            success_count = sum(1 for r in all_results if r.get('extraction_status') == '成功' or 'extraction_status' not in r)
            error_count = len(all_results) - success_count
            if error_count > 0:
                messages.warning(request, f'部分文件处理失败，成功: {success_count}, 失败: {error_count}')
            else:
                messages.success(request, f'成功处理 {success_count} 个文档！')

        # 存储在 session 供结果页展示
        request.session['extraction_results'] = all_results
        return redirect('show_result')

    return render(request, 'uploader/upload.html')

def show_result(request):
    """显示提取结果页面"""
    results = request.session.get('extraction_results', [])
    
    # 如果没有session结果（可能是直接访问），可以重定向到历史记录页面
    if not results:
        return redirect('file_history')
    

    
    # 对结果进行分组，区分成功和失败的结果
    success_results = []
    error_results = []
    
    for result in results:
        if result.get('extraction_status') == '失败' or 'error' in result:
            error_results.append(result)
        else:
            success_results.append(result)
    
    # 清除session中的结果
    if 'extraction_results' in request.session:
        del request.session['extraction_results']
    
    # 准备返回给模板的数据
    context = {
        'results': success_results,  # 直接使用results变量名以匹配模板中的引用
        'error_results': error_results,
        'total_success': len(success_results),
        'total_error': len(error_results),
        'total_processed': len(results)
    }
    # 同时提供一个按单号分组的单号列表（包含成功与失败），供前端侧边栏展示
    all_results = success_results + error_results

    def _clean_order_code(name: str) -> str:
        """从文件名或 code_part 中提取干净的单号。

        优先匹配以 EOSC_ 开头的片段（例如 EOSC_4712510225808891_KC）。
        否则去掉类似 upload_<hex>_ 的前缀并移除扩展名，作为回退。
        """
        if not name:
            return '未知'

        # 去除扩展名
        base = os.path.splitext(name)[0]

        # 优先查找 EOSC_ 开头的单号片段
        m = re.search(r'(EOSC_[A-Za-z0-9_\-]+)', base)
        if m:
            return m.group(1)

        # 回退：去掉 upload_<hex>_ 前缀（如果存在）
        base2 = re.sub(r'^upload_[0-9a-fA-F]+_', '', base)
        return base2 or base

    unique_codes = []
    seen = set()
    for r in all_results:
        raw = r.get('order_code') or r.get('file_name') or '未知'
        code = _clean_order_code(raw)
        # 在每个结果中保存清洗后的显示单号，模板可以使用它来显示更友好的单号
        try:
            r['display_code'] = code
        except Exception:
            # 保持稳健：如果 r 不是可变映射，忽略
            pass
        if code not in seen:
            unique_codes.append(code)
            seen.add(code)

    context['all_results'] = all_results
    context['unique_codes'] = unique_codes
    
    return render(request, 'uploader/result.html', context)


def file_history(request):
    """显示历史上传记录"""
    # 获取所有上传文件记录，按上传时间倒序排列
    uploaded_files = UploadedFile.objects.all().order_by('-uploaded_at')
    
    # 计算统计信息
    total_files = uploaded_files.count()
    processed_files = uploaded_files.filter(is_processed=True).count()
    error_files = uploaded_files.filter(processing_error__isnull=False).count()
    
    # 准备上下文数据
    context = {
        'uploaded_files': uploaded_files,
        'total_files': total_files,
        'processed_files': processed_files,
        'error_files': error_files
    }
    
    return render(request, 'uploader/file_history.html', context)


def file_detail(request, file_id):
    """显示特定文件的详细提取结果"""
    try:
        # 获取上传文件记录
        uploaded_file = UploadedFile.objects.get(id=file_id)
        
        # 获取该文件的所有提取信息
        extracted_info_list = uploaded_file.extracted_infos.all()
        
        # 将数据库模型转换为字典格式，以适配result.html模板
        all_results = []
        success_results = []
        error_results = []
        
        for info in extracted_info_list:
            result = {
                'file_name': info.document_name,
                'order_code': info.order_code,
                'extraction_status': info.extraction_status,
                'error': info.extraction_error,
                'maintenance_fee': info.maintenance_fee,
                'service_fee': info.service_fee,
                'terminal_fee': info.terminal_fee,
                'total_fees': info.total_fees,
                'doc_maintenance_total': info.doc_maintenance_total,
                'overall_total_price': info.overall_total_price,
                'total_price': info.total_price,
                'fiber_info': info.fiber_info,
                'equipment_items': info.equipment_items,
                'verification_passed': info.verification_passed,
                'verification_message': info.verification_message,
                'document_content': info.document_content,
            }
            all_results.append(result)
            
            if info.extraction_status == '失败' or info.extraction_error:
                error_results.append(result)
            else:
                success_results.append(result)

        # 准备返回给模板的数据
        context = {
            'results': success_results,
            'error_results': error_results,
            'all_results': all_results,
            'total_success': len(success_results),
            'total_error': len(error_results),
            'total_processed': len(all_results),
            'uploaded_file': uploaded_file,  # 传递原始文件信息
            'back_url': 'file_history', # 标记返回路径
        }
        
        # 生成单号列表逻辑 (复用show_result的逻辑)
        def _clean_order_code(name: str) -> str:
            if not name:
                return '未知'
            base = os.path.splitext(name)[0]
            m = re.search(r'(EOSC_[A-Za-z0-9_\-]+)', base)
            if m:
                return m.group(1)
            base2 = re.sub(r'^upload_[0-9a-fA-F]+_', '', base)
            return base2 or base

        unique_codes = []
        seen = set()
        for r in all_results:
            raw = r.get('order_code') or r.get('file_name') or '未知'
            code = _clean_order_code(raw)
            try:
                r['display_code'] = code
            except Exception:
                pass
            if code not in seen:
                unique_codes.append(code)
                seen.add(code)

        context['unique_codes'] = unique_codes
        
        return render(request, 'uploader/result.html', context)
        
    except UploadedFile.DoesNotExist:
        messages.error(request, '找不到指定的文件记录')
        return redirect('file_history')
    except Exception as e:
        logger.error(f"查看文件详情时出错: {str(e)}")
        messages.error(request, f'查看文件详情时出错: {str(e)}')
        return redirect('file_history')


def download_file(request, file_id):
    """下载上传的文件"""
    try:
        # 获取上传文件记录
        uploaded_file = UploadedFile.objects.get(id=file_id)
        
        # 检查文件是否存在
        if not uploaded_file.file or not os.path.exists(uploaded_file.file.path):
            messages.error(request, '文件不存在或已被删除')
            return redirect('file_detail', file_id=file_id)
        
        # 获取文件的MIME类型
        content_type = mimetypes.guess_type(uploaded_file.file.path)[0] or 'application/octet-stream'
        
        # 创建文件包装器
        file_wrapper = FileWrapper(open(uploaded_file.file.path, 'rb'))
        
        # 创建响应
        response = HttpResponse(file_wrapper, content_type=content_type)
        response['Content-Length'] = os.path.getsize(uploaded_file.file.path)
        response['Content-Disposition'] = f'attachment; filename="{uploaded_file.original_filename}"'
        
        return response
        
    except UploadedFile.DoesNotExist:
        messages.error(request, '找不到指定的文件记录')
        return redirect('file_history')
    except Exception as e:
        logger.error(f"下载文件时出错: {str(e)}")
        messages.error(request, f'下载文件时出错: {str(e)}')
        return redirect('file_detail', file_id=file_id)
