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

def dashboard(request, file_id=None):
    """统一的仪表盘视图，处理上传和显示结果"""
    # 获取历史记录供侧边栏使用
    history_list = UploadedFile.objects.all().order_by('-uploaded_at')[:50]
    
    context = {
        'history_list': history_list,
        'selected_file_id': file_id,
        'results': [],
        'unique_codes': []
    }

    # 处理上传
    if request.method == 'POST':
        # 复用 upload_file 的核心逻辑
        uploaded_files = request.FILES.getlist('files') or []
        if not uploaded_files:
            # 兼容旧字段
            if request.FILES.get('zip_file'): uploaded_files = [request.FILES.get('zip_file')]
            elif request.FILES.get('word_file'): uploaded_files = [request.FILES.get('word_file')]

        if not uploaded_files:
            messages.error(request, '请上传至少一个文件！')
            return redirect('dashboard')

        # 处理文件
        last_processed_id = None
        for file in uploaded_files:
            try:
                # 验证文件类型
                name_lower = file.name.lower()
                is_zip = name_lower.endswith('.zip')
                is_word = name_lower.endswith('.doc') or name_lower.endswith('.docx')

                if not (is_zip or is_word): continue

                # 创建记录
                uploaded_file = UploadedFile(
                    original_filename=file.name,
                    file_size=file.size,
                    file_type='zip' if is_zip else 'word'
                )
                uploaded_file.file = file
                uploaded_file.save()
                
                last_processed_id = uploaded_file.id
                file_path = uploaded_file.file.path
                
                # 提取
                if is_zip:
                    results = extract_info_from_zip(file_path, file.name)
                else:
                    results = extract_info_from_word(file_path, file.name)
                
                # 保存结果到数据库
                if results:
                    for result in results:
                        maintenance_fee = float(result.get('maintenance_fee', 0))
                        service_fee = float(result.get('service_fee', 0))
                        terminal_fee = float(result.get('terminal_fee', 0))
                        total_fees = float(result.get('total_fees', 0))
                        doc_maintenance_total = float(result.get('doc_maintenance_total', 0)) if result.get('doc_maintenance_total') else None
                        overall_total_price = float(result.get('overall_total_price', 0)) if result.get('overall_total_price') else None
                        total_price = float(result.get('total_price', 0)) if result.get('total_price') else None
                        other_fees = float(result.get('other_fees', 0)) if result.get('other_fees') else 0
                        
                        ExtractedInfo.objects.create(
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
                    uploaded_file.document_count = len(results)
                
                uploaded_file.is_processed = True
                uploaded_file.processed_at = timezone.now()
                uploaded_file.save()

            except Exception as e:
                logger.error(f"Error processing {file.name}: {e}")
                if 'uploaded_file' in locals():
                    uploaded_file.processing_error = str(e)
                    uploaded_file.save()

        # 上传完成后，重定向到该文件的详情页（如果只上传了一个，或者是最后一个）
        if last_processed_id:
            return redirect('dashboard_with_id', file_id=last_processed_id)
        return redirect('dashboard')

    # 处理显示结果 (GET)
    if file_id:
        try:
            uploaded_file = UploadedFile.objects.get(id=file_id)
            extracted_infos = uploaded_file.extracted_infos.all()
            
            # 构造结果列表
            results = []
            unique_codes = set()
            
            for info in extracted_infos:
                # 确定显示单号
                code = info.order_code
                if not code or code == '未知':
                    # 尝试从文件名提取简化的单号用于过滤
                    base = os.path.splitext(info.document_name)[0]
                    m = re.search(r'(EOSC_[A-Za-z0-9_\-]+)', base)
                    code = m.group(1) if m else base
                
                unique_codes.add(code)
                
                results.append({
                    'file_name': info.document_name,
                    'order_code': info.order_code,
                    'display_code': code, # 用于前端过滤
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
                    'document_content': info.document_content,
                })
            
            context['results'] = results
            context['unique_codes'] = sorted(list(unique_codes))
            
        except UploadedFile.DoesNotExist:
            pass

    return render(request, 'uploader/dashboard.html', context)

def upload_file(request):
    """(已弃用) 文件上传页面"""
    return redirect('dashboard')

def show_result(request):
    """(已弃用) 显示提取结果页面"""
    return redirect('dashboard')

def file_history(request):
    """(已弃用) 显示历史上传记录"""
    return redirect('dashboard')

def file_detail(request, file_id):
    """(已弃用) 显示特定文件的详细提取结果"""
    return redirect('dashboard_with_id', file_id=file_id)


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
