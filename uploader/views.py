import os
import logging
import re
from django.shortcuts import render, redirect
from django.conf import settings
from django.contrib import messages
from django.utils import timezone
from django.http import HttpResponse, JsonResponse
from django.views.decorators.http import require_POST
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
                if not is_zip:
                    messages.error(request, f'仅支持ZIP文件: {file.name}')
                    continue

                # 创建记录
                filename_base = os.path.splitext(file.name)[0]
                filename_base = re.sub(r'\(\d+\)$', '', filename_base)
                parts = filename_base.split('+')
                zip_group_name = parts[1] if len(parts) > 1 else ''
                zip_address = '+'.join(parts[2:]) if len(parts) > 2 else ''
                uploaded_file = UploadedFile(
                    original_filename=file.name,
                    file_size=file.size,
                    file_type='zip',
                    group_name=zip_group_name or None,
                    address=zip_address or None,
                    is_marked=True
                )
                uploaded_file.file = file
                uploaded_file.save()
                
                last_processed_id = uploaded_file.id
                file_path = uploaded_file.file.path
                
                # 提取
                results = extract_info_from_zip(file_path, file.name)
                
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
            filename_base = os.path.splitext(uploaded_file.original_filename)[0]
            filename_base = re.sub(r'\(\d+\)$', '', filename_base)
            parts = filename_base.split('+')
            zip_order_code = parts[0] if len(parts) > 0 else ''
            zip_group_name = parts[1] if len(parts) > 1 else ''
            zip_address = '+'.join(parts[2:]) if len(parts) > 2 else ''
            if uploaded_file.group_name:
                zip_group_name = uploaded_file.group_name
            if uploaded_file.address:
                zip_address = uploaded_file.address
            
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
                    'zip_order_code': zip_order_code,
                    'zip_group_name': zip_group_name,
                    'zip_address': zip_address,
                })
            
            context['results'] = results
            context['unique_codes'] = sorted(list(unique_codes))
            context['zip_order_code'] = zip_order_code
            context['zip_group_name'] = zip_group_name
            context['zip_address'] = zip_address
            
        except UploadedFile.DoesNotExist:
            pass

    return render(request, 'uploader/dashboard.html', context)


@require_POST
def toggle_upload_mark(request, file_id):
    try:
        uploaded_file = UploadedFile.objects.get(id=file_id)
    except UploadedFile.DoesNotExist:
        return JsonResponse({'ok': False, 'error': 'not_found'}, status=404)

    uploaded_file.is_marked = not uploaded_file.is_marked
    uploaded_file.save(update_fields=['is_marked'])
    return JsonResponse({'ok': True, 'is_marked': uploaded_file.is_marked})

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
