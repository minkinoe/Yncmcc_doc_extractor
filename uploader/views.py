import os
import logging
from django.shortcuts import render, redirect
from django.conf import settings
from django.contrib import messages
from .utils import extract_info_from_zip, extract_info_from_word

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

            # 将上传文件保存到临时文件
            try:
                # 使用唯一文件名以避免冲突
                import uuid
                tmp_name = f"upload_{uuid.uuid4().hex}_{file.name}"
                file_path = os.path.join(settings.MEDIA_ROOT, tmp_name)

                with open(file_path, 'wb+') as destination:
                    for chunk in file.chunks():
                        destination.write(chunk)

                logger.info(f"开始处理文件: {file.name}")
                if is_zip:
                    results = extract_info_from_zip(file_path)
                else:
                    results = extract_info_from_word(file_path)

                # results 可能是列表（zip 情况），也可能是单个文件的列表，统一合并
                if results:
                    all_results.extend(results)
                else:
                    all_results.append({
                        'file_name': file.name,
                        'extraction_status': '失败',
                        'error': '未提取到任何信息'
                    })

            except Exception as e:
                logger.error(f"处理文件 {file.name} 时出错: {str(e)}")
                all_results.append({
                    'file_name': file.name,
                    'extraction_status': '失败',
                    'error': str(e)
                })

            finally:
                # 尝试删除临时文件
                try:
                    if 'file_path' in locals() and os.path.exists(file_path):
                        os.remove(file_path)
                except Exception as e:
                    logger.warning(f"无法删除临时文件 {file_path}: {str(e)}")

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
    
    return render(request, 'uploader/result.html', context)
