import os
import logging
from django.shortcuts import render, redirect
from django.conf import settings
from django.contrib import messages
from .utils import extract_info_from_zip, extract_info_from_word

# 配置日志
logger = logging.getLogger(__name__)

def upload_file(request):
    """文件上传页面"""
    if request.method == 'POST' and (request.FILES.get('zip_file') or request.FILES.get('word_file')):
        # 检查是ZIP文件还是Word文件
        file = request.FILES.get('zip_file') or request.FILES.get('word_file')
        is_zip = file.name.lower().endswith('.zip')
        is_word = file.name.lower().endswith('.doc') or file.name.lower().endswith('.docx')
        
        # 验证文件类型
        if not (is_zip or is_word):
            messages.error(request, '请上传ZIP格式或Word格式(.doc, .docx)的文件！')
            return render(request, 'uploader/upload.html')
        
        try:
            # 保存上传的文件
            os.makedirs(settings.MEDIA_ROOT, exist_ok=True)
            file_path = os.path.join(settings.MEDIA_ROOT, file.name)
            
            with open(file_path, 'wb+') as destination:
                try:
                    for chunk in file.chunks():
                        destination.write(chunk)
                except Exception as e:
                    logger.error(f"文件保存失败: {str(e)}")
                    messages.error(request, '文件上传保存失败，请重试！')
                    return render(request, 'uploader/upload.html')
            
            # 根据文件类型选择不同的处理函数
            logger.info(f"开始处理文件: {file.name}")
            if is_zip:
                results = extract_info_from_zip(file_path)
            else:
                results = extract_info_from_word(file_path)
            
            # 检查结果是否为空
            if not results:
                messages.warning(request, '未能从上传的文件中提取到任何有效信息！')
            else:
                # 统计成功和失败的数量
                success_count = sum(1 for r in results if r.get('extraction_status') == '成功' or 'extraction_status' not in r)
                error_count = len(results) - success_count
                if error_count > 0:
                    messages.warning(request, f'部分文件处理失败，成功: {success_count}, 失败: {error_count}')
                else:
                    messages.success(request, f'成功处理 {success_count} 个文档！')
            
            # 将结果存储在session中以便在结果页面显示
            request.session['extraction_results'] = results
            
        except Exception as e:
            logger.error(f"文件处理过程中出错: {str(e)}")
            messages.error(request, f'文件处理失败: {str(e)}')
            results = []
            request.session['extraction_results'] = results
            
        finally:
            # 确保临时文件被删除
            if 'file_path' in locals() and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception as e:
                    logger.warning(f"无法删除临时文件: {str(e)}")
            
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
