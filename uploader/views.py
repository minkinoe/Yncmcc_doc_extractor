import os
import logging
import re
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
                # 保存原始文件名，供提取函数使用
                original_name = file.name

                with open(file_path, 'wb+') as destination:
                    for chunk in file.chunks():
                        destination.write(chunk)

                logger.info(f"开始处理文件: {file.name}")
                if is_zip:
                    results = extract_info_from_zip(file_path, original_name)
                else:
                    results = extract_info_from_word(file_path, original_name)

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
        raw = r.get('code_part') or r.get('file_name') or '未知'
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
