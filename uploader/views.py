import os
import logging
import re
import json
import zipfile
from functools import lru_cache
from xml.etree import ElementTree
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
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

def _normalize_street_name(value):
    if value is None:
        return ''
    s = str(value).strip()
    if not s:
        return ''
    s = s.replace(' ', '').replace('\u3000', '')
    if s.endswith('街道办事处'):
        s = s[:-5]
    return s

def _xlsx_first_sheet_xml_path(zf):
    workbook_xml = zf.read('xl/workbook.xml')
    wb_root = ElementTree.fromstring(workbook_xml)
    wb_ns = ''
    if wb_root.tag.startswith('{'):
        wb_ns = wb_root.tag.split('}')[0].lstrip('{')

    sheets = wb_root.find(f'.//{{{wb_ns}}}sheets') if wb_ns else wb_root.find('.//sheets')
    if sheets is None or len(list(sheets)) == 0:
        return None

    first_sheet = list(sheets)[0]
    rel_attr = None
    for k, v in first_sheet.attrib.items():
        if k.endswith('}id') or k == 'r:id':
            rel_attr = v
            break
    if not rel_attr:
        return 'xl/worksheets/sheet1.xml'

    rels_xml = zf.read('xl/_rels/workbook.xml.rels')
    rels_root = ElementTree.fromstring(rels_xml)
    rels_ns = ''
    if rels_root.tag.startswith('{'):
        rels_ns = rels_root.tag.split('}')[0].lstrip('{')

    for rel in rels_root.findall(f'.//{{{rels_ns}}}Relationship' if rels_ns else './/Relationship'):
        if rel.attrib.get('Id') == rel_attr:
            target = rel.attrib.get('Target') or ''
            target = target.lstrip('/')
            if target.startswith('xl/'):
                return target
            return f"xl/{target}"

    return 'xl/worksheets/sheet1.xml'

def _xlsx_shared_strings(zf):
    try:
        xml_bytes = zf.read('xl/sharedStrings.xml')
    except KeyError:
        return []

    root = ElementTree.fromstring(xml_bytes)
    ns = ''
    if root.tag.startswith('{'):
        ns = root.tag.split('}')[0].lstrip('{')

    strings = []
    for si in root.findall(f'.//{{{ns}}}si' if ns else './/si'):
        t_nodes = si.findall(f'.//{{{ns}}}t' if ns else './/t')
        text = ''.join([(t.text or '') for t in t_nodes])
        strings.append(text)
    return strings

def _xlsx_sheet_rows(zf, sheet_xml_path, shared_strings):
    xml_bytes = zf.read(sheet_xml_path)
    root = ElementTree.fromstring(xml_bytes)
    ns = ''
    if root.tag.startswith('{'):
        ns = root.tag.split('}')[0].lstrip('{')

    rows = []
    for row in root.findall(f'.//{{{ns}}}row' if ns else './/row'):
        row_values = {}
        for c in row.findall(f'.//{{{ns}}}c' if ns else './/c'):
            cell_ref = c.attrib.get('r') or ''
            col = ''.join([ch for ch in cell_ref if ch.isalpha()]).upper()
            cell_type = c.attrib.get('t')
            v_node = c.find(f'{{{ns}}}v' if ns else 'v')
            value = v_node.text if v_node is not None else None

            if cell_type == 's':
                try:
                    idx = int(value) if value is not None else -1
                    row_values[col] = shared_strings[idx] if 0 <= idx < len(shared_strings) else ''
                except (TypeError, ValueError):
                    row_values[col] = ''
            elif cell_type == 'inlineStr':
                t_node = c.find(f'.//{{{ns}}}t' if ns else './/t')
                row_values[col] = (t_node.text or '') if t_node is not None else ''
            else:
                row_values[col] = value or ''

        if row_values:
            rows.append(row_values)
    return rows

@lru_cache(maxsize=1)
def load_street_to_construction_unit_mapping():
    xlsx_path = getattr(settings, 'STREET_TEAM_XLSX_PATH', None)
    if not xlsx_path:
        return {}

    xlsx_path_str = str(xlsx_path)
    if not os.path.exists(xlsx_path_str):
        return {}

    try:
        with zipfile.ZipFile(xlsx_path_str, 'r') as zf:
            sheet_path = _xlsx_first_sheet_xml_path(zf)
            if not sheet_path:
                return {}

            shared_strings = _xlsx_shared_strings(zf)
            rows = _xlsx_sheet_rows(zf, sheet_path, shared_strings)

        if not rows:
            return {}

        header = rows[0]
        header_a = _normalize_street_name(header.get('A'))
        header_b = _normalize_street_name(header.get('B'))

        street_col = 'A'
        unit_col = 'B'
        if any(x in header_a for x in ['施工', '单位', '队']) and '街道' in header_b:
            street_col, unit_col = 'B', 'A'

        mapping = {}
        for r in rows[1:] if ('街道' in header_a or '街道' in header_b) else rows:
            street = _normalize_street_name(r.get(street_col))
            unit = (r.get(unit_col) or '').strip()
            if not street:
                continue
            mapping[street] = unit
        return mapping
    except (zipfile.BadZipFile, KeyError, ElementTree.ParseError, OSError):
        return {}

def get_construction_unit_from_township(township):
    street = _normalize_street_name(township)
    if not street:
        return None

    mapping = load_street_to_construction_unit_mapping()
    if not mapping:
        return None

    unit = mapping.get(street)
    if unit:
        return unit

    for k, v in mapping.items():
        if not k or not v:
            continue
        if street in k or k in street:
            return v

    return None

def _amap_get_json(endpoint, params, timeout_seconds=6):
    url = f"{endpoint}?{urlencode(params)}"
    req = Request(url, headers={'User-Agent': 'wordextractor/1.0'})
    with urlopen(req, timeout=timeout_seconds) as resp:
        raw = resp.read().decode('utf-8', errors='replace')
    return json.loads(raw)

def get_township_from_address(address):
    amap_key = getattr(settings, 'AMAP_API_KEY', None)
    if not amap_key:
        return None

    if not address:
        return None

    normalized_address = str(address).replace('+', ' ').strip()
    if not normalized_address:
        return None

    try:
        geo = _amap_get_json(
            'https://restapi.amap.com/v3/geocode/geo',
            {'key': amap_key, 'address': normalized_address},
        )
        if str(geo.get('status')) != '1':
            return None

        geocodes = geo.get('geocodes') or []
        if not geocodes:
            return None

        location = geocodes[0].get('location')
        if not location:
            return None

        regeo = _amap_get_json(
            'https://restapi.amap.com/v3/geocode/regeo',
            {'key': amap_key, 'location': location, 'radius': 1000, 'extensions': 'base'},
        )
        if str(regeo.get('status')) != '1':
            return None

        address_component = (regeo.get('regeocode') or {}).get('addressComponent') or {}
        township = address_component.get('township') or ''
        if township:
            return township

        street = ((address_component.get('streetNumber') or {}).get('street')) or ''
        if street:
            return street

        return None
    except (HTTPError, URLError, json.JSONDecodeError, TimeoutError, ValueError):
        return None

def dashboard(request, file_id=None):
    """统一的仪表盘视图，处理上传和显示结果"""
    # 获取历史记录供侧边栏使用
    history_list = UploadedFile.objects.all().order_by('-uploaded_at')[:50]
    
    context = {
        'history_list': history_list,
        'selected_file_id': file_id,
        'results': [],
        'unique_codes': [],
        'selected_upload_id': file_id
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
                zip_township = get_township_from_address(zip_address)
                zip_construction_unit = get_construction_unit_from_township(zip_township)
                uploaded_file = UploadedFile(
                    original_filename=file.name,
                    file_size=file.size,
                    file_type='zip',
                    group_name=zip_group_name or None,
                    address=zip_address or None,
                    township=zip_township or None,
                    construction_unit=zip_construction_unit or None,
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
            zip_township = uploaded_file.township
            if not zip_township and zip_address:
                zip_township = get_township_from_address(zip_address)
                if zip_township:
                    uploaded_file.township = zip_township
                    uploaded_file.save(update_fields=['township'])

            zip_construction_unit = uploaded_file.construction_unit
            if not zip_construction_unit and zip_township:
                zip_construction_unit = get_construction_unit_from_township(zip_township)
                if zip_construction_unit:
                    uploaded_file.construction_unit = zip_construction_unit
                    uploaded_file.save(update_fields=['construction_unit'])

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
                    'zip_township': zip_township,
                    'zip_construction_unit': zip_construction_unit,
                })
            
            context['results'] = results
            context['unique_codes'] = sorted(list(unique_codes))
            context['zip_order_code'] = zip_order_code
            context['zip_group_name'] = zip_group_name
            context['zip_address'] = zip_address
            context['zip_township'] = zip_township
            context['zip_construction_unit'] = zip_construction_unit
            context['selected_upload_id'] = uploaded_file.id
            
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

@require_POST
def update_construction_unit(request, file_id):
    try:
        uploaded_file = UploadedFile.objects.get(id=file_id)
    except UploadedFile.DoesNotExist:
        return JsonResponse({'ok': False, 'error': 'not_found'}, status=404)

    try:
        payload = json.loads((request.body or b'{}').decode('utf-8', errors='replace'))
    except json.JSONDecodeError:
        payload = {}

    unit = payload.get('construction_unit')
    unit = (str(unit).strip() if unit is not None else '')
    if unit == '':
        uploaded_file.construction_unit = None
    else:
        uploaded_file.construction_unit = unit[:255]

    uploaded_file.save(update_fields=['construction_unit'])
    return JsonResponse({'ok': True, 'construction_unit': uploaded_file.construction_unit})

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
