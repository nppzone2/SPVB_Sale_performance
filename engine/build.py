"""Build site/index.html from input/*.xlsx and src/template.html.

- Sale Out (BC Sản lượng theo KH theo SKU), ASO list (Danh sách chi tiết ASO),
  Loyalty list (DS ASO Tham gia Loyalty) are read from input/ and embedded
  as compressed JSON. All calculations stay in the browser, as before.
- Target file (Target_*.xlsx, exported from tab Nhập Target) is embedded as
  the shared target set.
"""
import base64, datetime, glob, gzip, json, os, sys, unicodedata
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_DIR = os.path.join(ROOT, 'input')
SRC = os.path.join(ROOT, 'src', 'template.html')
OUT = os.path.join(ROOT, 'site', 'index.html')

KEEP = {
    'sale': ['Region', 'Group', 'ASM', 'Mã NPP', 'Tên NPP', 'SS', 'Mã Route', 'Tên Route', 'Loại Route', 'Mã KH', 'Tên KH',
             'Loại hình chi tiết', 'Segment', 'Ngành', 'Nhãn', 'Hương', 'Bao Bì', 'Dung Tích', 'Sản Lượng', 'Tần suất VT'],
    'loy': ['Khu vực', 'Area', 'Mã Chương Trình', 'Tên Chương Trình', 'Mức Tham Gia', 'Mã NPP', 'Tên NPP', 'Mã Route',
            'Mã Khách Hàng', 'Tên Khách Hàng', 'Segment', 'Số Suất', 'Trạng Thái',
            'Chỉ Tiêu (Đ/v các CT DCR tự nhập chỉ tiêu)'],
    'aso': ['Khu vực', 'Area', 'Mã NPP', 'Tên NPP', 'Mã KH', 'Tên KH', 'Club', 'Segment', 'Route Code', 'DCR Name',
            'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'Trang thai KH', 'ASO Status'],
}
TARGET_COLS = {'sale': 'sale', 'aso': 'aso', 'sdo': 'sdo'}


def fold(s):
    s = unicodedata.normalize('NFD', str(s)).encode('ascii', 'ignore').decode().lower()
    return ''.join(ch if ch.isalnum() else '_' for ch in s)


def classify(path):
    n = fold(os.path.basename(path))
    if 'target' in n:
        return 'target'
    if 'loyal' in n:
        return 'loy'
    if 'danh_sach' in n or 'aso' in n and 'danh' in n:
        return 'aso'
    if 'san_luong' in n or 'sku' in n:
        return 'sale'
    return None


def cell(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime('%Y-%m-%d')
    return str(v).strip()


def read_table(path, keep):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    head_i, head = None, None
    rows_iter = ws.iter_rows(values_only=True)
    buf = []
    for i, r in enumerate(rows_iter):
        buf.append(r)
        names = {cell(c) for c in r}
        if len(names & set(keep)) >= max(3, len(keep) // 2):
            head_i, head = i, [cell(c) for c in r]
            break
        if i > 80:
            break
    if head is None:
        raise SystemExit(f'Không tìm thấy dòng tiêu đề trong {os.path.basename(path)}')
    idx = [head.index(c) if c in head else None for c in keep]
    out = []
    for r in rows_iter:
        if r is None or all(c is None or str(c).strip() == '' for c in r):
            continue
        out.append([cell(r[j]) if j is not None and j < len(r) else '' for j in idx])
    wb.close()
    print(f'  {os.path.basename(path)}: {len(out)} dòng, thiếu cột: {[c for c, j in zip(keep, idx) if j is None] or "không"}')
    return [keep] + out


def read_targets(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    head, code_i, cols = None, None, {}
    for r in ws.iter_rows(values_only=True):
        h = [cell(c) for c in r]
        if head is None and 'Mã route' in h:
            head = h
            code_i = h.index('Mã route')
            for j, name in enumerate(h):
                n = fold(name)
                if 'target' in n and 'sale' in n: cols[j] = 'sale'
                elif 'target' in n and 'aso' in n: cols[j] = 'aso'
                elif 'target' in n and 'sdo' in n: cols[j] = 'sdo'
            continue
        if head is None or not cols:
            continue
        code = cell(r[code_i]) if code_i < len(r) else ''
        if not code:
            continue
        for j, k in cols.items():
            v = r[j] if j < len(r) else None
            if v is None or str(v).strip() == '':
                continue
            try:
                num = float(str(v).replace(',', ''))
            except ValueError:
                continue
            TGT.setdefault(k, {})[code] = int(num) if num.is_integer() else num
    wb.close()
    print(f'  {os.path.basename(path)}: ' + ', '.join(f'{k}={len(TGT.get(k, {}))}' for k in ('sale', 'aso', 'sdo')))


TGT = {}


def main():
    files = sorted(glob.glob(os.path.join(IN_DIR, '*.xlsx')))
    found = {}
    for f in files:
        k = classify(f)
        if k:
            found.setdefault(k, []).append(f)
    print('File đầu vào:')
    missing = [k for k in ('sale', 'loy', 'aso') if k not in found]
    if missing:
        sys.exit(f'Thiếu file cho: {missing}. Đặt file vào thư mục input/.')
    data = {}
    for k in ('sale', 'loy', 'aso'):
        if len(found[k]) > 1:
            sys.exit(f'Có nhiều file cho {k}: {found[k]}')
        data[k] = read_table(found[k][0], KEEP[k])
    tjson = os.path.join(ROOT, 'targets', 'targets.json')
    if os.path.exists(tjson):
        saved = json.load(open(tjson, encoding='utf-8'))
        for k in ('sale', 'aso', 'sdo'):
            TGT[k] = {str(c): v for c, v in (saved.get(k) or {}).items() if v not in ('', None)}
        print('  targets/targets.json: ' + ', '.join(f'{k}={len(TGT.get(k, {}))}' for k in ('sale', 'aso', 'sdo')))
    else:
        for f in found.get('target', []):
            read_targets(f)
    payload = base64.b64encode(gzip.compress(json.dumps(data, ensure_ascii=False, separators=(',', ':')).encode('utf-8'), 9)).decode()
    tpl = open(SRC, encoding='utf-8').read()
    assert tpl.count('__SAMPLE__') == 1 and tpl.count('__TARGETS__') == 1
    html = tpl.replace('__SAMPLE__', payload).replace('__TARGETS__', json.dumps(TGT, ensure_ascii=False))
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    open(OUT, 'w', encoding='utf-8').write(html)
    tj = os.path.join(ROOT, 'targets', 'targets.json')
    if os.path.exists(tj):
        import shutil
        shutil.copy(tj, os.path.join(os.path.dirname(OUT), 'targets.json'))
    print(f'Đã tạo {os.path.relpath(OUT, ROOT)} ({len(html) // 1024} KB)')


if __name__ == '__main__':
    main()
