"""Build site/index.html and site/data.enc.

Input (all encrypted with the group password, AES-GCM, key = PBKDF2-SHA256):
  input/sale.xlsx.enc   BC Sản lượng theo KH theo SKU
  input/loy.xlsx.enc    DS ASO tham gia Loyalty
  input/aso.xlsx.enc    Danh sách chi tiết ASO
  targets/targets.enc   Target (gửi từ trang)

The password is read from env SPVB_PASSWORD (GitHub Actions secret).
Output: site/data.enc (encrypted payload, decrypted in the browser) and
site/index.html (template only, no plain data).
Calculations stay in the browser, as before.
"""
import gzip, hashlib, io, json, os, sys, unicodedata
import openpyxl
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN_DIR = os.path.join(ROOT, 'input')
TGT_FILE = os.path.join(ROOT, 'targets', 'targets.enc')
SRC = os.path.join(ROOT, 'src', 'template.html')
OUT_DIR = os.path.join(ROOT, 'site')
ITER = 200000

KEEP = {
    'sale': ['Region', 'Group', 'ASM', 'Mã NPP', 'Tên NPP', 'SS', 'Mã Route', 'Tên Route', 'Loại Route', 'Mã KH', 'Tên KH',
             'Loại hình chi tiết', 'Segment', 'Ngành', 'Nhãn', 'Hương', 'Bao Bì', 'Dung Tích', 'Sản Lượng', 'Tần suất VT'],
    'loy': ['Khu vực', 'Area', 'Mã Chương Trình', 'Tên Chương Trình', 'Mức Tham Gia', 'Mã NPP', 'Tên NPP', 'Mã Route',
            'Mã Khách Hàng', 'Tên Khách Hàng', 'Segment', 'Số Suất', 'Trạng Thái',
            'Chỉ Tiêu (Đ/v các CT DCR tự nhập chỉ tiêu)'],
    'aso': ['Khu vực', 'Area', 'Mã NPP', 'Tên NPP', 'Mã KH', 'Tên KH', 'Club', 'Segment', 'Route Code', 'DCR Name',
            'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'Trang thai KH', 'ASO Status'],
}
# Dấu nhận biết từng loại (đọc theo nội dung, không theo tên file)
MARK = [('sale', 'Sản Lượng'), ('loy', 'Mã Chương Trình'), ('aso', 'Route Code'), ('aso', 'ASO Status')]


def key_of(pw, salt):
    return hashlib.pbkdf2_hmac('sha256', pw.encode('utf-8'), salt, ITER, dklen=32)


def encrypt(pw, data: bytes) -> bytes:
    salt, iv = os.urandom(16), os.urandom(12)
    return salt + iv + AESGCM(key_of(pw, salt)).encrypt(iv, data, None)


def decrypt(pw, blob: bytes) -> bytes:
    salt, iv, ct = blob[:16], blob[16:28], blob[28:]
    return AESGCM(key_of(pw, salt)).decrypt(iv, ct, None)


def cell(v):
    if v is None:
        return ''
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (int,)):
        return str(v)
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    return str(v).strip()


def read_source(xlsx_bytes, label):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    kind, head_i, head = None, None, None
    rows = ws.iter_rows(values_only=True)
    for i, r in enumerate(rows):
        names = [cell(c) for c in r]
        for k, mark in MARK:
            if mark in names:
                kind, head_i, head = k, i, names
                break
        if kind:
            break
        if i > 80:
            break
    if not kind:
        sys.exit(f'Không nhận ra loại file: {label}')
    idx = [head.index(c) if c in head else None for c in KEEP[kind]]
    out = []
    for r in rows:
        if r is None or all(c is None or str(c).strip() == '' for c in r):
            continue
        out.append([cell(r[j]) if j is not None and j < len(r) else '' for j in idx])
    wb.close()
    missing = [c for c, j in zip(KEEP[kind], idx) if j is None]
    print(f'  {label} -> {kind}: {len(out)} dòng, thiếu cột: {missing or "không"}')
    return kind, [KEEP[kind]] + out


def read_targets(pw):
    if not os.path.exists(TGT_FILE):
        return {}
    data = json.loads(gzip.decompress(decrypt(pw, open(TGT_FILE, 'rb').read())))
    out = {k: {str(c): v for c, v in (data.get(k) or {}).items() if v not in ('', None)} for k in ('sale', 'aso', 'sdo')}
    print('  targets: ' + ', '.join(f'{k}={len(out[k])}' for k in out))
    return out


def main():
    pw = os.environ.get('SPVB_PASSWORD', '')
    if not pw:
        sys.exit('Thiếu mật khẩu: thêm secret SPVB_PASSWORD trong Settings → Secrets and variables → Actions.')
    plain = [f for f in os.listdir(IN_DIR) if f.lower().endswith('.xlsx')] if os.path.isdir(IN_DIR) else []
    if plain:
        sys.exit(f'Có file Excel chưa mã hoá trong input/: {plain}. Chỉ được đưa file đã mã hoá (.xlsx.enc).')
    data, seen = {}, {}
    print('Dữ liệu đầu vào:')
    for k in ('sale', 'loy', 'aso'):
        path = os.path.join(IN_DIR, f'{k}.xlsx.enc')
        if not os.path.exists(path):
            sys.exit(f'Thiếu file input/{k}.xlsx.enc')
        raw = decrypt(pw, open(path, 'rb').read())
        kind, rows = read_source(raw, f'{k}.xlsx.enc')
        if kind != k:
            sys.exit(f'File {k}.xlsx.enc thực tế là loại "{kind}". Hãy gửi lại đúng file.')
        data[k] = rows
    data['targets'] = read_targets(pw)
    payload = encrypt(pw, gzip.compress(json.dumps(data, ensure_ascii=False, separators=(',', ':')).encode('utf-8'), 9))
    tpl = open(SRC, encoding='utf-8').read()
    assert tpl.count('__SAMPLE__') == 1 and tpl.count('__TARGETS__') == 1
    html = tpl.replace('__SAMPLE__', '').replace('__TARGETS__', '{}')
    os.makedirs(OUT_DIR, exist_ok=True)
    open(os.path.join(OUT_DIR, 'data.enc'), 'wb').write(payload)
    open(os.path.join(OUT_DIR, 'index.html'), 'w', encoding='utf-8').write(html)
    print(f'Đã tạo site/index.html ({len(html) // 1024} KB) và site/data.enc ({len(payload) // 1024} KB)')


if __name__ == '__main__':
    main()
