"""
月銷售排行整合引擎
將各電子書平台原始報表解析、彙整，並依範例格式輸出 Excel
"""

import openpyxl
import xlrd
import re
import io
from collections import defaultdict
from pathlib import Path
from datetime import date

# ── 書種對照 ────────────────────────────────────────────────
CAT_MAP = {
    '漫畫': '漫畫',
    '輕小說': '輕小說',
    '文學': '文學．小說',
    '實用書': '一般．實用書',
    '雜誌': '雜誌',
}

# ── 工具函式 ────────────────────────────────────────────────

def _to_bytesio(file_obj):
    """將 Streamlit UploadedFile 或一般 bytes 轉為 BytesIO"""
    if file_obj is None:
        return None
    if hasattr(file_obj, 'getvalue'):       # Streamlit UploadedFile
        return io.BytesIO(file_obj.getvalue())
    if hasattr(file_obj, 'read'):
        content = file_obj.read()
        if hasattr(file_obj, 'seek'):
            file_obj.seek(0)
        return io.BytesIO(content)
    if isinstance(file_obj, (bytes, bytearray)):
        return io.BytesIO(file_obj)
    return file_obj


def _col_letter(col_idx):
    """欄索引（1-based）→ Excel 欄位字母"""
    result = ''
    while col_idx > 0:
        col_idx, r = divmod(col_idx - 1, 26)
        result = chr(65 + r) + result
    return result


def load_genre_catalog(catalog_path=None):
    """
    載入 BW 書種目錄。
    catalog_path 可為：檔案路徑(str/Path)、BytesIO 或 Streamlit UploadedFile。
    找不到或解析失敗則回傳空字典（書種欄位將留空）。
    """
    if catalog_path is None:
        base = Path(__file__).parent.parent
        catalog_path = (
            base / "銷售報表" / "Bookwalker書籍商品ID&資料" / "product_fixed.xlsx"
        )

    # 若傳入的是檔案物件（BytesIO / UploadedFile），直接用內容載入
    if hasattr(catalog_path, 'read') or isinstance(catalog_path, io.IOBase):
        source = _to_bytesio(catalog_path)
    else:
        if not Path(str(catalog_path)).exists():
            return {}
        source = str(catalog_path)

    wb = openpyxl.load_workbook(source, read_only=True)
    ws = wb.active
    title_cat = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 2:
            continue
        title = row[3]
        category = row[10]
        if title and category:
            main_cat = str(category).split(' > ')[0]
            title_cat[str(title)] = CAT_MAP.get(main_cat, main_cat)
    wb.close()
    return title_cat


def get_genre(title, catalog):
    """多層次書種查詢：精確→去特典標記→去集數→空白"""
    t = str(title).strip()
    checks = [
        t,
        re.sub(r'\s*【[^】]*】\s*$', '', t).strip(),
        re.sub(r'\s*[\(（]\d+[\)）]\s*$',
               '', re.sub(r'\s*【[^】]*】\s*$', '', t).strip()).strip(),
    ]
    for c in checks:
        g = catalog.get(c, '')
        if g:
            return g
    return ''


# ── 各平台解析 ────────────────────────────────────────────────

def parse_google(google_file):
    """
    Google 稅務試算 xlsx
    使用 '應or免稅' 分頁（已樞紐彙整）
    回傳 list of (publisher, isbn, title, qty, amount)
    """
    wb = openpyxl.load_workbook(_to_bytesio(google_file), data_only=True)
    target = next((s for s in wb.sheetnames if '應or免稅' in s), None)
    if not target:
        raise ValueError("Google 檔案：找不到含「應or免稅」的分頁")

    ws = wb[target]
    agg = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 3:
            continue
        if row[0] is None:
            continue
        pub   = str(row[0]).strip()
        isbn  = str(row[1]) if row[1] else ''
        title = str(row[2]).strip() if row[2] else ''
        if not title or pub == '總計' or pub.endswith(' 合計'):
            continue
        qty = float(row[3] or 0)
        amt = float(row[4] or 0)
        key = (pub, isbn, title)
        if key not in agg:
            agg[key] = [0.0, 0.0]
        agg[key][0] += qty
        agg[key][1] += amt

    result = [(k[0], k[1], k[2], v[0], v[1]) for k, v in agg.items()]
    result.sort(key=lambda x: -x[4])
    return result


def parse_kobo(kobo_files):
    """
    Kobo：3 個 xlsx（BOOKWALKER / GAEABOOKS / KADOKAWA）
    自動辨識，合併 Kobo 分頁
    回傳 list of (publisher, eisbn, title, qty, amount)
    """
    agg = {}
    for f in kobo_files:
        fname = getattr(f, 'name', f'kobo_{id(f)}')
        bio = _to_bytesio(f)
        wb = openpyxl.load_workbook(bio)
        if 'Kobo' not in wb.sheetnames:
            continue
        ws = wb['Kobo']
        empty_eisbn_counter = 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            if row[0] is None:
                continue
            pub   = str(row[1]).strip()
            eisbn = str(row[6]).strip() if row[6] else ''
            title = str(row[8]).strip() if row[8] else ''
            qty   = float(row[5] or 0)
            amt   = float(row[14] or 0)
            # eISBN 空白時補流水號，避免同書名不同書種被誤合併
            if not eisbn:
                empty_eisbn_counter += 1
                eisbn = f'__no_eisbn_{fname}_{empty_eisbn_counter}__'
            key = (pub, eisbn, title)
            if key not in agg:
                agg[key] = [0.0, 0.0]
            agg[key][0] += qty
            agg[key][1] += amt

    # 輸出時還原假 eISBN 為空字串
    result = [(k[0], k[1] if not k[1].startswith('__no_eisbn_') else '', k[2], v[0], v[1])
              for k, v in agg.items()]
    result.sort(key=lambda x: -x[3])
    return result


def parse_readmoo(readmoo_file):
    """
    Readmoo 拆帳報表
    欄位：出版社(1), 書名(2), 排除免費冊數(13), 總金額(12)
    回傳 list of (publisher, '', title, qty, amount)
    """
    wb = openpyxl.load_workbook(_to_bytesio(readmoo_file), data_only=True)
    ws = wb.active
    agg = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 3:
            continue
        if row[1] is None:
            continue
        pub   = str(row[1]).strip()
        title = str(row[2]).strip() if row[2] else ''
        qty   = float(row[13] or 0)
        amt   = float(row[12] or 0)
        key = (pub, title)
        if key not in agg:
            agg[key] = [0.0, 0.0]
        agg[key][0] += qty
        agg[key][1] += amt

    result = [(k[0], '', k[1], v[0], v[1]) for k, v in agg.items()]
    result.sort(key=lambda x: -x[4])
    return result


def _find_isbn_col(header_row):
    """從標題列找 ISBN / 書號 / 商品代碼 欄位索引"""
    for j, h in enumerate(header_row):
        h_str = str(h)
        if any(kw in h_str for kw in ['ISBN', '書號', '商品代碼', '商品號']):
            return j
    return None


def parse_bokuto(bokuto_files):
    """
    博客來：bookwalker.xls（all 分頁）+ bookwalker1.xls（sale_report(O)，雜誌）
    自動依檔名辨識（含 'bookwalker1' → 雜誌）
    key 改用 (pub, isbn, title) 避免同書名不同書種被合併
    回傳 list of (publisher, isbn, title, qty, amount)
    """
    agg = {}
    for f in bokuto_files:
        fname = getattr(f, 'name', '').lower()
        is_mag = 'bookwalker1' in fname
        content = f.getvalue() if hasattr(f, 'getvalue') else f.read()
        wb = xlrd.open_workbook(file_contents=content)

        sheet_name = 'sale_report(O)' if is_mag else (
            'all' if 'all' in wb.sheet_names() else 'sale_report(O)'
        )
        if sheet_name not in wb.sheet_names():
            continue

        ws = wb.sheet_by_name(sheet_name)
        header_found = False
        isbn_col = None
        for i in range(ws.nrows):
            row = ws.row_values(i)
            if str(row[0]).startswith('結算帳月') or str(row[2]) == '廠商名稱':
                header_found = True
                isbn_col = _find_isbn_col(row)
                continue
            if not header_found or not row[0]:
                continue
            pub   = str(row[2]).strip()
            title = str(row[3]).strip()
            isbn  = str(row[isbn_col]).strip() if isbn_col is not None and row[isbn_col] else ''
            qty   = float(row[14] or 0) if is_mag else (
                    float(row[17] or 0) if len(row) > 17 else float(row[14] or 0))
            amt   = float(row[15] or 0)
            key = (pub, isbn, title)
            if key not in agg:
                agg[key] = [0.0, 0.0]
            agg[key][0] += qty
            agg[key][1] += amt

    result = [(k[0], k[1], k[2], v[0], v[1]) for k, v in agg.items()]
    result.sort(key=lambda x: -x[3])
    return result


def parse_luanda(luanda_file):
    """
    亂搭 明細 分頁
    金額手動計算：售價 × 數量 × (1 - 手續費) × 拆帳比例
    回傳 list of (publisher, '', title, qty, amount)
    """
    wb = openpyxl.load_workbook(_to_bytesio(luanda_file), data_only=True)
    ws = wb['明細']
    agg = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if row[8] is None:
            continue
        pub   = str(row[8]).strip()
        title = str(row[6]).strip() if row[6] else ''
        qty   = float(row[12] or 0)
        if qty == 0:
            continue  # 跳過零銷售列
        price = float(row[11] or 0)
        fee   = float(row[14] or 0)
        ratio = float(row[15] or 0)
        amt   = price * qty * (1 - fee) * ratio
        key = (pub, title)
        if key not in agg:
            agg[key] = [0.0, 0.0]
        agg[key][0] += qty
        agg[key][1] += amt

    result = [(k[0], '', k[1], v[0], v[1]) for k, v in agg.items()]
    result.sort(key=lambda x: -x[3])
    return result


def parse_pubu(pubu_files):
    """
    Pubu：合併 PUBU.xls + 誠品.xls + PChome.xls
    自動依檔名辨識：
      PUBU → 出版社(2), 商品名稱(3), 您的收入(11)，每列=1冊
      誠品/PChome → 出版者(6), 商品名稱(4), 收入(15)，每列=1冊
    回傳 list of (publisher, '', title, qty, amount)
    """
    agg = {}
    for f in pubu_files:
        fname = getattr(f, 'name', '').lower()
        content = f.getvalue() if hasattr(f, 'getvalue') else f.read()
        wb = xlrd.open_workbook(file_contents=content)

        # 依檔名判斷平台
        is_pubu = 'pubu' in fname and '誠品' not in fname and 'pchome' not in fname

        if is_pubu:
            ws = wb.sheet_by_name('Pubu 訂單報表')
            header_row = -1
            for i in range(ws.nrows):
                if '訂單編號' in str(ws.row_values(i)):
                    header_row = i
                    break
            for i in range(header_row + 1, ws.nrows):
                row = ws.row_values(i)
                if not row[0]:
                    continue
                pub   = str(row[2]).strip()
                title = str(row[3]).strip()
                amt   = float(row[11] or 0)
                key = (pub, title)
                if key not in agg:
                    agg[key] = [0.0, 0.0]
                agg[key][0] += 1
                agg[key][1] += amt
        else:
            # 誠品 or PChome（同格式）
            ws = wb.sheet_by_index(0)
            for i in range(1, ws.nrows):
                row = ws.row_values(i)
                if not row[4]:
                    continue
                pub   = str(row[6]).strip()   # 出版者
                title = str(row[4]).strip()   # 商品名稱
                amt   = float(row[15] or 0)   # 收入
                key = (pub, title)
                if key not in agg:
                    agg[key] = [0.0, 0.0]
                agg[key][0] += 1
                agg[key][1] += amt

    result = [(k[0], '', k[1], v[0], v[1]) for k, v in agg.items()]
    result.sort(key=lambda x: -x[3])
    return result


# ── 建立分頁 ────────────────────────────────────────────────

def _build_sheet(ws, platform, data, genre_catalog):
    """
    依範例格式（Feb 202602）將資料寫入分頁：
      左側：全部書目（排序後）
      右1：Top 20 by 數量
      右2：Top 10 by 金額
    """
    HEADERS = {
        'google':  ([None, '列標籤', 'Primary ISBN', 'Title',
                     '加總 - Qty', '加總 - 扣除美國地區30%稅額'], 6, 3, 10),
        'kobo':    ([None, '列標籤', 'eISBN', '書名',
                     '加總 - 数量', '加總 - 付款幣別基準之付款金額'], 6, 3, 10),
        'readmoo': ([None, '列標籤', '書名',
                     '加總 - 排除免費冊數', '加總 - 總金額'], 5, 1, 7),
        'bokuto':  ([None, '廠商名稱', '商品名稱',
                     '加總 - 排除免費冊數', '加總 - 成本小計'], 5, 2, 8),
        'luanda':  ([None, '列標籤', '書名',
                     '加總 - 銷售數量', '加總 - 應付帳款'], 5, 2, 8),
        'pubu':    ([None, '出版社', '商品名稱',
                     '計數 - 商品名稱', '加總 - 您的收入'], 5, 1, 7),
    }
    left_hdr, left_cols, sep1, r1_col = HEADERS[platform]

    R1_HDR = ['名次', '\xa0\xa0書種\xa0', '作品名(含集數)\xa0\xa0',
              '銷售冊數', '\xa0\xa0出版社', '金額']
    R2_HDR = ['名次', '\xa0\xa0書種\xa0', '作品名(含集數)\xa0\xa0', '銷售冊數']

    r1_end = r1_col + 5
    sep2   = r1_end + 1
    r2_col = sep2 + 1

    # 標題列（第3列）
    col = 1
    for v in left_hdr:
        ws.cell(3, col, v); col += 1
    col += sep1
    for v in R1_HDR:
        ws.cell(3, col, v); col += 1
    col += 1
    for v in R2_HDR:
        ws.cell(3, col, v); col += 1

    # 左側資料
    for idx, (pub, isbn, title, qty, amt) in enumerate(data):
        r = idx + 4
        ws.cell(r, 1, idx + 1)
        ws.cell(r, 2, pub)
        if left_cols == 6:       # google / kobo（有 isbn 欄）
            ws.cell(r, 3, isbn or None)
            ws.cell(r, 4, title)
            ws.cell(r, 5, qty)
            ws.cell(r, 6, round(amt, 4) if amt else 0)
        else:                    # 其他（無 isbn 欄）
            ws.cell(r, 3, title)
            ws.cell(r, 4, qty)
            ws.cell(r, 5, round(amt, 4) if amt else 0)

    # 排行資料（過濾零數量與零金額）
    paid     = [x for x in data if x[3] > 0 and x[4] > 0]
    top_qty  = sorted(paid, key=lambda x: -x[3])[:20]
    top_amt  = sorted(paid, key=lambda x: -x[4])[:10]

    def clean(t):
        return re.sub(r'\s*[\(（]電子書[\)）]', '', t).strip()

    for rank, (pub, isbn, title, qty, amt) in enumerate(top_qty, 1):
        ct = clean(title)
        r  = rank + 3
        ws.cell(r, r1_col,     rank)
        ws.cell(r, r1_col + 1, get_genre(ct, genre_catalog))
        ws.cell(r, r1_col + 2, ct)
        ws.cell(r, r1_col + 3, int(qty))
        ws.cell(r, r1_col + 4, pub)
        ws.cell(r, r1_col + 5, round(amt, 4) if amt else 0)

    for rank, (pub, isbn, title, qty, amt) in enumerate(top_amt, 1):
        ct = clean(title)
        r  = rank + 3
        ws.cell(r, r2_col,     rank)
        ws.cell(r, r2_col + 1, get_genre(ct, genre_catalog))
        ws.cell(r, r2_col + 2, ct)
        ws.cell(r, r2_col + 3, int(qty))
        if platform in ('google', 'bokuto'):
            ws.cell(r, r2_col + 4, round(amt, 4) if amt else 0)
            ql = _col_letter(r2_col + 3)
            al = _col_letter(r2_col + 4)
            ws.cell(r, r2_col + 5, f'={al}{r}/{ql}{r}')


# ── 主要輸出函式 ─────────────────────────────────────────────

def generate_ranking(
    month: str,
    google_file=None,
    kobo_files=None,
    readmoo_file=None,
    bokuto_files=None,
    luanda_file=None,
    pubu_files=None,
    template_file=None,
    genre_catalog: dict = None,
    progress_cb=None,
) -> bytes:
    """
    整合所有平台資料，產生月銷售排行 Excel，回傳 bytes。

    progress_cb(step, total, message): 可選的進度回呼，供 Streamlit 使用。
    """
    def progress(step, total, msg):
        if progress_cb:
            progress_cb(step, total, msg)

    if genre_catalog is None:
        genre_catalog = {}

    kobo_files  = kobo_files  or []
    bokuto_files = bokuto_files or []
    pubu_files  = pubu_files  or []

    STEPS = 8
    progress(0, STEPS, "初始化…")

    # ── 建立工作簿 ──
    if template_file is not None:
        wb = openpyxl.load_workbook(_to_bytesio(template_file))
        progress(1, STEPS, "已載入上月範本")
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)      # 移除預設空白頁
        progress(1, STEPS, "未提供範本，建立空白工作簿")

    CONFIGS = [
        (f'Google_{month}',  'google',  google_file,  parse_google,  [google_file]),
        (f'Kobo_{month}',    'kobo',    kobo_files,   parse_kobo,    [kobo_files]),
        (f'Readmoo_{month}', 'readmoo', readmoo_file, parse_readmoo, [readmoo_file]),
        (f'博客來_{month}',   'bokuto',  bokuto_files, parse_bokuto,  [bokuto_files]),
        (f'亂搭_{month}',     'luanda',  luanda_file,  parse_luanda,  [luanda_file]),
        (f'pubu_{month}',    'pubu',    pubu_files,   parse_pubu,    [pubu_files]),
    ]

    PLATFORM_NAMES = {
        'google': 'Google', 'kobo': 'Kobo', 'readmoo': 'Readmoo',
        'bokuto': '博客來', 'luanda': '亂搭', 'pubu': 'Pubu',
    }

    for step_i, (sheet_name, platform, files, parser, args) in enumerate(CONFIGS, 2):
        pname = PLATFORM_NAMES[platform]
        # 判斷是否有上傳檔案
        has_file = bool(files) if isinstance(files, list) else (files is not None)
        if not has_file:
            progress(step_i, STEPS, f"跳過 {pname}（未上傳）")
            continue

        progress(step_i, STEPS, f"解析 {pname}…")
        try:
            data = parser(*args)
        except Exception as e:
            raise RuntimeError(f"{pname} 解析失敗：{e}") from e

        ws = wb.create_sheet(sheet_name)
        _build_sheet(ws, platform, data, genre_catalog)
        progress(step_i, STEPS, f"{pname} 完成（{len(data)} 筆）")

    progress(STEPS, STEPS, "儲存中…")
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
