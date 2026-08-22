import pandas as pd
import numpy as np
import re
import unicodedata

# 1. Đọc dữ liệu
file_name = 'Code A check.xlsx'
df = pd.read_excel(file_name)

# 2. Đổi tên cột 
rename_map = {
    'Item': 'item code A',
    'Item Ksys': 'item ksys',
    'Code mới': 'item new',
    'Item Description / Spec': 'Description code A',
    'Description Ksys': 'Description ksys',
    'Spect mới': 'Description new',
    'UOM': 'uom code A',
    'UOM Ksys': 'uom ksys',
    'Unit': 'uom new'
}

# 3. Tối ưu cột
df = df[list(rename_map.keys())].rename(columns=rename_map)

# 4. Làm sạch mã Item (Xóa đuôi .0 và khoảng trắng thừa)
for col in ['item code A', 'item ksys', 'item new']:
    df[col] = (
        df[col]
        .fillna('')
        .astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.strip()
    )

# 5. Kế thừa ô gộp (ffill) cho các cột dữ liệu Code A gốc
ffill_cols = ['item code A', 'Description code A', 'uom code A']
df[ffill_cols] = df[ffill_cols].ffill()

# 6. Loại bỏ các hàng có 'Description new' là Null, N/A, khoảng trắng, bằng 0,
#    hoặc không chứa chữ cái nào (toàn số/toàn ký tự đặc biệt -> chắc chắn là dữ liệu rác/dòng thừa)
df = df.dropna(subset=['Description new']).copy()
desc_new_clean = df['Description new'].astype(str).str.strip().str.lower()
invalid_values = ['n/a', 'na', 'null', 'nan', 'none', '', '0', '0.0']
is_invalid_value = desc_new_clean.isin(invalid_values)
has_no_letter = ~desc_new_clean.str.contains(r'[a-zA-ZÀ-ỹ]', regex=True, na=True)
df = df[~(is_invalid_value | has_no_letter)].copy()

# ==================== PHẦN 1: SO SÁNH UOM ====================

# Từ điển ánh xạ từ viết tắt sang tên chuẩn
uom_mapping = {
    'rol': 'roll',
    'sht': 'sheet',
    'pnl': 'panel'
}

def clean_uom(series):
    return (
        series.fillna('')
        .astype(str)
        .str.replace(r'[\r\n\t\xa0]', '', regex=True)
        .str.strip()
        .str.lower()
        .replace(uom_mapping)
    )

uom_a_clean = clean_uom(df['uom code A'])
uom_ksys_clean = clean_uom(df['uom ksys'])
uom_new_clean = clean_uom(df['uom new'])

df['uom_ksys_match'] = (uom_a_clean == uom_ksys_clean)
df['uom_new_match'] = (uom_a_clean == uom_new_clean)

df['Trạng thái UOM (Code A vs Ksys)'] = df['uom_ksys_match'].map({True: 'Khớp', False: 'Lệch'})
df['Trạng thái UOM (Code A vs New)'] = df['uom_new_match'].map({True: 'Khớp', False: 'Lệch'})


# ==================== PHẦN 2: SO SÁNH DESCRIPTION (SMART SPECS MATCHING - 5 LỚP) ====================
WORD_MAP = {
    # Màu sắc
# Bổ sung mã màu dây điện chuẩn IEC 60757 & các biến thể viết tắt
    'bk': 'black', 'blk': 'black', 'bla': 'black',
    'bn': 'brown', 'brn': 'brown',
    'rd': 'red',
    'og': 'orange', 'org': 'orange', 'orn': 'orange',
    'ye': 'yellow', 'yel': 'yellow', 'yl': 'yellow',
    'gn': 'green', 'grn': 'green',
    'bu': 'blue', 'blu': 'blue', 'bl': 'blue',
    'vt': 'violet', 'pur': 'violet', 'prp': 'violet',
    'gy': 'grey', 'gry': 'grey', 'gray': 'grey',
    'wh': 'white', 'wht': 'white',
    'pk': 'pink', 'pnk': 'pink',
    # Dây tiếp địa 2 màu (Green-Yellow)
    'gnye': 'greenyellow', 'gnyel': 'greenyellow', 'yegn': 'greenyellow',
    # Xuất xứ
    'vn': 'vietnam', 'vnm': 'vietnam', 'cn': 'china', 'chn': 'china',
    'kr': 'korea', 'kor': 'korea', 'jp': 'japan', 'jpn': 'japan',
    'tw': 'taiwan', 'twn': 'taiwan', 'de': 'germany', 'deu': 'germany',
    # thông số kỹ thuật (đơn vị mm² viết theo nhiều kiểu -> quy về 1 dạng chuẩn duy nhất)
    'sqmm': 'sqmm', 'mmsq': 'sqmm', 'sq': 'sqmm','fr': 'flameretardant',
    'fr': 'flameretardant','hpvc': 'hpvc','vina': 'vietnam'
}

COLOR_SET = {'black', 'red', 'blue', 'yellow', 'white', 'green', 'orange', 'grey', 'brown'}
ORIGIN_SET = {'vietnam', 'china', 'korea', 'japan', 'taiwan', 'germany'}

# ==================== HÀM BỔ TRỢ CHUẨN HÓA DỮ LIỆU ====================

# ==================== HÀM TIỀN XỬ LÝ & BÓC TÁCH DỮ LIỆU ====================

def remove_diacritics(text):
    """Xóa dấu tiếng Việt"""
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize('NFD', text)
    return re.sub(r'[\u0300-\u036f]', '', text).replace('đ', 'd').replace('Đ', 'D')

def normalize_numbers(text):
    """Chuẩn hóa định dạng số thập phân và quy đổi số nguyên"""
    # LƯU Ý: Đã khảo sát toàn bộ dữ liệu thực tế và xác nhận dấu phẩy trong mô tả
    # LUÔN được dùng làm dấu phân cách giữa các trường thông số (vd '630,1250A' = 2 số),
    # KHÔNG dùng làm dấu thập phân (số thập phân thật luôn viết bằng dấu chấm: '0.6KV', '1.5CL').
    # Vì vậy KHÔNG quy đổi dấu phẩy -> dấu chấm nữa để tránh ghép nhầm 2 số liệt kê thành 1 số thập phân.
    text = re.sub(r'(\d+)\.0+(?!\d)', r'\1', text)          # 6.0 -> 6
    text = re.sub(r'(\d+\.\d*?[1-9])0+(?!\d)', r'\1', text) # 1.50 -> 1.5
    return text

def split_digit_letter_runs(token):
    """
    Tách một token dính liền chữ+số thành các token nhỏ hơn theo ranh giới số/chữ,
    giữ nguyên số thập phân (vd 31.5 không bị tách rời).
    Đây là chìa khóa xử lý 2 nhóm lỗi:
      - Số dính liền đơn vị do thiếu khoảng trắng nguồn: '16mmsq' -> ['16','mmsq']
      - Số dính liền tên khác (thiếu dấu phẩy nguồn): '64hengzhu' -> ['64','hengzhu']
      - Số/chữ bị đảo thứ tự giữa 2 mô tả: '110vdc' -> ['110','vdc'], 'dc110v' -> ['dc','110','v']
        (sau khi tách, số '110' ở cả 2 bên đều là token riêng độc lập -> so khớp được dù thứ tự đảo)
    """
    return re.findall(r'\d+\.\d+|\d+|[a-z]+', token)

def tokenize_and_clean(text, remove_class5=False):
    """
    Tách từ & làm sạch: coi ký tự đặc biệt và khoảng trắng là dấu phân cách.
    Chỉ trích xuất chữ và số, giữ lại số thập phân.
    """
    if text is None or pd.isna(text):
        return []

    s = remove_diacritics(str(text).lower())
    s = re.sub(r'\bflame[\s-]+retardant\b', 'fr', s)
    s = re.sub(r'\bh[\s-]*pvc\b', 'hpvc', s)
    s = re.sub(r'\bcosmolink\s+vina\b', 'cosmolink', s)
    s = normalize_numbers(s)

    # CLASS 5 thừa ở Code New không được đem đi so sánh
    if remove_class5:
        s = re.sub(
            r'\bclass\s*[-:=]?\s*5(?:\.0)?\b',
            ' ',
            s
        )

    # W40 -> 40, H60 -> 60, L2000 -> 2000, T2 -> 2, DIA10 -> 10
    s = re.sub(r'\b([whldt]|phi|dia)(\d+)', r'\2', s)

    raw_tokens = re.findall(
        r'[a-z0-9]+(?:\.[a-z0-9]+)*',
        s
    )

    tokens = []

    for token in raw_tokens:
        token = token.strip('.')

        if not token:
            continue

        for piece in split_digit_letter_runs(token):
            tokens.append(WORD_MAP.get(piece, piece))

    return tokens

def tokenize_no_parenthetical(text, remove_class5=False):
    if text is None or pd.isna(text):
        return []

    s = re.sub(r'\([^)]*\)', ' ', str(text))

    return tokenize_and_clean(
        s,
        remove_class5=remove_class5
    )
# ==================== PHẦN 2: SMART SPECS MATCHING (5 LỚP TỐI ƯU) ====================

def smart_specs_matching(desc1, desc2,ignore_class5=False, threshold=0.8):
    t1 = tokenize_and_clean(desc1, remove_class5=ignore_class5)
    t2 = tokenize_and_clean(desc2, remove_class5=ignore_class5)

    # -------------------------------------------------------------
    # LỚP 0: Lọc rỗng & Khớp tuyệt đối
    # -------------------------------------------------------------
    if not t1 and not t2:
        return "Khớp tuyệt đối", "Cả hai mô tả đều rỗng"
    if not t1 or not t2:
        return "Không khớp", "Một trong hai mô tả bị thiếu (rỗng)"
    if t1 == t2:
        return "Khớp tuyệt đối", "Hai chuỗi hoàn toàn giống nhau"

    # -------------------------------------------------------------
    # LỚP 1: Kiểm tra Màu sắc & Xuất xứ (Chỉ báo lỗi khi CẢ 2 BÊN CÙNG CÓ nhưng KHÁC NHAU)
    # Bỏ qua nội dung trong ngoặc () vì đó thường là chú giải viết tắt, không phải màu/xuất xứ thật khác biệt.
    # -------------------------------------------------------------
    t1_np, t2_np = tokenize_no_parenthetical(desc1, remove_class5=ignore_class5), tokenize_no_parenthetical(desc2, remove_class5=ignore_class5)

    c1, c2 = {w for w in t1_np if w in COLOR_SET}, {w for w in t2_np if w in COLOR_SET}
    if c1 and c2 and c1 != c2:
        return "Lỗi lệch màu/xuất xứ", f"Lệch màu sắc: {c1} vs {c2}"

    o1, o2 = {w for w in t1_np if w in ORIGIN_SET}, {w for w in t2_np if w in ORIGIN_SET}
    if o1 and o2 and o1 != o2:
        return "Lỗi lệch màu/xuất xứ", f"Lệch xuất xứ: {o1} vs {o2}"

    # Phân định chuỗi Dài / Ngắn
    if len(t1) >= len(t2):
        t_long, t_short = t1, t2
    else:
        t_long, t_short = t2, t1

    # -------------------------------------------------------------
    # LỚP 2: Kiểm tra Con số / Thông số kỹ thuật
    # -------------------------------------------------------------
    specs_short = [w for w in t_short if re.search(r'\d', w)]
    missing_specs = []
    
    for s in specs_short:
        # Khớp từ độc lập HOẶC nằm trong một mã Model gộp (vd: '1215' nằm trong 'ysr1215gw')
        if (s in t_long) or (len(s) >= 3 and any(s in token for token in t_long)):
            continue
        missing_specs.append(s)

    if missing_specs:
        return "Lỗi lệch con số/kích thước", f"Sai lệch thông số: {', '.join(missing_specs)}"

    # -------------------------------------------------------------
    # LỚP 3: Kiểm tra Khớp chứa nhau (Bản tóm tắt)
    # -------------------------------------------------------------
    matched_words = sum(
        1 for w in t_short 
        if (w in t_long) or (len(w) >= 3 and any(w in token for token in t_long))
    )

    if matched_words == len(t_short):
        return "Khớp chứa nhau", "Chuỗi ngắn là bản tóm tắt nằm trọn trong chuỗi dài"

    # -------------------------------------------------------------
    # LỚP 4: Độ phủ từ còn lại (Word Coverage Ratio)
    # -------------------------------------------------------------
    coverage = matched_words / len(t_short) if t_short else 0.0
    if coverage >= threshold:
        return "Khớp tương đối", f"Độ phủ từ đạt {coverage*100:.1f}%"
    
    return "Không khớp", f"Độ phủ từ chỉ đạt {coverage*100:.1f}% (< {int(threshold*100)}%)"


# ==================== PHẦN 3: THI CÔNG SO SÁNH & XUẤT BÁO CÁO ====================

def run_matching(df, col_a, col_target, prefix, ignore_class5=False):
    """Hàm gộp xử lý matching theo cặp cột"""
    res = df.apply(lambda row: smart_specs_matching(
        row[col_a], 
        row[col_target], 
        ignore_class5=ignore_class5
    ), axis=1)
    df[f'Trạng thái Desc ({prefix})'] = [r[0] for r in res]
    df[f'Chi tiết Desc ({prefix})'] = [r[1] for r in res]

# Thực thi kiểm tra cho cả Ksys và New
run_matching(df, 'Description code A', 'Description ksys', 'Code A vs Ksys', ignore_class5=False)
run_matching(df, 'Description code A', 'Description new', 'Code A vs New', ignore_class5=True)

# Tập hợp các trạng thái được coi là hợp lệ (Không tính là lệch)
MATCH_STATUSES = {'Khớp tuyệt đối', 'Khớp chứa nhau', 'Khớp tương đối'}

# Lọc danh sách lệch UOM
df_uom_diff_ksys = df[~df['uom_ksys_match']].copy()
df_uom_diff_new = df[~df['uom_new_match']].copy()

# Lọc danh sách lệch Description
df_desc_diff_ksys = df[~df['Trạng thái Desc (Code A vs Ksys)'].isin(MATCH_STATUSES)].copy()
# Lọc lệch Description giữa Code A và New
df_desc_diff_new = df[~df['Trạng thái Desc (Code A vs New)'].isin(MATCH_STATUSES)].copy()

# Lưu số dòng trước khi loại duplicate
desc_diff_new_before_dedup = len(df_desc_diff_new)

# Loại hàng trùng hoàn toàn trên tất cả các cột
df_desc_diff_new = (
    df_desc_diff_new
    .drop_duplicates(keep='first')
    .reset_index(drop=True)
)

removed_desc_diff_new_duplicates = (
    desc_diff_new_before_dedup - len(df_desc_diff_new)
)
# Xuất dữ liệu đa sheet ra Excel
output_filename = 'Check uom & description.xlsx'

with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    df_desc_diff_ksys.to_excel(writer, sheet_name='Lệch Desc (Code A vs Ksys)', index=False)
    df_desc_diff_new.to_excel(writer, sheet_name='Lệch Desc (Code A vs New)', index=False)
    df_uom_diff_ksys.to_excel(writer, sheet_name='Lệch UOM (Code A vs Ksys)', index=False)
    df_uom_diff_new.to_excel(writer, sheet_name='Lệch UOM (Code A vs New)', index=False)
    df.to_excel(writer, sheet_name='Toàn bộ Data Checked', index=False)

print(f"--> Xuất thành công file gộp: '{output_filename}'")
print(f"    + Lệch UOM  : Ksys ({len(df_uom_diff_ksys)}) | New ({len(df_uom_diff_new)})")
print(f"    + Lệch Desc : Ksys ({len(df_desc_diff_ksys)}) | New ({len(df_desc_diff_new)})")

# ==================== PHẦN 4: TÌM ITEM NEW NGHI TRÙNG MÔ TẢ ====================
# Mục tiêu: trong hệ mã New, tìm những mã item KHÁC NHAU nhưng mô tả (sau chuẩn hóa)
# lại khớp nhau theo đúng thuật toán smart_specs_matching ở trên -> nghi bị tạo trùng item.

# =============================================================
# 4.1. HÀM HỖ TRỢ
# =============================================================

def join_unique_values(series):
    """Gộp các giá trị duy nhất và bỏ giá trị rỗng."""
    values = {
        str(value).strip()
        for value in series
        if pd.notna(value) and str(value).strip()
    }

    return ' | '.join(sorted(values))


def normalize_description_part4(description):
    """Chuẩn hóa Description New riêng cho Phần 4."""
    if description is None or pd.isna(description):
        return ''

    s = remove_diacritics(str(description).lower())

    # Giữ cùng quy tắc chuẩn hóa với code hiện tại
    s = re.sub(r'\bflame[\s-]+retardant\b', 'fr', s)
    s = re.sub(r'\bh[\s-]*pvc\b', 'hpvc', s)
    s = re.sub(r'\bcosmolink\s+vina\b', 'cosmolink', s)

    # Bỏ CLASS 5
    s = re.sub(
        r'\bclass\s*[-:=]?\s*5(?:\.0)?\b',
        ' ',
        s
    )

    s = normalize_numbers(s)
    s = re.sub(r'\s+', ' ', s).strip()

    return s


def make_description_key(description):
    """
    Khóa toàn bộ nội dung Description New.

    Không phụ thuộc thứ tự và không tính token lặp lại.
    Thiếu hoặc khác một token sẽ tạo khóa khác.
    """
    tokens = tokenize_and_clean(
        description,
        remove_class5=True
    )

    if not tokens:
        return ''

    return '|'.join(sorted(set(tokens)))


def make_technical_key(description):
    """
    Khóa thông số kỹ thuật có giữ quan hệ chữ và số.

    Ví dụ:
    - M12 khác M10.
    - L110 khác L30.
    - M12,L30 khác M30,L12.
    - 100VA khác 200VA.
    """
    s = normalize_description_part4(description)

    if not s:
        return ''

    raw_tokens = re.findall(
        r'[a-z0-9]+(?:\.[a-z0-9]+)*',
        s
    )

    technical_tokens = {
        token.strip('.')
        for token in raw_tokens
        if re.search(r'\d', token)
    }

    return '|'.join(sorted(technical_tokens))


def make_category_key(description):
    """Lấy tên/loại vật tư ở phần đầu trước dấu phẩy."""
    if description is None or pd.isna(description):
        return ''

    first_part = str(description).split(',', 1)[0]

    tokens = tokenize_and_clean(
        first_part,
        remove_class5=True
    )

    return '|'.join(tokens)


def make_color_key(description):
    """Lấy toàn bộ màu sắc, kể cả màu nằm một phía."""
    strict_color_set = set(COLOR_SET) | {
        'violet',
        'pink',
        'greenyellow'
    }

    tokens = tokenize_no_parenthetical(
        description,
        remove_class5=True
    )

    colors = {
        token
        for token in tokens
        if token in strict_color_set
    }

    return '|'.join(sorted(colors))


def make_origin_key(description):
    """Lấy toàn bộ thông tin xuất xứ."""
    tokens = tokenize_no_parenthetical(
        description,
        remove_class5=True
    )

    origins = {
        token
        for token in tokens
        if token in ORIGIN_SET
    }

    return '|'.join(sorted(origins))


# =============================================================
# 4.2. CHUẨN BỊ DỮ LIỆU
# =============================================================

df_new_check = df[
    [
        'item new',
        'Description new',
        'item code A',
        'Description code A',
        'item ksys'
    ]
].copy()

# Làm sạch các cột mã
for col in ['item new', 'item code A', 'item ksys']:
    df_new_check[col] = (
        df_new_check[col]
        .fillna('')
        .astype(str)
        .str.strip()
    )

# Loại Item New và Description New rỗng
df_new_check = df_new_check[
    df_new_check['item new'].ne('')
    & df_new_check['Description new'].notna()
    & df_new_check[
        'Description new'
    ].astype(str).str.strip().ne('')
].copy()

# Loại các dòng trùng hoàn toàn
df_new_check = (
    df_new_check
    .drop_duplicates(keep='first')
    .reset_index(drop=True)
)


# =============================================================
# 4.3. TỔNG HỢP THEO ITEM NEW VÀ DESCRIPTION NEW
# =============================================================

df_new_unique = (
    df_new_check
    .groupby(
        ['item new', 'Description new'],
        as_index=False
    )
    .agg({
        'item code A': join_unique_values,
        'Description code A': join_unique_values,
        'item ksys': join_unique_values
    })
)

df_new_unique['Description key'] = (
    df_new_unique['Description new']
    .apply(make_description_key)
)

df_new_unique['Technical key'] = (
    df_new_unique['Description new']
    .apply(make_technical_key)
)

df_new_unique['Category key'] = (
    df_new_unique['Description new']
    .apply(make_category_key)
)

df_new_unique['Color key'] = (
    df_new_unique['Description new']
    .apply(make_color_key)
)

df_new_unique['Origin key'] = (
    df_new_unique['Description new']
    .apply(make_origin_key)
)

# Khóa tổng hợp nghiêm ngặt
df_new_unique['Duplicate key'] = (
    df_new_unique['Category key']
    + '||'
    + df_new_unique['Description key']
    + '||'
    + df_new_unique['Technical key']
    + '||'
    + df_new_unique['Color key']
    + '||'
    + df_new_unique['Origin key']
)

df_new_unique = df_new_unique[
    df_new_unique['Description key'].ne('')
].copy()


# =============================================================
# 4.4. TÌM NHIỀU ITEM NEW CÓ CÙNG DESCRIPTION NEW
# =============================================================

item_count_by_key = (
    df_new_unique
    .groupby('Duplicate key')['item new']
    .nunique()
)

duplicate_keys = item_count_by_key[
    item_count_by_key > 1
].index.tolist()

STRICT_DUPLICATE_STATUSES = {
    'Khớp tuyệt đối',
    'Khớp chứa nhau'
}

pair_records = []
confirmed_keys = set()


# =============================================================
# 4.5. XÁC NHẬN BẰNG SMART_SPECS_MATCHING
# =============================================================

for duplicate_key in duplicate_keys:

    group = df_new_unique[
        df_new_unique['Duplicate key'] == duplicate_key
    ]

    rows = group.to_dict('records')

    for i in range(len(rows)):
        row_1 = rows[i]

        for j in range(i + 1, len(rows)):
            row_2 = rows[j]

            if row_1['item new'] == row_2['item new']:
                continue

            desc_1 = row_1['Description new']
            desc_2 = row_2['Description new']

            status, detail = smart_specs_matching(
                desc_1,
                desc_2,
                ignore_class5=True
            )

            # Kiểm tra nghiêm ngặt lần cuối
            same_description = (
                row_1['Description key']
                == row_2['Description key']
            )

            same_technical_specs = (
                row_1['Technical key']
                == row_2['Technical key']
            )

            same_category = (
                row_1['Category key']
                == row_2['Category key']
            )

            same_color = (
                row_1['Color key']
                == row_2['Color key']
            )

            same_origin = (
                row_1['Origin key']
                == row_2['Origin key']
            )

            is_duplicate = (
                same_description
                and same_technical_specs
                and same_category
                and same_color
                and same_origin
                and status in STRICT_DUPLICATE_STATUSES
            )

            if not is_duplicate:
                continue

            confirmed_keys.add(duplicate_key)

            pair_records.append({
                'Duplicate key': duplicate_key,

                'item new 1': row_1['item new'],
                'Description new 1': desc_1,
                'item code A 1': row_1['item code A'],
                'Description code A 1':
                    row_1['Description code A'],
                'item ksys 1': row_1['item ksys'],

                'item new 2': row_2['item new'],
                'Description new 2': desc_2,
                'item code A 2': row_2['item code A'],
                'Description code A 2':
                    row_2['Description code A'],
                'item ksys 2': row_2['item ksys'],

                'Thông số kỹ thuật':
                    row_1['Technical key'],
                'Màu sắc':
                    row_1['Color key'],
                'Xuất xứ':
                    row_1['Origin key'],

                'Trạng thái': status,
                'Chi tiết': detail
            })


# =============================================================
# 4.6. TẠO BẢNG NHÓM ITEM NEW TRÙNG DESCRIPTION
# =============================================================

group_ids = {
    key: number
    for number, key in enumerate(
        sorted(confirmed_keys),
        start=1
    )
}

df_duplicate_groups = df_new_unique[
    df_new_unique['Duplicate key'].isin(confirmed_keys)
].copy()

if not df_duplicate_groups.empty:

    df_duplicate_groups['Nhóm trùng số'] = (
        df_duplicate_groups['Duplicate key']
        .map(group_ids)
    )

    df_duplicate_groups = df_duplicate_groups[
        [
            'Nhóm trùng số',
            'item new',
            'Description new',
            'item code A',
            'Description code A',
            'item ksys',
            'Technical key',
            'Color key',
            'Origin key'
        ]
    ].rename(
        columns={
            'Technical key': 'Thông số kỹ thuật',
            'Color key': 'Màu sắc',
            'Origin key': 'Xuất xứ'
        }
    ).sort_values(
        ['Nhóm trùng số', 'item new']
    ).reset_index(drop=True)

else:
    df_duplicate_groups = pd.DataFrame(
        columns=[
            'Nhóm trùng số',
            'item new',
            'Description new',
            'item code A',
            'Description code A',
            'item ksys',
            'Thông số kỹ thuật',
            'Màu sắc',
            'Xuất xứ'
        ]
    )


# =============================================================
# 4.7. TẠO BẢNG CHI TIẾT CÁC CẶP TRÙNG
# =============================================================

for record in pair_records:
    record['Nhóm trùng số'] = group_ids[
        record['Duplicate key']
    ]

    del record['Duplicate key']


pair_columns = [
    'Nhóm trùng số',

    'item new 1',
    'Description new 1',
    'item code A 1',
    'Description code A 1',
    'item ksys 1',

    'item new 2',
    'Description new 2',
    'item code A 2',
    'Description code A 2',
    'item ksys 2',

    'Thông số kỹ thuật',
    'Màu sắc',
    'Xuất xứ',
    'Trạng thái',
    'Chi tiết'
]

df_duplicate_pairs = (
    pd.DataFrame(
        pair_records,
        columns=pair_columns
    )
    .drop_duplicates(keep='first')
    .reset_index(drop=True)
)

if not df_duplicate_pairs.empty:
    df_duplicate_pairs = df_duplicate_pairs.sort_values(
        [
            'Nhóm trùng số',
            'item new 1',
            'item new 2'
        ]
    ).reset_index(drop=True)


# =============================================================
# 4.8. TÌM ITEM NEW CÓ NHIỀU ITEM CODE A KHÁC NHAU
# =============================================================

df_new_code_a_relation = (
    df_new_check[
        df_new_check['item code A'].ne('')
    ]
    .drop_duplicates(
        subset=['item new', 'item code A'],
        keep='first'
    )
)

code_a_count_by_item_new = (
    df_new_code_a_relation
    .groupby('item new')['item code A']
    .nunique()
)

multiple_code_a_items = code_a_count_by_item_new[
    code_a_count_by_item_new > 1
].index.tolist()


# Chi tiết Item New có nhiều Item Code A
df_multiple_code_a_detail = (
    df_new_check[
        df_new_check['item new'].isin(
            multiple_code_a_items
        )
        & df_new_check['item code A'].ne('')
    ]
    .drop_duplicates(keep='first')
    [
        [
            'item new',
            'Description new',
            'item code A',
            'Description code A',
            'item ksys'
        ]
    ]
    .sort_values(
        ['item new', 'item code A']
    )
    .reset_index(drop=True)
)


# Tổng hợp Item New có nhiều Item Code A
if not df_multiple_code_a_detail.empty:

    df_multiple_code_a_summary = (
        df_multiple_code_a_detail
        .groupby('item new', as_index=False)
        .agg(
            so_luong_code_a=(
                'item code A',
                'nunique'
            ),
            danh_sach_code_a=(
                'item code A',
                join_unique_values
            ),
            danh_sach_description_new=(
                'Description new',
                join_unique_values
            ),
            danh_sach_item_ksys=(
                'item ksys',
                join_unique_values
            )
        )
        .rename(
            columns={
                'so_luong_code_a':
                    'Số lượng Item Code A',
                'danh_sach_code_a':
                    'Danh sách Item Code A',
                'danh_sach_description_new':
                    'Danh sách Description New',
                'danh_sach_item_ksys':
                    'Danh sách Item Ksys'
            }
        )
        .sort_values(
            ['Số lượng Item Code A', 'item new'],
            ascending=[False, True]
        )
        .reset_index(drop=True)
    )

else:
    df_multiple_code_a_summary = pd.DataFrame(
        columns=[
            'item new',
            'Số lượng Item Code A',
            'Danh sách Item Code A',
            'Danh sách Description New',
            'Danh sách Item Ksys'
        ]
    )


# =============================================================
# 4.9. GHI KẾT QUẢ VÀO FILE EXCEL
# =============================================================

with pd.ExcelWriter(
    output_filename,
    engine='openpyxl',
    mode='a',
    if_sheet_exists='replace'
) as writer:

    df_duplicate_groups.to_excel(
        writer,
        sheet_name='Item New trùng Description',
        index=False
    )

    df_duplicate_pairs.to_excel(
        writer,
        sheet_name='Chi tiết Item New trùng',
        index=False
    )

    df_multiple_code_a_summary.to_excel(
        writer,
        sheet_name='Item New nhiều Code A',
        index=False
    )

    df_multiple_code_a_detail.to_excel(
        writer,
        sheet_name='Chi tiết New nhiều Code A',
        index=False
    )


# =============================================================
# 4.10. THỐNG KÊ
# =============================================================

n_duplicate_groups = (
    df_duplicate_groups['Nhóm trùng số'].nunique()
    if not df_duplicate_groups.empty
    else 0
)

n_duplicate_items = (
    df_duplicate_groups['item new'].nunique()
    if not df_duplicate_groups.empty
    else 0
)

n_multiple_code_a = len(multiple_code_a_items)

print(
    f"    + Item New trùng Description: "
    f"{n_duplicate_groups} nhóm / {n_duplicate_items} mã"
)

print(
    f"    + Cặp Item New trùng Description: "
    f"{len(df_duplicate_pairs)} cặp"
)

print(
    f"    + Item New có nhiều Item Code A: "
    f"{n_multiple_code_a} mã"
)