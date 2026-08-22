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

# 4.1. Hàm gộp các giá trị duy nhất
def join_unique_values(series):
    values = {
        str(value).strip()
        for value in series
        if pd.notna(value) and str(value).strip()
    }
    return ' | '.join(sorted(values))


# 4.2. Chuẩn bị dữ liệu
df_new_check = df[
    ['item new', 'Description new', 'item code A', 'item ksys']
].copy()

# Làm sạch Item New
df_new_check['item new'] = (
    df_new_check['item new']
    .fillna('')
    .astype(str)
    .str.strip()
)

# Loại Item New và Description New rỗng
df_new_check = df_new_check[
    df_new_check['item new'].ne('')
    & df_new_check['Description new'].notna()
    & df_new_check['Description new'].astype(str).str.strip().ne('')
].copy()


# 4.3. Gộp các Item Code A liên quan đến từng Item New + Description New
df_new_unique = (
    df_new_check
    .groupby(
        ['item new', 'Description new'],
        as_index=False
    )
    .agg({
        'item code A': join_unique_values,
        'item ksys': join_unique_values
    })
)


# 4.4. Tạo nhóm phân loại để giảm số cặp cần so sánh
def get_category(desc):
    """
    Lấy phần đầu mô tả trước dấu phẩy và chuẩn hóa thành category.
    """
    if desc is None or pd.isna(desc):
        return ''

    first_part = str(desc).split(',', 1)[0]

    tokens = tokenize_and_clean(
        first_part,
        remove_class5=True
    )

    return ' '.join(tokens)


df_new_unique['category'] = (
    df_new_unique['Description new']
    .apply(get_category)
)


# 4.5. Union-Find để tạo nhóm Item New trùng
parent = {
    item: item
    for item in df_new_unique['item new'].unique()
}


def find(item):
    while parent[item] != item:
        parent[item] = parent[parent[item]]
        item = parent[item]

    return item


def union(item_a, item_b):
    root_a = find(item_a)
    root_b = find(item_b)

    if root_a != root_b:
        parent[root_b] = root_a


# Khớp tuyệt đối và khớp chứa nhau được xem là tin cậy cao
STRICT_MATCH_STATUSES = {
    'Khớp tuyệt đối',
    'Khớp chứa nhau'
}

pair_records = []


# 4.6. So sánh Description New bằng smart_specs_matching
for category, group in df_new_unique.groupby('category'):

    rows = group[
        [
            'item new',
            'Description new',
            'item code A',
            'item ksys'
        ]
    ].to_dict('records')

    number_of_rows = len(rows)

    for i in range(number_of_rows):
        row_i = rows[i]

        for j in range(i + 1, number_of_rows):
            row_j = rows[j]

            item_i = row_i['item new']
            item_j = row_j['item new']

            # Không so sánh cùng một Item New
            if item_i == item_j:
                continue

            desc_i = row_i['Description new']
            desc_j = row_j['Description new']

            # Tận dụng cùng thuật toán kiểm tra lệch Description.
            # Bỏ CLASS 5 ở cả hai Description New.
            status, detail = smart_specs_matching(
                desc_i,
                desc_j,
                ignore_class5=True
            )

            if status in MATCH_STATUSES:
                pair_records.append({
                    'item new 1': item_i,
                    'Description new 1': desc_i,
                    'item code A 1': row_i['item code A'],
                    'item ksys 1': row_i['item ksys'],

                    'item new 2': item_j,
                    'Description new 2': desc_j,
                    'item code A 2': row_j['item code A'],
                    'item ksys 2': row_j['item ksys'],

                    'Trạng thái': status,
                    'Chi tiết': detail
                })

                # Chỉ gộp nhóm với kết quả tin cậy cao
                if status in STRICT_MATCH_STATUSES:
                    union(item_i, item_j)


# 4.7. Tạo bảng chi tiết các cặp nghi trùng
pair_columns = [
    'item new 1',
    'Description new 1',
    'item code A 1',
    'item ksys 1',
    'item new 2',
    'Description new 2',
    'item code A 2',
    'item ksys 2',
    'Trạng thái',
    'Chi tiết'
]

df_pair_detail = pd.DataFrame(
    pair_records,
    columns=pair_columns
)

# Loại cặp bị ghi trùng
df_pair_detail = df_pair_detail.drop_duplicates()


# Phân chia kết quả tin cậy cao và cần rà thêm
df_pair_detail_strict = df_pair_detail[
    df_pair_detail['Trạng thái'].isin(
        STRICT_MATCH_STATUSES
    )
].copy()

df_pair_detail_loose = df_pair_detail[
    ~df_pair_detail['Trạng thái'].isin(
        STRICT_MATCH_STATUSES
    )
].copy()


# 4.8. Tạo bảng nhóm Item New trùng tin cậy cao
strict_items = set()

if not df_pair_detail_strict.empty:
    strict_items.update(
        df_pair_detail_strict['item new 1'].tolist()
    )
    strict_items.update(
        df_pair_detail_strict['item new 2'].tolist()
    )


# Tổng hợp toàn bộ Code A theo Item New
df_item_information = (
    df_new_check
    .groupby('item new', as_index=False)
    .agg({
        'Description new': join_unique_values,
        'item code A': join_unique_values,
        'item ksys': join_unique_values
    })
)

df_item_information = df_item_information[
    df_item_information['item new'].isin(strict_items)
].copy()


if not df_item_information.empty:
    df_item_information['group_root'] = (
        df_item_information['item new']
        .apply(find)
    )

    # Chỉ giữ nhóm có ít nhất hai Item New
    group_sizes = (
        df_item_information
        .groupby('group_root')['item new']
        .transform('nunique')
    )

    df_duplicate_groups = df_item_information[
        group_sizes >= 2
    ].copy()

    # Đánh số nhóm
    unique_roots = sorted(
        df_duplicate_groups['group_root'].unique()
    )

    group_ids = {
        root: group_number
        for group_number, root in enumerate(
            unique_roots,
            start=1
        )
    }

    df_duplicate_groups['Nhóm nghi trùng số'] = (
        df_duplicate_groups['group_root']
        .map(group_ids)
    )

    df_duplicate_groups = df_duplicate_groups[
        [
            'Nhóm nghi trùng số',
            'item new',
            'Description new',
            'item code A',
            'item ksys'
        ]
    ].sort_values(
        [
            'Nhóm nghi trùng số',
            'item new'
        ]
    )

else:
    df_duplicate_groups = pd.DataFrame(
        columns=[
            'Nhóm nghi trùng số',
            'item new',
            'Description new',
            'item code A',
            'item ksys'
        ]
    )


# 4.9. Ghi kết quả vào Excel
with pd.ExcelWriter(
    output_filename,
    engine='openpyxl',
    mode='a',
    if_sheet_exists='replace'
) as writer:

    df_duplicate_groups.to_excel(
        writer,
        sheet_name='Nhóm Item New trùng',
        index=False
    )

    df_pair_detail_strict.to_excel(
        writer,
        sheet_name='Chi tiết trùng tin cậy',
        index=False
    )

    df_pair_detail_loose.to_excel(
        writer,
        sheet_name='Nghi trùng cần rà',
        index=False
    )


# 4.10. Thống kê
n_groups = (
    df_duplicate_groups['Nhóm nghi trùng số'].nunique()
    if not df_duplicate_groups.empty
    else 0
)

n_items = (
    df_duplicate_groups['item new'].nunique()
    if not df_duplicate_groups.empty
    else 0
)

n_code_a = (
    df_duplicate_groups['item code A'].nunique()
    if not df_duplicate_groups.empty
    else 0)

print(f"    + Item New trùng tin cậy cao: "f"{n_groups} nhóm / {n_items} mã")
print(f"    + Cặp trùng tin cậy cao: "f"{len(df_pair_detail_strict)} cặp")
print(f"    + Cặp nghi trùng cần rà thêm: "f"{len(df_pair_detail_loose)} cặp")
print(f"    + Item Code A liên quan: "f"{n_code_a} nhóm giá trị")