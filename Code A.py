# -*- coding: utf-8 -*-
"""
=============================================================================
 KIỂM TRA ĐỐI CHIẾU DANH MỤC VẬT TƯ: CODE A  <->  KSYS  <->  CODE NEW
=============================================================================
 Phần 1 : So sánh UOM
 Phần 2 : So sánh Description (Smart Specs Matching - 5 lớp)
 Phần 3 : Xuất báo cáo lệch UOM / Description
 Phần 4 : Kiểm tra trùng lặp trong hệ mã New  
 Phần 5 : Cảnh báo chất lượng dữ liệu         
=============================================================================
"""

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field as dc_field
import pandas as pd


# =============================================================================
# 0. CẤU HÌNH CHUNG
# =============================================================================

INPUT_FILE = 'Code A check.xlsx'
OUTPUT_FILE = 'Check uom & description.xlsx'

COLUMN_RENAME_MAP = {
    'Item': 'item code A',
    'Item Ksys': 'item ksys',
    'Code mới': 'item new',
    'Item Description / Spec': 'Description code A',
    'Description Ksys': 'Description ksys',
    'Spect mới': 'Description new',
    'UOM': 'uom code A',
    'UOM Ksys': 'uom ksys',
    'Unit': 'uom new',
}

CODE_COLUMNS = ['item code A', 'item ksys', 'item new']
FFILL_COLUMNS = ['item code A', 'Description code A', 'uom code A']

INVALID_DESCRIPTION_VALUES = [
    'n/a', 'na', 'null', 'nan', 'none', '', '0', '0.0',
]


# =============================================================================
# 1. ĐỌC & LÀM SẠCH DỮ LIỆU
# =============================================================================

def load_dataset(file_name=INPUT_FILE):
    """Đọc file nguồn, đổi tên cột, làm sạch mã và loại dòng rác."""
    raw = pd.read_excel(file_name)

    df = raw[list(COLUMN_RENAME_MAP.keys())].rename(columns=COLUMN_RENAME_MAP)

    # Làm sạch mã Item (xóa đuôi '.0' do Excel ép kiểu số và khoảng trắng thừa)
    for column in CODE_COLUMNS:
        df[column] = (
            df[column]
            .fillna('')
            .astype(str)
            .str.replace(r'\.0$', '', regex=True)
            .str.strip()
        )

    # Kế thừa ô gộp (merged cell) cho khối dữ liệu Code A gốc
    df[FFILL_COLUMNS] = df[FFILL_COLUMNS].ffill()

    return df


def drop_invalid_description_new(df):
    """
    Loại dòng có 'Description new' rỗng / N/A / bằng 0 / không chứa chữ cái.
    Trả về (df_hợp_lệ, df_bị_loại) để Phần 5 còn truy vết được.
    """
    has_value = df['Description new'].notna()

    cleaned = df['Description new'].astype(str).str.strip().str.lower()
    is_invalid_value = cleaned.isin(INVALID_DESCRIPTION_VALUES)
    has_no_letter = ~cleaned.str.contains(r'[a-zA-ZÀ-ỹ]', regex=True, na=True)

    keep_mask = has_value & ~(is_invalid_value | has_no_letter)

    return df[keep_mask].copy(), df[~keep_mask].copy()


# =============================================================================
# PHẦN 1: SO SÁNH UOM
# =============================================================================

UOM_MAPPING = {
    'rol': 'roll',
    'sht': 'sheet',
    'pnl': 'panel',
}


def clean_uom(series):
    return (
        series.fillna('')
        .astype(str)
        .str.replace(r'[\r\n\t\xa0]', '', regex=True)
        .str.strip()
        .str.lower()
        .replace(UOM_MAPPING)
    )


def compare_uom(df):
    uom_a = clean_uom(df['uom code A'])

    df['uom_ksys_match'] = uom_a == clean_uom(df['uom ksys'])
    df['uom_new_match'] = uom_a == clean_uom(df['uom new'])

    df['Trạng thái UOM (Code A vs Ksys)'] = df['uom_ksys_match'].map(
        {True: 'Khớp', False: 'Lệch'}
    )
    df['Trạng thái UOM (Code A vs New)'] = df['uom_new_match'].map(
        {True: 'Khớp', False: 'Lệch'}
    )

    return df


# =============================================================================
# PHẦN 2: SO SÁNH DESCRIPTION (SMART SPECS MATCHING - 5 LỚP)
# =============================================================================

WORD_MAP = {
    # --- Màu sắc: mã màu dây điện IEC 60757 & các biến thể viết tắt ---
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
    # Dây tiếp địa 2 màu
    'gnye': 'greenyellow', 'gnyel': 'greenyellow', 'yegn': 'greenyellow',
    # --- Xuất xứ ---
    'vn': 'vietnam', 'vnm': 'vietnam', 'cn': 'china', 'chn': 'china',
    'kr': 'korea', 'kor': 'korea', 'jp': 'japan', 'jpn': 'japan',
    'tw': 'taiwan', 'twn': 'taiwan', 'de': 'germany', 'deu': 'germany',
    'vina': 'vietnam',
    # --- Thông số kỹ thuật (mm² viết nhiều kiểu -> quy về 1 dạng chuẩn) ---
    'sqmm': 'sqmm', 'mmsq': 'sqmm', 'sq': 'sqmm',
    'fr': 'flameretardant',
    'hpvc': 'hpvc',
}

COLOR_SET = {
    'black', 'red', 'blue', 'yellow', 'white',
    'green', 'orange', 'grey', 'brown',
}

ORIGIN_SET = {'vietnam', 'china', 'korea', 'japan', 'taiwan', 'germany'}

MATCH_STATUSES = {'Khớp tuyệt đối', 'Khớp chứa nhau', 'Khớp tương đối'}


# ---------------------------------------------------------------------------
# 2.1. Hàm tiền xử lý dùng chung
# ---------------------------------------------------------------------------

def remove_diacritics(text):
    """Xóa dấu tiếng Việt."""
    if not isinstance(text, str):
        return ''

    text = unicodedata.normalize('NFD', text)

    return (
        re.sub(r'[\u0300-\u036f]', '', text)
        .replace('đ', 'd')
        .replace('Đ', 'D')
    )


def normalize_numbers(text):
    """
    Chuẩn hóa số thập phân: 6.0 -> 6, 1.50 -> 1.5.

    LƯU Ý (đã khảo sát toàn bộ dữ liệu thực tế): dấu phẩy trong mô tả LUÔN
    là dấu phân cách trường ('630,1250A' = 2 số), KHÔNG phải dấu thập phân
    (số thập phân thật luôn viết bằng dấu chấm: '0.6KV', '1.5CL'). Vì vậy
    tuyệt đối KHÔNG quy đổi ',' -> '.'.
    """
    text = re.sub(r'(\d+)\.0+(?!\d)', r'\1', text)
    text = re.sub(r'(\d+\.\d*?[1-9])0+(?!\d)', r'\1', text)

    return text

def split_digit_letter_runs(token):
    """
    Tách token dính liền chữ+số theo ranh giới số/chữ, giữ nguyên số thập phân.
        '16mmsq'    -> ['16', 'mmsq']       (thiếu khoảng trắng nguồn)
        '64hengzhu' -> ['64', 'hengzhu']    (thiếu dấu phẩy nguồn)
        '110vdc' / 'dc110v' -> số '110' tách riêng ở cả hai bên nên vẫn khớp
                               được dù thứ tự chữ/số bị đảo.
    """
    return re.findall(r'\d+\.\d+|\d+|[a-z]+', token)

def tokenize_and_clean(text, remove_class5=False):
    """Tách từ & làm sạch: ký tự đặc biệt và khoảng trắng đều là dấu phân cách."""
    if text is None or pd.isna(text):
        return []

    s = remove_diacritics(str(text).lower())
    s = re.sub(r'\bflame[\s-]+retardant\b', 'fr', s)
    s = re.sub(r'\bh[\s-]*pvc\b', 'hpvc', s)
    s = re.sub(r'\bcosmolink\s+vina\b', 'cosmolink', s)
    s = normalize_numbers(s)

    # 'CLASS 5' thừa ở Code New không đem đi so sánh
    if remove_class5:
        s = re.sub(r'\bclass\s*[-:=]?\s*5(?:\.0)?\b', ' ', s)

    # W40 -> 40, H60 -> 60, L2000 -> 2000, T2 -> 2, DIA10 -> 10
    s = re.sub(r'\b([whldt]|phi|dia)(\d+)', r'\2', s)
    tokens = []

    for raw_token in re.findall(r'[a-z0-9]+(?:\.[a-z0-9]+)*', s):
        raw_token = raw_token.strip('.')

        if not raw_token:
            continue

        for piece in split_digit_letter_runs(raw_token):
            tokens.append(WORD_MAP.get(piece, piece))

    return tokens


def tokenize_no_parenthetical(text, remove_class5=False):
    """Như `tokenize_and_clean` nhưng bỏ nội dung trong ngoặc đơn."""
    if text is None or pd.isna(text):
        return []
    return tokenize_and_clean(
        re.sub(r'\([^)]*\)', ' ', str(text)),
        remove_class5=remove_class5,
    )


# ---------------------------------------------------------------------------
# 2.2. Smart Specs Matching
# ---------------------------------------------------------------------------

def smart_specs_matching(desc1, desc2, ignore_class5=False, threshold=0.8):
    t1 = tokenize_and_clean(desc1, remove_class5=ignore_class5)
    t2 = tokenize_and_clean(desc2, remove_class5=ignore_class5)

    # --- LỚP 0: Lọc rỗng & khớp tuyệt đối ---
    if not t1 and not t2:
        return 'Khớp tuyệt đối', 'Cả hai mô tả đều rỗng'
    if not t1 or not t2:
        return 'Không khớp', 'Một trong hai mô tả bị thiếu (rỗng)'
    if t1 == t2:
        return 'Khớp tuyệt đối', 'Hai chuỗi hoàn toàn giống nhau'

    # --- LỚP 1: Màu sắc & xuất xứ (chỉ báo lỗi khi CẢ HAI cùng có mà khác) ---
    # Bỏ qua nội dung trong ngoặc vì đó thường là chú giải viết tắt.
    t1_np = tokenize_no_parenthetical(desc1, remove_class5=ignore_class5)
    t2_np = tokenize_no_parenthetical(desc2, remove_class5=ignore_class5)

    c1 = {w for w in t1_np if w in COLOR_SET}
    c2 = {w for w in t2_np if w in COLOR_SET}
    if c1 and c2 and c1 != c2:
        return 'Lỗi lệch màu/xuất xứ', f'Lệch màu sắc: {c1} vs {c2}'

    o1 = {w for w in t1_np if w in ORIGIN_SET}
    o2 = {w for w in t2_np if w in ORIGIN_SET}
    if o1 and o2 and o1 != o2:
        return 'Lỗi lệch màu/xuất xứ', f'Lệch xuất xứ: {o1} vs {o2}'

    t_long, t_short = (t1, t2) if len(t1) >= len(t2) else (t2, t1)

    def is_covered(word):
        return (
            word in t_long
            or (len(word) >= 3 and any(word in token for token in t_long))
        )

    # --- LỚP 2: Con số / thông số kỹ thuật ---
    missing_specs = [
        w for w in t_short
        if re.search(r'\d', w) and not is_covered(w)
    ]

    if missing_specs:
        return (
            'Lỗi lệch con số/kích thước',
            f"Sai lệch thông số: {', '.join(missing_specs)}",
        )

    # --- LỚP 3: Khớp chứa nhau (bản tóm tắt) ---
    matched_words = sum(1 for w in t_short if is_covered(w))

    if matched_words == len(t_short):
        return (
            'Khớp chứa nhau',
            'Chuỗi ngắn là bản tóm tắt nằm trọn trong chuỗi dài',
        )

    # --- LỚP 4: Độ phủ từ còn lại ---
    coverage = matched_words / len(t_short) if t_short else 0.0

    if coverage >= threshold:
        return 'Khớp tương đối', f'Độ phủ từ đạt {coverage * 100:.1f}%'

    return (
        'Không khớp',
        f'Độ phủ từ chỉ đạt {coverage * 100:.1f}% (< {int(threshold * 100)}%)',
    )
def run_matching(df, col_a, col_target, prefix, ignore_class5=False):
    """Chạy matching cho một cặp cột và ghi kết quả vào df."""
    results = df.apply(
        lambda row: smart_specs_matching(
            row[col_a], row[col_target], ignore_class5=ignore_class5
        ),
        axis=1,
    )

    df[f'Trạng thái Desc ({prefix})'] = [r[0] for r in results]
    df[f'Chi tiết Desc ({prefix})'] = [r[1] for r in results]

    return df


# =============================================================================
# PHẦN 4: KIỂM TRA TRÙNG LẶP TRONG HỆ MÃ NEW
# =============================================================================
"""
Triết lý thiết kế (đáp ứng yêu cầu 4 & 5):

    KHÔNG kết luận nghi trùng dựa trên "độ giống nhau của câu chữ".
    Mỗi Description New được bóc tách thành một CHỮ KÝ CÓ CẤU TRÚC
    (ItemSignature), sau đó so khớp THUỘC TÍNH-VỚI-THUỘC TÍNH.

Vì sao phải so khớp theo TRƯỜNG (field) thay vì "túi từ" (bag of words)?
    Mô tả trong dữ liệu luôn có dạng phân tách bằng dấu phẩy:
        <LOẠI VẬT TƯ>,<TÊN/MODEL>,<THÔNG SỐ>,...,<MÀU>,<XUẤT XỨ>
    Làm phẳng thành tập hợp từ sẽ phá vỡ quan hệ nhãn-giá trị và sinh ra
    kết luận SAI. Hai ví dụ CÓ THẬT trong dữ liệu này:

      (a) 'NUT,HEX NUT,M5,SWCH,Ep-Fe/Zn 5/CM2(Cr3/White)'
          'NUT,HEX NUT,M4,SWCH,Ep-Fe/Zn 5/CM2(Cr3/White)'
          -> tập từ bản M5 lại là TẬP CON của bản M4 (vì số 5 vẫn còn trong
             'Zn 5'). Túi từ kết luận trùng. THỰC TẾ: khác cỡ ren.

      (b) 'LAMP,YSBRL34-DL11-RW,1x5(R/2+W/3),DC110V,...'
          'LAMP,YSBRL34-DL11-RW,1x5(R/3+W/2),DC110V,...'
          -> hai tập từ {1,5,r,2,w,3} giống hệt nhau. Túi từ kết luận trùng.
             THỰC TẾ: 2 đỏ+3 trắng khác 3 đỏ+2 trắng.

    So khớp theo trường giữ nguyên thứ tự bên trong từng trường nên chặn
    được cả hai lỗi trên.

Ba bài toán độc lập:
    4A. Item New nghi trùng Description New
        - Mức 1 (Trùng hoàn toàn): mọi trường khớp tuyệt đối sau chuẩn hóa.
        - Mức 2 (Thiếu thông tin) : một bên là tập con của bên kia VÀ phần
                                    chênh lệch không chứa bất kỳ con số nào.
    4B. Item New gắn với nhiều Item Code A khác nhau.
    4C. (xem Phần 5) Cảnh báo chất lượng dữ liệu.
"""

# ---------------------------------------------------------------------------
# 4.0. Cấu hình riêng cho Phần 4
# ---------------------------------------------------------------------------
# Phần 4 dùng bộ chuẩn hóa RIÊNG, KHÔNG sửa `tokenize_and_clean` của Phần 2
# để không làm thay đổi output của Phần 1-3.

PART4_PHRASE_RULES = [
    (re.compile(r'\bflame[\s\-]+retardant\b'), 'fr'),
    (re.compile(r'\bh[\s\-]*pvc\b'), 'hpvc'),
    (re.compile(r'\bcosmolink\s+vina\b'), 'cosmolink'),
    # 'VIET NAM' viết tách phải quy về 'vietnam', nếu không bộ lọc xuất xứ
    # sẽ bị vô hiệu với toàn bộ các dòng đang viết tách (159 lượt trong data).
    (re.compile(r'\bviet\s+nam\b'), 'vietnam'),
    (re.compile(r'\bsouth\s+korea\b'), 'korea'),
]

CLASS5_PATTERN = re.compile(r'\bclass\s*[-:=]?\s*5(?:\.0)?\b')
TOKEN_PATTERN = re.compile(r'\d+\.\d+|\d+|[a-z]+')
PARENTHETICAL_PATTERN = re.compile(r'\([^)]*\)')
DIGIT_PATTERN = re.compile(r'\d')

# Nhận diện cặp NHÃN-GIÁ TRỊ dính liền: 'M12' -> m=12, '110V' -> v=110
LABEL_BEFORE_VALUE = re.compile(r'\b([a-z]{1,3})(\d+(?:\.\d+)?)\b')
VALUE_BEFORE_LABEL = re.compile(r'\b(\d+(?:\.\d+)?)([a-z]{1,4})\b')

# Mã model/quy cách: cụm chữ-số có dấu nối, ví dụ 'ysbrl34-dl11-rw', 'din912'
MODEL_CODE_PATTERN = re.compile(r'[a-z0-9]+(?:[\-/][a-z0-9]+)*')

PART4_COLOR_SET = COLOR_SET | {'violet', 'pink', 'greenyellow'}

# Bổ sung các quốc gia thực sự xuất hiện trong dữ liệu nhưng thiếu ở ORIGIN_SET
PART4_ORIGIN_SET = ORIGIN_SET | {
    'italy', 'indonesia', 'turkey', 'malaysia',
    'india', 'usa', 'france', 'thailand', 'singapore',
}

LEVEL_EXACT = 'Mức 1 - Khớp tuyệt đối'
LEVEL_SUBSET = 'Mức 2 - Khớp chứa nhau (thiếu dữ kiện)'


# ---------------------------------------------------------------------------
# 4.1. Chuẩn hóa & bóc tách thuộc tính
# ---------------------------------------------------------------------------

def _normalize_raw_text(text):
    """Đưa mô tả thô về chữ thường, bỏ dấu, gộp cụm đồng nghĩa, bỏ CLASS 5."""
    s = remove_diacritics(str(text).lower())

    for pattern, replacement in PART4_PHRASE_RULES:
        s = pattern.sub(replacement, s)

    s = CLASS5_PATTERN.sub(' ', s)

    return normalize_numbers(s)


def _tokenize_field(text):
    """
    Tách một trường thành danh sách token, GIỮ NGUYÊN THỨ TỰ.

    Chỉ sắp xếp lại các token MÀU với nhau (giữ nguyên vị trí các token khác).
    Nhờ vậy 'GREEN/YELLOW' == 'YELLOW/GREEN' (cùng một loại dây tiếp địa)
    nhưng 'R/2+W/3' vẫn KHÁC 'R/3+W/2'.
    """
    tokens = [
        WORD_MAP.get(token, token)
        for token in TOKEN_PATTERN.findall(text)
    ]

    color_positions = [
        index for index, token in enumerate(tokens)
        if token in PART4_COLOR_SET
    ]

    if len(color_positions) > 1:
        for index, color in zip(
            color_positions,
            sorted(tokens[index] for index in color_positions),
        ):
            tokens[index] = color

    return tokens


def _extract_measures(text):
    """
    Bóc tách thông số kỹ thuật dạng NHÃN-GIÁ TRỊ, giữ nguyên ràng buộc.

        'M12,L35' -> {'m': {'12'}, 'l': {'35'}}
        'DC110V'  -> {'v': {'110'}, 'dc': ...}
        '1Cx2.5'  -> {'c': {'1'}, 'x': {'2.5'}}

    Nhờ giữ ràng buộc này, 'M10,L20' KHÁC 'M20,L10' - điều mà cách xử lý cũ
    (cắt bỏ tiền tố L/W/H/D/T rồi so số trần) không phân biệt được.
    """
    measures = {}

    for label, value in LABEL_BEFORE_VALUE.findall(text):
        measures.setdefault(label, set()).add(value)

    for value, label in VALUE_BEFORE_LABEL.findall(text):
        measures.setdefault(label, set()).add(value)

    return {label: frozenset(values) for label, values in measures.items()}


@dataclass(frozen=True)
class ItemSignature:
    """Chữ ký có cấu trúc của một Description New."""

    category: tuple = ()                                  # Tên/loại vật tư
    fields: Counter = dc_field(default_factory=Counter)   # Bội tập các trường
    field_source: tuple = ()                              # (chuẩn hóa -> nguyên văn)
    colors: frozenset = frozenset()                       # Màu sắc
    origins: frozenset = frozenset()                      # Nguồn gốc/xuất xứ
    measures: tuple = ()                                  # Cặp (nhãn, giá trị)
    model_codes: frozenset = frozenset()                  # Mã model/quy cách

    @property
    def is_empty(self):
        return not self.fields

    @property
    def measure_map(self):
        return dict(self.measures)

    @property
    def source_map(self):
        return dict(self.field_source)

    @property
    def technical_text(self):
        return ' | '.join(
            f'{label.upper()}={"/".join(sorted(values))}'
            for label, values in self.measures
        )


def build_signature(description):
    """Bóc tách một Description New thành ItemSignature."""
    if description is None or pd.isna(description):
        return ItemSignature()

    normalized = _normalize_raw_text(description)

    # --- Các trường ngăn cách bởi dấu phẩy ---
    # Giữ song song bản chuẩn hóa (để so khớp) và bản nguyên văn (để in ra
    # báo cáo cho người đọc).
    raw_parts = str(description).split(',')
    normalized_parts = normalized.split(',')

    field_token_lists = []
    field_source = {}

    for index, part in enumerate(normalized_parts):
        tokens = _tokenize_field(part)

        if not tokens:
            continue

        field_token_lists.append(tokens)

        key = ' '.join(tokens)
        raw_text = (
            raw_parts[index].strip() if index < len(raw_parts) else key
        )
        field_source.setdefault(key, raw_text or key)

    if not field_token_lists:
        return ItemSignature()

    # --- Màu & xuất xứ: bỏ nội dung trong ngoặc ---
    # Ngoặc thường là chú giải viết tắt ('B-W-R-BK(Blue-White-Red-Black)');
    # nếu tính vào sẽ báo lệch màu giả.
    outside_tokens = _tokenize_field(
        _normalize_raw_text(
            PARENTHETICAL_PATTERN.sub(' ', str(description))
        )
    )

    return ItemSignature(
        category=tuple(field_token_lists[0]),
        fields=Counter(' '.join(tokens) for tokens in field_token_lists),
        field_source=tuple(sorted(field_source.items())),
        colors=frozenset(t for t in outside_tokens if t in PART4_COLOR_SET),
        origins=frozenset(t for t in outside_tokens if t in PART4_ORIGIN_SET),
        measures=tuple(sorted(_extract_measures(normalized).items())),
        model_codes=frozenset(
            token
            for token in MODEL_CODE_PATTERN.findall(normalized)
            if DIGIT_PATTERN.search(token) and re.search(r'[a-z]', token)
        ),
    )


# ---------------------------------------------------------------------------
# 4.2. So khớp hai chữ ký
# ---------------------------------------------------------------------------

def _format_fields(field_counter, *source_maps):
    """In các trường ra dạng NGUYÊN VĂN để người đọc báo cáo hiểu ngay."""
    lookup = {}

    for source_map in source_maps:
        lookup.update(source_map)

    return ' | '.join(
        lookup.get(key, key) for key in sorted(field_counter.elements())
    )


def compare_signatures(signature_1, signature_2):
    """
    So khớp hai chữ ký theo từng THUỘC TÍNH.

    Trả về (level, reason, missing_text):
        level        : None nếu KHÔNG nghi trùng.
        reason       : lý do kết luận (hoặc lý do loại trừ).
        missing_text : phần dữ kiện mà bên ngắn còn thiếu (chỉ có ở Mức 2).

    Thứ tự kiểm tra đi từ thuộc tính "đắt giá" nhất xuống, nên lý do trả về
    luôn là nguyên nhân loại trừ đầu tiên - dễ giải thích cho nghiệp vụ.
    """
    if signature_1.is_empty or signature_2.is_empty:
        return None, 'Thiếu mô tả', ''

    # --- Lớp 1: Tên/loại vật tư ---
    if signature_1.category != signature_2.category:
        return None, (
            f"Khác tên vật tư: '{' '.join(signature_1.category)}' "
            f"vs '{' '.join(signature_2.category)}'"
        ), ''

    # --- Lớp 2: Màu sắc (chỉ loại khi CẢ HAI cùng có nhưng khác nhau) ---
    if (
        signature_1.colors and signature_2.colors
        and signature_1.colors != signature_2.colors
    ):
        return None, (
            f'Khác màu sắc: {sorted(signature_1.colors)} '
            f'vs {sorted(signature_2.colors)}'
        ), ''

    # --- Lớp 3: Nguồn gốc / xuất xứ ---
    if (
        signature_1.origins and signature_2.origins
        and signature_1.origins != signature_2.origins
    ):
        return None, (
            f'Khác xuất xứ: {sorted(signature_1.origins)} '
            f'vs {sorted(signature_2.origins)}'
        ), ''

    # --- Lớp 4: Thông số kỹ thuật theo cặp nhãn-giá trị ---
    # Đây là lớp bảo vệ CHÍNH cho yêu cầu "khác thông số kỹ thuật thì không
    # được tính là nghi trùng". Vì `measures` được bóc từ TOÀN BỘ mô tả, mọi
    # mâu thuẫn nhãn-giá trị đều bị chặn tại đây, kể cả khi thông số đó nằm
    # ở phần chênh lệch giữa hai bên.
    map_1, map_2 = signature_1.measure_map, signature_2.measure_map

    conflicts = sorted(
        f'{label.upper()}={sorted(map_1[label])} vs {sorted(map_2[label])}'
        for label in map_1.keys() & map_2.keys()
        if map_1[label] != map_2[label]
    )

    if conflicts:
        return None, 'Khác thông số kỹ thuật: ' + '; '.join(conflicts), ''

    # --- Lớp 5: Mã model / quy cách ---
    # Chỉ loại khi MỖI BÊN đều có model riêng (tức là hai model đối chọi
    # nhau). Nếu chỉ một bên có model còn bên kia không ghi thì đó là
    # "thiếu dữ kiện", để Lớp 7 xử lý.
    only_in_1 = signature_1.model_codes - signature_2.model_codes
    only_in_2 = signature_2.model_codes - signature_1.model_codes

    if only_in_1 and only_in_2:
        return None, (
            f'Khác model/quy cách: {sorted(only_in_1)} vs {sorted(only_in_2)}'
        ), ''

    # --- Lớp 6: So khớp bội tập các trường mô tả ---
    source_1, source_2 = signature_1.source_map, signature_2.source_map

    extra_in_1 = signature_1.fields - signature_2.fields
    extra_in_2 = signature_2.fields - signature_1.fields

    if not extra_in_1 and not extra_in_2:
        return LEVEL_EXACT, 'Mọi trường mô tả khớp tuyệt đối sau chuẩn hóa', ''

    # Cả hai bên đều có trường riêng -> đây là MÂU THUẪN thông tin, không
    # phải thiếu thông tin. Ví dụ 'Words: English' vs 'Words: English and
    # Vietnamese' (nhãn song ngữ) hay 'M5' vs 'M4' (khác cỡ ren).
    if extra_in_1 and extra_in_2:
        return None, (
            f'Khác nội dung mô tả: '
            f'[{_format_fields(extra_in_1, source_1)}] '
            f'vs [{_format_fields(extra_in_2, source_2)}]'
        ), ''

    # --- Lớp 7: KHỚP CHỨA NHAU (một bên là tập con của bên kia) ---
    # Đến đây đã chắc chắn: cùng tên vật tư, cùng màu, cùng xuất xứ, KHÔNG
    # có cặp nhãn-giá trị nào mâu thuẫn (Lớp 4) và không có model đối chọi
    # (Lớp 5). Phần chênh lệch vì vậy là thông tin mà bên ngắn KHÔNG ĐỦ DỮ
    # KIỆN để so sánh, chứ không phải thông số khác biệt.
    #
    # LƯU Ý QUAN TRỌNG: ở đây KHÔNG chặn theo "phần thừa có chứa số".
    # Chặn như vậy sẽ loại oan các trường hợp đúng nghiệp vụ, ví dụ:
    #     'BOLT,HEX BOLT,M6,L25,...'
    #     'BOLT,HEX BOLT,P/W*2+S/W*1+NUT*1,M6,L25,...'
    # Việc bảo vệ khỏi lệch thông số đã do Lớp 4 đảm nhiệm: nếu phần thừa
    # chứa một nhãn kỹ thuật đã tồn tại với giá trị khác (M6 vs M8, L25 vs
    # L30...) thì cặp đó đã bị loại từ trước khi tới đây.
    missing_text = _format_fields(
        extra_in_1 or extra_in_2, source_1, source_2
    )

    return (
        LEVEL_SUBSET,
        f'Mô tả ngắn nằm trọn trong mô tả dài; bên ngắn thiếu dữ kiện: '
        f'[{missing_text}]',
        missing_text,
    )


# ---------------------------------------------------------------------------
# 4.3. Gom nhóm bắc cầu (Union-Find)
# ---------------------------------------------------------------------------

class UnionFind:
    """Gom các mã nghi trùng thành nhóm theo quan hệ bắc cầu."""

    def __init__(self):
        self._parent = {}

    def find(self, item):
        self._parent.setdefault(item, item)

        while self._parent[item] != item:
            self._parent[item] = self._parent[self._parent[item]]
            item = self._parent[item]

        return item

    def union(self, item_1, item_2):
        root_1, root_2 = self.find(item_1), self.find(item_2)

        if root_1 != root_2:
            self._parent[root_2] = root_1

    def groups(self):
        result = {}

        for item in self._parent:
            result.setdefault(self.find(item), []).append(item)

        return result


# ---------------------------------------------------------------------------
# 4.4. Chuẩn bị bảng Item New duy nhất
# ---------------------------------------------------------------------------

def join_unique_values(series):
    """Gộp các giá trị duy nhất, bỏ giá trị rỗng."""
    values = {
        str(value).strip()
        for value in series
        if pd.notna(value) and str(value).strip()
    }

    return ' | '.join(sorted(values))


def build_new_item_table(df):
    """
    Trích bảng quan hệ Item New <-> Code A <-> Ksys, bỏ mã rỗng/placeholder.

    Mã placeholder (toàn số 0, ví dụ '00000000000') nghĩa là 'chưa cấp mã
    mới' - nếu để lại sẽ gom nhầm hàng trăm vật tư khác nhau vào một nhóm.
    """
    columns = [
        'item new', 'Description new',
        'item code A', 'Description code A', 'item ksys',
    ]

    table = df[columns].copy()

    for column in ['item new', 'item code A', 'item ksys']:
        table[column] = table[column].fillna('').astype(str).str.strip()

    is_real_code = (
        table['item new'].ne('')
        & ~table['item new'].str.fullmatch(r'0+')
    )
    has_description = (
        table['Description new'].notna()
        & table['Description new'].astype(str).str.strip().ne('')
    )

    return (
        table[is_real_code & has_description]
        .drop_duplicates(keep='first')
        .reset_index(drop=True)
    )


def aggregate_by_new_item(new_item_table):
    """Gộp về mức (item new, Description new) và gắn chữ ký cấu trúc."""
    aggregated = (
        new_item_table
        .groupby(['item new', 'Description new'], as_index=False)
        .agg({
            'item code A': join_unique_values,
            'Description code A': join_unique_values,
            'item ksys': join_unique_values,
        })
    )

    aggregated['signature'] = aggregated['Description new'].map(build_signature)

    aggregated = aggregated[
        aggregated['signature'].map(lambda s: not s.is_empty)
    ].reset_index(drop=True)

    aggregated['Thông số kỹ thuật'] = aggregated['signature'].map(
        lambda s: s.technical_text
    )
    aggregated['Màu sắc'] = aggregated['signature'].map(
        lambda s: '|'.join(sorted(s.colors))
    )
    aggregated['Xuất xứ'] = aggregated['signature'].map(
        lambda s: '|'.join(sorted(s.origins))
    )
    aggregated['Category key'] = aggregated['signature'].map(
        lambda s: ' '.join(s.category)
    )

    return aggregated


# ---------------------------------------------------------------------------
# 4.5. BÀI TOÁN 4A - Item New nghi trùng Description New
# ---------------------------------------------------------------------------

PAIR_COLUMNS = [
    'Nhóm trùng số',
    'item new 1', 'Description new 1',
    'item code A 1', 'Description code A 1', 'item ksys 1',
    'item new 2', 'Description new 2',
    'item code A 2', 'Description code A 2', 'item ksys 2',
    'Thông số kỹ thuật', 'Màu sắc', 'Xuất xứ',
    'Trạng thái', 'Chi tiết', 'Mã có mô tả đầy đủ hơn', 'Dữ kiện còn thiếu',
]

GROUP_COLUMNS = [
    'Nhóm trùng số', 'item new', 'Description new',
    'item code A', 'Description code A', 'item ksys',
    'Thông số kỹ thuật', 'Màu sắc', 'Xuất xứ',
]


def find_duplicate_pairs(aggregated):
    """
    Sinh danh sách cặp nghi trùng.

    Chặn (blocking) theo 'Category key' để chỉ so các vật tư CÙNG LOẠI -
    vừa đúng nghiệp vụ (khác tên vật tư thì không thể trùng) vừa giảm mạnh
    số phép so sánh.
    """
    pairs = []

    for _, block in aggregated.groupby('Category key', sort=False):
        if len(block) < 2:
            continue

        rows = block.to_dict('records')

        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                row_1, row_2 = rows[i], rows[j]

                if row_1['item new'] == row_2['item new']:
                    continue

                level, reason, missing_text = compare_signatures(
                    row_1['signature'], row_2['signature']
                )

                if level is None:
                    continue

                # Chuẩn hóa thứ tự cặp để kết quả ổn định giữa các lần chạy
                first, second = sorted(
                    (row_1, row_2), key=lambda r: r['item new']
                )

                # Mã nào có mô tả đầy đủ hơn -> mã còn lại là mã thiếu dữ kiện
                fuller = max(
                    (row_1, row_2),
                    key=lambda r: sum(r['signature'].fields.values()),
                )

                pairs.append({
                    'level': level,
                    'reason': reason,
                    'missing': missing_text,
                    'fuller': fuller['item new'] if missing_text else '',
                    'first': first,
                    'second': second,
                })

    return pairs


def build_duplicate_reports(pairs, level):
    """Dựng cặp bảng (nhóm, chi tiết cặp) cho một mức nghi trùng."""
    selected = [pair for pair in pairs if pair['level'] == level]

    if not selected:
        return (
            pd.DataFrame(columns=GROUP_COLUMNS),
            pd.DataFrame(columns=PAIR_COLUMNS),
        )

    # --- Gom nhóm bắc cầu ---
    union_find = UnionFind()

    for pair in selected:
        union_find.union(pair['first']['item new'], pair['second']['item new'])

    group_id_by_item = {}

    for number, (_, members) in enumerate(
        sorted(union_find.groups().items(), key=lambda kv: min(kv[1])),
        start=1,
    ):
        for member in members:
            group_id_by_item[member] = number

    # --- Bảng chi tiết cặp ---
    pair_records = []

    for pair in selected:
        first, second = pair['first'], pair['second']

        pair_records.append({
            'Nhóm trùng số': group_id_by_item[first['item new']],

            'item new 1': first['item new'],
            'Description new 1': first['Description new'],
            'item code A 1': first['item code A'],
            'Description code A 1': first['Description code A'],
            'item ksys 1': first['item ksys'],

            'item new 2': second['item new'],
            'Description new 2': second['Description new'],
            'item code A 2': second['item code A'],
            'Description code A 2': second['Description code A'],
            'item ksys 2': second['item ksys'],

            'Thông số kỹ thuật': first['Thông số kỹ thuật'],
            'Màu sắc': first['Màu sắc'],
            'Xuất xứ': first['Xuất xứ'],

            'Trạng thái': pair['level'],
            'Chi tiết': pair['reason'],
            'Mã có mô tả đầy đủ hơn': pair['fuller'],
            'Dữ kiện còn thiếu': pair['missing'],
        })

    df_pairs = (
        pd.DataFrame(pair_records, columns=PAIR_COLUMNS)
        .drop_duplicates(keep='first')
        .sort_values(['Nhóm trùng số', 'item new 1', 'item new 2'])
        .reset_index(drop=True)
    )

    # --- Bảng nhóm ---
    member_rows = {}

    for pair in selected:
        for side in ('first', 'second'):
            row = pair[side]
            member_rows[row['item new']] = row

    df_groups = pd.DataFrame([
        {
            'Nhóm trùng số': group_id_by_item[item_new],
            'item new': item_new,
            'Description new': row['Description new'],
            'item code A': row['item code A'],
            'Description code A': row['Description code A'],
            'item ksys': row['item ksys'],
            'Thông số kỹ thuật': row['Thông số kỹ thuật'],
            'Màu sắc': row['Màu sắc'],
            'Xuất xứ': row['Xuất xứ'],
        }
        for item_new, row in member_rows.items()
    ], columns=GROUP_COLUMNS)

    df_groups = (
        df_groups
        .sort_values(['Nhóm trùng số', 'item new'])
        .reset_index(drop=True)
    )

    return df_groups, df_pairs


# ---------------------------------------------------------------------------
# 4.6. BÀI TOÁN 4B - Item New gắn nhiều Item Code A
# ---------------------------------------------------------------------------

MULTI_CODE_A_SUMMARY_COLUMNS = [
    'item new',
    'Số lượng Item Code A',
    'Danh sách Item Code A',
    'Danh sách Description New',
    'Danh sách Item Ksys',
]

CODE_A_VERDICT_COLUMNS = [
    'item new', 'Description new',
    'item code A 1', 'Description code A 1',
    'item code A 2', 'Description code A 2',
    'Kết luận', 'Chi tiết', 'Dữ kiện còn thiếu',
]


def find_items_with_multiple_code_a(new_item_table):
    """Tìm Item New được gắn với nhiều hơn một Item Code A."""
    relation = new_item_table[new_item_table['item code A'].ne('')]

    code_a_count = (
        relation
        .drop_duplicates(subset=['item new', 'item code A'], keep='first')
        .groupby('item new')['item code A']
        .nunique()
    )

    flagged_items = code_a_count[code_a_count > 1].index

    df_detail = (
        relation[relation['item new'].isin(flagged_items)]
        .drop_duplicates(keep='first')
        [[
            'item new', 'Description new',
            'item code A', 'Description code A', 'item ksys',
        ]]
        .sort_values(['item new', 'item code A'])
        .reset_index(drop=True)
    )

    if df_detail.empty:
        return (
            pd.DataFrame(columns=MULTI_CODE_A_SUMMARY_COLUMNS),
            df_detail,
        )

    df_summary = (
        df_detail
        .groupby('item new', as_index=False)
        .agg(**{
            'Số lượng Item Code A': ('item code A', 'nunique'),
            'Danh sách Item Code A': ('item code A', join_unique_values),
            'Danh sách Description New': ('Description new', join_unique_values),
            'Danh sách Item Ksys': ('item ksys', join_unique_values),
        })
        .sort_values(
            ['Số lượng Item Code A', 'item new'], ascending=[False, True]
        )
        .reset_index(drop=True)
    )

    return df_summary, df_detail


def judge_code_a_within_new_item(df_multi_code_a_detail):
    """
    Với mỗi Item New gắn nhiều Item Code A: so chính các mô tả Code A đó
    với nhau bằng đúng engine của 4A.

    Ý nghĩa nghiệp vụ - trả lời câu hỏi "gộp mã như vậy có đúng không?":
      * Nếu các Code A khớp tuyệt đối / khớp chứa nhau -> hệ Code A cũ vốn
        đã tạo trùng, việc gộp về một mã New là ĐÚNG.
      * Nếu các Code A khác thông số/màu/xuất xứ -> đây là LỖI ÁNH XẠ:
        nhiều vật tư khác nhau bị gán nhầm chung một mã New.
    """
    records = []

    for item_new, block in df_multi_code_a_detail.groupby('item new'):
        rows = (
            block
            .drop_duplicates(subset=['item code A'], keep='first')
            .to_dict('records')
        )

        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                first, second = sorted(
                    (rows[i], rows[j]), key=lambda r: r['item code A']
                )

                level, reason, missing_text = compare_signatures(
                    build_signature(first['Description code A']),
                    build_signature(second['Description code A']),
                )

                records.append({
                    'item new': item_new,
                    'Description new': first['Description new'],
                    'item code A 1': first['item code A'],
                    'Description code A 1': first['Description code A'],
                    'item code A 2': second['item code A'],
                    'Description code A 2': second['Description code A'],
                    'Kết luận': (
                        f'Gộp hợp lý ({level})' if level
                        else 'LỖI ÁNH XẠ - Code A không trùng nhau'
                    ),
                    'Chi tiết': reason,
                    'Dữ kiện còn thiếu': missing_text,
                })

    return (
        pd.DataFrame(records, columns=CODE_A_VERDICT_COLUMNS)
        .sort_values(['Kết luận', 'item new', 'item code A 1'])
        .reset_index(drop=True)
    )


# =============================================================================
# PHẦN 5: CẢNH BÁO CHẤT LƯỢNG DỮ LIỆU
# =============================================================================
# Các trường hợp trước đây bị loại âm thầm hoặc chưa được kiểm tra.
# Tách riêng nên KHÔNG ảnh hưởng kết quả Phần 1-4.

def build_data_quality_report(df_full, df_dropped, new_item_table):
    """Tổng hợp các cảnh báo dữ liệu thành một bảng phẳng, dễ lọc."""
    records = []

    def add(loai, ma, mo_ta):
        records.append({'Loại cảnh báo': loai, 'Mã': ma, 'Chi tiết': mo_ta})

    # (1) Dòng bị loại vì thiếu Description New -> KHÔNG được kiểm tra
    for _, row in df_dropped.iterrows():
        add(
            'Thiếu Description New (không kiểm tra được)',
            str(row['item code A']),
            f"Code A: {row['Description code A']} | item new: {row['item new']}",
        )

    # (2) Mã New placeholder (toàn số 0) = chưa cấp mã
    placeholder = df_full[df_full['item new'].str.fullmatch(r'0+', na=False)]

    for _, row in placeholder.iterrows():
        add(
            'Item New là mã placeholder (toàn số 0)',
            str(row['item code A']),
            f"item new = '{row['item new']}' | {row['Description code A']}",
        )

    # (3) Một Item New nhưng có nhiều Description New khác nhau
    desc_count = new_item_table.groupby('item new')['Description new'].nunique()

    for item_new in desc_count[desc_count > 1].index:
        descriptions = new_item_table.loc[
            new_item_table['item new'] == item_new, 'Description new'
        ]
        add(
            'Item New có nhiều Description New khác nhau',
            item_new,
            join_unique_values(descriptions),
        )

    # (4) Một Item Code A ánh xạ sang nhiều Item New (dấu hiệu tách mã sai)
    code_a_map = (
        new_item_table[new_item_table['item code A'].ne('')]
        .groupby('item code A')['item new']
        .nunique()
    )

    for code_a in code_a_map[code_a_map > 1].index:
        items = new_item_table.loc[
            new_item_table['item code A'] == code_a, 'item new'
        ]
        add(
            'Một Item Code A ánh xạ nhiều Item New',
            code_a,
            join_unique_values(items),
        )

    # (5) Một Item New ứng với nhiều Item Ksys
    ksys_map = (
        new_item_table[new_item_table['item ksys'].ne('')]
        .groupby('item new')['item ksys']
        .nunique()
    )

    for item_new in ksys_map[ksys_map > 1].index:
        ksys_values = new_item_table.loc[
            new_item_table['item new'] == item_new, 'item ksys'
        ]
        add(
            'Item New ứng với nhiều Item Ksys',
            item_new,
            join_unique_values(ksys_values),
        )

    return pd.DataFrame(
        records, columns=['Loại cảnh báo', 'Mã', 'Chi tiết']
    )


UOM_CONFLICT_COLUMNS = ['Mức nghi trùng', 'Nhóm trùng số', 'item new', 'UOM new']


def check_uom_conflict_in_groups(group_frames_by_level, df_full):
    """
    Trong mỗi nhóm nghi trùng, kiểm tra các Item New có cùng UOM hay không.

    Cùng mô tả nhưng khác đơn vị tính thường là khác quy cách đóng gói,
    cần người review xác nhận trước khi gộp mã.

    `group_frames_by_level`: dict {tên mức: DataFrame nhóm}. Số nhóm được
    đánh lại độc lập ở từng mức nên phải gộp kèm nhãn mức, tránh trộn nhầm
    'Nhóm 1 - Mức 1' với 'Nhóm 1 - Mức 2'.
    """
    frames = []

    for level, df_groups in group_frames_by_level.items():
        if df_groups.empty:
            continue

        frame = df_groups[['Nhóm trùng số', 'item new']].copy()
        frame.insert(0, 'Mức nghi trùng', level)
        frames.append(frame)

    if not frames:
        return pd.DataFrame(columns=UOM_CONFLICT_COLUMNS)

    merged = pd.concat(frames, ignore_index=True)

    uom_by_item = (
        df_full[df_full['item new'].ne('')]
        .groupby('item new')['uom new']
        .apply(join_unique_values)
    )

    merged['UOM new'] = merged['item new'].map(uom_by_item).fillna('')

    conflict = (
        merged
        .groupby(['Mức nghi trùng', 'Nhóm trùng số'], sort=False)
        .filter(lambda block: block['UOM new'].nunique() > 1)
    )

    return (
        conflict[UOM_CONFLICT_COLUMNS]
        .sort_values(['Mức nghi trùng', 'Nhóm trùng số', 'item new'])
        .reset_index(drop=True)
    )


# =============================================================================
# ĐIỀU PHỐI & XUẤT BÁO CÁO
# =============================================================================

def main():
    # ---- Đọc & làm sạch ----
    df_full = load_dataset()
    df, df_dropped = drop_invalid_description_new(df_full)

    # ---- Phần 1 & 2 ----
    df = compare_uom(df)
    df = run_matching(
        df, 'Description code A', 'Description ksys',
        'Code A vs Ksys', ignore_class5=False,
    )
    df = run_matching(
        df, 'Description code A', 'Description new',
        'Code A vs New', ignore_class5=True,
    )

    # ---- Phần 3: lọc danh sách lệch ----
    df_uom_diff_ksys = df[~df['uom_ksys_match']].copy()
    df_uom_diff_new = df[~df['uom_new_match']].copy()

    df_desc_diff_ksys = df[
        ~df['Trạng thái Desc (Code A vs Ksys)'].isin(MATCH_STATUSES)
    ].copy()

    df_desc_diff_new = (
        df[~df['Trạng thái Desc (Code A vs New)'].isin(MATCH_STATUSES)]
        .drop_duplicates(keep='first')
        .reset_index(drop=True)
    )

    # ---- Phần 4 ----
    new_item_table = build_new_item_table(df)
    aggregated = aggregate_by_new_item(new_item_table)

    pairs = find_duplicate_pairs(aggregated)

    df_exact_groups, df_exact_pairs = build_duplicate_reports(pairs, LEVEL_EXACT)
    df_subset_groups, df_subset_pairs = build_duplicate_reports(pairs, LEVEL_SUBSET)

    df_multi_code_a_summary, df_multi_code_a_detail = (
        find_items_with_multiple_code_a(new_item_table)
    )
    df_code_a_verdict = judge_code_a_within_new_item(df_multi_code_a_detail)

    # ---- Phần 5 ----
    df_data_quality = build_data_quality_report(
        df_full, df_dropped, new_item_table
    )
    df_uom_conflict = check_uom_conflict_in_groups(
        {
            LEVEL_EXACT: df_exact_groups,
            LEVEL_SUBSET: df_subset_groups,
        },
        df,
    )

    # ---- Xuất Excel ----
    sheets = {
        'Lệch Desc (Code A vs Ksys)': df_desc_diff_ksys,
        'Lệch Desc (Code A vs New)': df_desc_diff_new,
        'Lệch UOM (Code A vs Ksys)': df_uom_diff_ksys,
        'Lệch UOM (Code A vs New)': df_uom_diff_new,
        'Toàn bộ Data Checked': df,

        'Item New trùng Description': df_exact_groups,
        'Chi tiết Item New trùng': df_exact_pairs,
        'Nghi trùng khớp chứa nhau': df_subset_groups,
        'Chi tiết khớp chứa nhau': df_subset_pairs,

        'Item New nhiều Code A': df_multi_code_a_summary,
        'Chi tiết New nhiều Code A': df_multi_code_a_detail,
        'Đối chiếu Code A trong 1 New': df_code_a_verdict,

        'Cảnh báo dữ liệu': df_data_quality,
        'Cảnh báo lệch UOM nhóm trùng': df_uom_conflict,
    }

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

    # ---- Thống kê ----
    print(f"--> Xuất thành công file gộp: '{OUTPUT_FILE}'")
    print(f"    + Lệch UOM  : Ksys ({len(df_uom_diff_ksys)}) "
          f"| New ({len(df_uom_diff_new)})")
    print(f"    + Lệch Desc : Ksys ({len(df_desc_diff_ksys)}) "
          f"| New ({len(df_desc_diff_new)})")
    print(f"    + [4A] Khớp tuyệt đối  : "
          f"{df_exact_groups['Nhóm trùng số'].nunique()} nhóm / "
          f"{df_exact_groups['item new'].nunique()} mã / "
          f"{len(df_exact_pairs)} cặp")
    print(f"    + [4A] Khớp chứa nhau  : "
          f"{df_subset_groups['Nhóm trùng số'].nunique()} nhóm / "
          f"{df_subset_groups['item new'].nunique()} mã / "
          f"{len(df_subset_pairs)} cặp")
    print(f"    + [4B] Item New nhiều Code A: "
          f"{len(df_multi_code_a_summary)} mã / "
          f"{len(df_code_a_verdict)} cặp Code A đối chiếu")

    if not df_code_a_verdict.empty:
        for verdict, count in (
            df_code_a_verdict['Kết luận'].value_counts().items()
        ):
            print(f"           - {verdict}: {count} cặp")

    print(f"    + [5]  Cảnh báo dữ liệu     : {len(df_data_quality)} dòng")

    return locals()


if __name__ == '__main__':
    main()