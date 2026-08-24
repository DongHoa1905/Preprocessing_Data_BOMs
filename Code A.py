# -*- coding: utf-8 -*-
"""
=============================================================================
 KIỂM TRA ĐỐI CHIẾU DANH MỤC VẬT TƯ: CODE A  <->  KSYS  <->  CODE NEW
=============================================================================
 Phần 1 : So sánh UOM
 Phần 2 : So sánh Description (dùng chung engine chữ ký của Phần 4)
 Phần 3 : Xuất báo cáo lệch UOM / Description
 Phần 4 : Kiểm tra trùng lặp trong hệ mã New
 Phần 5 : Cảnh báo chất lượng dữ liệu

 Nguyên tắc phân loại (áp dụng cho cả Phần 2 và Phần 4):
   - KHÔNG kết luận theo độ giống câu chữ. Mỗi mô tả được bóc thành CHỮ KÝ
     có cấu trúc (loại vật tư / các trường / màu / xuất xứ / thông số) rồi
     so THUỘC TÍNH-VỚI-THUỘC TÍNH.
   - Kết quả rơi vào một trong các nhãn: Khớp Description, Khớp chứa nhau,
     Lệch Item Name (khớp phần sau), hoặc các loại Lệch (loại vật tư /
     thông số / màu / xuất xứ).
   - Thứ tự từ được GIỮ NGUYÊN để phân biệt màu ('GREEN/YELLOW' khác
     'YELLOW/GREEN') và cấu hình ('R/2+W/3' khác 'R/3+W/2'). 
=============================================================================
"""

import re
import unicodedata
from collections import Counter, OrderedDict
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

PLACEHOLDER_CODE_PATTERN = re.compile(r'^0+$')
PLACEHOLDER_CODE_VALUES = {'n/a', 'na', 'nan', 'none', ''}


def is_placeholder_code(value):
    """Mã placeholder: toàn số 0 hoặc N/A -> coi như CHƯA CẤP MÃ, bỏ qua."""
    text = str(value).strip().lower()
    return (
        text in PLACEHOLDER_CODE_VALUES
        or bool(PLACEHOLDER_CODE_PATTERN.match(text))
    )


# =============================================================================
# 1. ĐỌC & LÀM SẠCH DỮ LIỆU
# =============================================================================

def load_dataset(file_name=INPUT_FILE):
    raw = pd.read_excel(file_name)
    df = raw[list(COLUMN_RENAME_MAP.keys())].rename(columns=COLUMN_RENAME_MAP)

    for column in CODE_COLUMNS:
        df[column] = (
            df[column].fillna('').astype(str)
            .str.replace(r'\.0$', '', regex=True).str.strip()
        )

    df[FFILL_COLUMNS] = df[FFILL_COLUMNS].ffill()
    return df


def drop_invalid_description_new(df):
    has_value = df['Description new'].notna()
    cleaned = df['Description new'].astype(str).str.strip().str.lower()
    is_invalid_value = cleaned.isin(INVALID_DESCRIPTION_VALUES)
    has_no_letter = ~cleaned.str.contains(r'[a-zA-ZÀ-ỹ]', regex=True, na=True)
    keep = has_value & ~(is_invalid_value | has_no_letter)
    return df[keep].copy(), df[~keep].copy()


# =============================================================================
# PHẦN 1: SO SÁNH UOM
# =============================================================================

UOM_MAPPING = {'rol': 'roll', 'sht': 'sheet', 'pnl': 'panel'}


def clean_uom(series):
    return (
        series.fillna('').astype(str)
        .str.replace(r'[\r\n\t\xa0]', '', regex=True)
        .str.strip().str.lower().replace(UOM_MAPPING)
    )


def compare_uom(df):
    uom_a = clean_uom(df['uom code A'])
    df['uom_ksys_match'] = uom_a == clean_uom(df['uom ksys'])
    df['uom_new_match'] = uom_a == clean_uom(df['uom new'])
    df['Trạng thái UOM (Code A vs Ksys)'] = df['uom_ksys_match'].map(
        {True: 'Khớp', False: 'Lệch'})
    df['Trạng thái UOM (Code A vs New)'] = df['uom_new_match'].map(
        {True: 'Khớp', False: 'Lệch'})
    return df


# =============================================================================
# 2. CHUẨN HÓA & TỪ ĐIỂN
# =============================================================================

WORD_MAP = {
    # Màu sắc (IEC 60757 & biến thể)
    'bk': 'black', 'blk': 'black', 'bla': 'black',
    'bn': 'brown', 'brn': 'brown',
    'rd': 'red',
    'og': 'orange', 'org': 'orange', 'orn': 'orange',
    'ye': 'yellow', 'yel': 'yellow', 'yl': 'yellow',
    'gn': 'green', 'grn': 'green',
    'blu': 'blue',
    'vt': 'violet', 'pur': 'violet', 'prp': 'violet',
    'gy': 'grey', 'gry': 'grey', 'gray': 'grey',
    'wh': 'white', 'wht': 'white',
    'pk': 'pink', 'pnk': 'pink',
    'gnye': 'greenyellow', 'gnyel': 'greenyellow', 'yegn': 'greenyellow',
    # Xuất xứ
    'vnm': 'vietnam', 'cn': 'china', 'chn': 'china',
    'kr': 'korea', 'kor': 'korea', 'jp': 'japan', 'jpn': 'japan',
    'tw': 'taiwan', 'twn': 'taiwan', 'deu': 'germany', 'vina': 'vietnam',
    # Thông số (mm² viết nhiều kiểu -> 1 dạng chuẩn)
    'sq': 'sqmm', 'sqmm': 'sqmm', 'mmsq': 'sqmm', 'sqm': 'sqmm',
    'fr': 'flameretardant', 'hpvc': 'hpvc',
}
# ĐÃ BỎ 'vn'->vietnam: trong dữ liệu 'Vn' là điện áp danh định (voltage
# nominal, vd '1.1Vn/CONT'), không phải Việt Nam. 'viet nam' và 'vnm' vẫn
# được gộp về vietnam qua PHRASE_RULES / WORD_MAP.
# 'sq' được gộp về 'sqmm' để '10.0SQ' == '10.0 SQmm'.

COLOR_SET = {
    'black', 'red', 'blue', 'yellow', 'white',
    'green', 'orange', 'grey', 'brown', 'violet', 'pink', 'greenyellow',
}

ORIGIN_SET = {
    'vietnam', 'china', 'korea', 'japan', 'taiwan', 'germany',
    'italy', 'indonesia', 'turkey', 'malaysia',
    'india', 'usa', 'france', 'thailand', 'singapore',
}

# Từ "thừa" - chỉ mang tính mô tả, không phân biệt vật tư. Bỏ khi so túi từ
# để 'GREEN COLOR / LED TYPE' == 'GREEN / LED', 'T/B' == "T/B ASS'Y".
FILLER_WORDS = {
    'color', 'colour', 'type', 'the', 'of', 'for', 'with', 'and',
    'w',  # 'w/' = with
    'assembly',  # ASS'Y / ASSY / ASSEMBLY đã gộp về 'assembly' ở PHRASE_RULES
}

# Nhãn KHÔNG phải thông số đo dạng nhãn-giá trị (đơn vị tiết diện, đơn vị
# đếm...) - bỏ khỏi extract_measures để tránh lệch giả '10sq' vs '10 sqmm'.
MEASURE_LABEL_BLACKLIST = {
    'sq', 'sqmm', 'sqm', 'mm', 'cm', 'pi', 'ea', 'pcs', 'set', 'cl',
}

# Mảnh model thường là "nhiễu" do dính vào mã: tên hãng/xuất xứ/tiền tố -
# không tính khi xét hai model có khác gốc hay không.
MODEL_NOISE_FRAGMENTS = {
    'ac', 'dc', 'v', 'ys',
    'hyundai', 'hengzhu', 'china', 'korea', 'vietnam', 'japan',
    'se', 'tse', 'vina',
}

# SH CABLE / SH VINA CABLE và COSMOLINK (VINA) là CÙNG một NCC (đổi tên).
SUPPLIER_ALIASES = {
    'sh cable': 'cosmolink',
    'sh vina cable': 'cosmolink',
    'sh vina': 'cosmolink',
    'cosmolink vina': 'cosmolink',
    'cosmolink': 'cosmolink',
}

PHRASE_RULES = [
    (re.compile(r'\bflame[\s\-]+retardant\b'), 'fr'),
    (re.compile(r'\bh[\s\-]*pvc\b'), 'hpvc'),
    (re.compile(r'\bsh\s+vina\s+cable\b'), 'cosmolink'),
    (re.compile(r'\bsh\s+cable\b'), 'cosmolink'),
    (re.compile(r'\bcosmolink\s+vina\b'), 'cosmolink'),
    (re.compile(r'\bviet\s+nam\b'), 'vietnam'),
    (re.compile(r'\bsouth\s+korea\b'), 'korea'),
    # 'AUX.VOLTAGE' và 'COIL VOLTAGE' cùng nghĩa: điện áp cuộn dây rơ-le.
    (re.compile(r'\b(?:aux|coil)[\s\.\-]*voltage\b'), 'voltage'),
    # Chuẩn hóa mọi biến thể ASS'Y / ASSY / ASSEMBLY về một token rồi coi
    # là từ thừa (T/B == T/B ASS'Y).
    (re.compile(r"\bass(?:'?y|embly)?\b"), 'assembly'),
]

CLASS5_PATTERN = re.compile(r'\bclass\s*[-:=]?\s*5(?:\.0)?\b')
TOKEN_PATTERN = re.compile(r'\d+\.\d+|\d+|[a-z]+')
PARENTHETICAL_PATTERN = re.compile(r'\([^)]*\)')
DIGIT_PATTERN = re.compile(r'\d')

LABEL_BEFORE_VALUE = re.compile(r'\b([a-z]{1,3})(\d+(?:\.\d+)?)\b')
VALUE_BEFORE_LABEL = re.compile(r'\b(\d+(?:\.\d+)?)([a-z]{1,4})\b')
MODEL_CODE_PATTERN = re.compile(r'[a-z0-9]+(?:[\-/][a-z0-9]+)*')

COLOR_ORDER_SENSITIVE = True


def remove_diacritics(text):
    if not isinstance(text, str):
        return ''
    text = unicodedata.normalize('NFD', text)
    return (re.sub(r'[\u0300-\u036f]', '', text)
            .replace('đ', 'd').replace('Đ', 'D'))


def normalize_numbers(text):
    text = re.sub(r'(\d+)\.0+(?!\d)', r'\1', text)
    text = re.sub(r'(\d+\.\d*?[1-9])0+(?!\d)', r'\1', text)
    return text


VOLTAGE_PREFIX = re.compile(r'\b(ac|dc)[\s\-]*(\d+(?:[./]\d+)*)\s*v\b')
VOLTAGE_SUFFIX = re.compile(r'\b(\d+(?:[./]\d+)*)\s*v(ac|dc)\b')
VOLTAGE_PLAIN = re.compile(r'\b(\d+(?:[./]\d+)*)\s*v\b')


def normalize_voltage(text):
    """
    Chuẩn hóa điện áp về một dạng thống nhất '<số>v<ac|dc>' để:
        'DC110V' == '110VDC' == '110V DC'  -> '110vdc'
        'AC200/220V' == '200/220VAC'       -> '200/220vac'
        '240V'                              -> '240v'
    Nhờ đó cách viết prefix/suffix AC/DC không còn gây lệch giả.
    """
    text = VOLTAGE_PREFIX.sub(lambda m: f'{m.group(2)}v{m.group(1)}', text)
    text = VOLTAGE_SUFFIX.sub(lambda m: f'{m.group(1)}v{m.group(2)}', text)
    return text


def normalize_text(text):
    s = remove_diacritics(str(text).lower())
    for pattern, replacement in PHRASE_RULES:
        s = pattern.sub(replacement, s)
    s = CLASS5_PATTERN.sub(' ', s)
    s = normalize_voltage(s)
    return normalize_numbers(s)


def tokenize_field(text):
    """Tách token GIỮ NGUYÊN THỨ TỰ."""
    return [WORD_MAP.get(t, t) for t in TOKEN_PATTERN.findall(text)]


def tokenize_field_color_agnostic(text):
    """Sắp riêng token MÀU (giữ vị trí token khác) - dùng xét cấu trúc."""
    tokens = tokenize_field(text)
    pos = [i for i, t in enumerate(tokens) if t in COLOR_SET]
    if len(pos) > 1:
        for i, color in zip(pos, sorted(tokens[i] for i in pos)):
            tokens[i] = color
    return tokens


def extract_measures(text):
    """Bóc thông số dạng nhãn-giá trị, giữ ràng buộc (M10,L20 != M20,L10)."""
    measures = {}
    for label, value in LABEL_BEFORE_VALUE.findall(text):
        if label in MEASURE_LABEL_BLACKLIST:
            continue
        measures.setdefault(label, set()).add(value)
    for value, label in VALUE_BEFORE_LABEL.findall(text):
        if label in MEASURE_LABEL_BLACKLIST:
            continue
        measures.setdefault(label, set()).add(value)
    return {k: frozenset(v) for k, v in measures.items()}


# =============================================================================
# 3. CHỮ KÝ VẬT TƯ
# =============================================================================

@dataclass(frozen=True)
class ItemSignature:
    raw: str = ''
    category: tuple = ()
    fields: Counter = dc_field(default_factory=Counter)
    field_source: tuple = ()
    color_sequence: tuple = ()
    origins: tuple = ()
    suppliers: frozenset = frozenset()
    measures: tuple = ()
    model_codes: frozenset = frozenset()
    word_bag: Counter = dc_field(default_factory=Counter)
    body_word_bag: Counter = dc_field(default_factory=Counter)

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
    def colors(self):
        return frozenset(self.color_sequence)

    @property
    def category_text(self):
        return ' '.join(self.category)

    @property
    def body_fields(self):
        """Trường SAU loại vật tư (để so khi lệch item name)."""
        first = ' '.join(self.category)
        body = self.fields.copy()
        if first in body:
            body[first] -= 1
            if body[first] <= 0:
                del body[first]
        return body

    @property
    def technical_text(self):
        return ' | '.join(
            f'{label.upper()}={"/".join(sorted(values))}'
            for label, values in self.measures)

    @property
    def color_text(self):
        return '/'.join(self.color_sequence).upper()

    @property
    def origin_text(self):
        return '/'.join(self.origins).upper()


def _extract_suppliers(normalized):
    return frozenset(
        canonical for alias, canonical in SUPPLIER_ALIASES.items()
        if alias in normalized)


def _collapse_color_annotation(normalized_part):
    """
    Bỏ ngoặc chú giải MÀU trong một trường.

    'b w r bk (blue white red black)' -> 'b w r bk'
    Chỉ bỏ khi MỌI từ trong ngoặc đều là màu (kể cả dạng viết tắt map sang
    màu). Ngoặc chứa số đo/định danh/thông tin khác được GIỮ NGUYÊN.
    """
    def replace(match):
        inner = match.group(1)
        words = re.findall(r'[a-z]+', inner)
        if words and all(WORD_MAP.get(w, w) in COLOR_SET for w in words):
            return ' '
        return match.group(0)

    return re.sub(r'\(([^)]*)\)', replace, normalized_part)


# Cụm "sạch" để lấy màu: chỉ gồm chữ và dấu ngăn màu (/ + - & khoảng trắng),
# KHÔNG chứa chữ số. 'GREEN/YELLOW' hợp lệ; 'DL11-GY' bị loại vì có số '11'.
COLOR_CLUSTER_PATTERN = re.compile(r'[a-z]+(?:[\s/&+\-][a-z]+)*')


def _extract_colors_in_order(text):
    """
    Lấy màu theo đúng thứ tự xuất hiện, KHỬ TRÙNG LẶP.

    Bước 1: XÓA mọi mã model (cụm chữ-số dính nhau như 'ysbsl33-dl11-gy')
            khỏi text. Nhờ đó hậu tố màu nằm trong mã model ('-GY', '-RGY')
            biến mất, không bị nhận nhầm là màu 'grey'.
    Bước 2: bóc màu từ phần còn lại, giữ thứ tự.
    """
    cleaned = MODEL_CODE_PATTERN.sub(
        lambda mt: ' ' if (
            DIGIT_PATTERN.search(mt.group()) and re.search(r'[a-z]', mt.group())
        ) else mt.group(),
        text,
    )

    colors = [
        WORD_MAP.get(piece, piece)
        for piece in re.findall(r'[a-z]+', cleaned)
        if WORD_MAP.get(piece, piece) in COLOR_SET
    ]

    return tuple(dict.fromkeys(colors))


def build_signature(description):
    if description is None or pd.isna(description):
        return ItemSignature()

    normalized = normalize_text(description)
    raw_parts = str(description).split(',')
    normalized_parts = normalized.split(',')

    field_token_lists = []
    field_source = OrderedDict()

    for index, part in enumerate(normalized_parts):
        # Chuẩn hóa mã màu viết tắt kèm chú giải trong ngoặc: khi một cụm
        # có dạng 'B-W-R-BK(Blue-White-Red-Black)' thì phần trong ngoặc chỉ
        # viết đầy đủ lại các màu đã có ở dạng viết tắt -> bỏ ngoặc để hai
        # cách viết tương đương. Ngoặc KHÁC (mang thông tin mới, số đo, định
        # danh) vẫn được GIỮ để không làm mất dữ liệu phân biệt.
        part = _collapse_color_annotation(part)

        tokens = tokenize_field_color_agnostic(part)
        if not tokens:
            continue
        field_token_lists.append(tokens)
        key = ' '.join(tokens)
        raw_text = raw_parts[index].strip() if index < len(raw_parts) else key
        field_source.setdefault(key, raw_text or key)

    if not field_token_lists:
        return ItemSignature(raw=str(description))

    outside = normalize_text(PARENTHETICAL_PATTERN.sub(' ', str(description)))

    # Màu chỉ hợp lệ khi đứng như một TRƯỜNG MÀU ĐỘC LẬP (một token viết
    # tách bằng khoảng trắng / dấu phẩy / dấu gạch chéo giữa các màu), KHÔNG
    # phải mảnh dính trong mã model. Ví dụ 'DL11-GY': 'gy' là hậu tố model,
    # không phải màu 'grey'. Ta bóc màu từ các cụm KHÔNG chứa chữ số và
    # không dính liền chữ-số qua dấu gạch.
    color_sequence = _extract_colors_in_order(outside)
    outside_tokens = tokenize_field(outside)
    origins = tuple(
        dict.fromkeys(t for t in outside_tokens if t in ORIGIN_SET))

    # Túi từ toàn mô tả: gom mọi token của tất cả trường (đã color-agnostic
    # nên màu sắp cùng nhau), bỏ từ thừa. Đây là nền so sánh CHÍNH - không
    # phụ thuộc dấu phân cách hay cách gộp/tách trường.
    all_tokens = [t for field in field_token_lists for t in field]
    word_bag = Counter(t for t in all_tokens if t not in FILLER_WORDS)

    body_tokens = [t for field in field_token_lists[1:] for t in field]
    body_word_bag = Counter(t for t in body_tokens if t not in FILLER_WORDS)

    # Loại vật tư (category) cũng bỏ từ thừa: "T/B ASS'Y" -> category ('t','b')
    # trùng với "T/B", nên hai mã được coi CÙNG loại (khớp), không phải lệch
    # item name.
    category = tuple(t for t in field_token_lists[0] if t not in FILLER_WORDS)
    if not category:
        category = tuple(field_token_lists[0])

    return ItemSignature(
        raw=str(description),
        category=category,
        fields=Counter(' '.join(t) for t in field_token_lists),
        field_source=tuple(field_source.items()),
        color_sequence=color_sequence,
        origins=origins,
        suppliers=_extract_suppliers(normalized),
        measures=tuple(sorted(extract_measures(normalized).items())),
        model_codes=frozenset(
            t for t in MODEL_CODE_PATTERN.findall(normalized)
            if DIGIT_PATTERN.search(t) and re.search(r'[a-z]', t)),
        word_bag=word_bag,
        body_word_bag=body_word_bag,
    )


# =============================================================================
# 4. SO KHỚP HAI CHỮ KÝ
# =============================================================================

MATCH_EXACT = 'Khớp Description'
MATCH_CONTAIN = 'Khớp chứa nhau'
MATCH_ITEMNAME = 'Lệch Item Name nhưng khớp phần sau'

DIFF_CATEGORY = 'Lệch loại vật tư'
DIFF_SPEC = 'Lệch thông số'
DIFF_COLOR = 'Lệch màu sắc'
DIFF_ORIGIN = 'Lệch xuất xứ'
DIFF_SUPPLIER = 'Lệch nhà sản xuất'
DIFF_CONTENT = 'Lệch nội dung mô tả'
DIFF_COLOR_ORDER = 'Lệch thứ tự màu sắc'

MATCH_LABELS = {MATCH_EXACT, MATCH_CONTAIN, MATCH_ITEMNAME}


def _format_fields(field_counter, *source_maps):
    lookup = {}
    for source_map in source_maps:
        lookup.update(source_map)
    return ' | '.join(
        lookup.get(key, key) for key in sorted(field_counter.elements()))


def _compare_attributes(sig1, sig2):
    """Kiểm tra màu / xuất xứ / NCC / thông số. None nếu tương thích."""
    if sig1.colors and sig2.colors:
        if sig1.colors != sig2.colors:
            return DIFF_COLOR, f'{sig1.color_text} ↔ {sig2.color_text}'
        if COLOR_ORDER_SENSITIVE and sig1.color_sequence != sig2.color_sequence:
            return DIFF_COLOR_ORDER, f'{sig1.color_text} ↔ {sig2.color_text}'

    if sig1.suppliers and sig2.suppliers and sig1.suppliers != sig2.suppliers:
        return DIFF_SUPPLIER, (
            f'{"/".join(sorted(sig1.suppliers))} ↔ '
            f'{"/".join(sorted(sig2.suppliers))}')

    if sig1.origins and sig2.origins and set(sig1.origins) != set(sig2.origins):
        return DIFF_ORIGIN, f'{sig1.origin_text} ↔ {sig2.origin_text}'

    map1, map2 = sig1.measure_map, sig2.measure_map
    conflicts = sorted(
        f'{label.upper()}={"/".join(sorted(map1[label]))}'
        f'↔{"/".join(sorted(map2[label]))}'
        for label in map1.keys() & map2.keys()
        if map1[label] != map2[label])
    if conflicts:
        return DIFF_SPEC, '; '.join(conflicts)

    # Model: KHÔNG so chuỗi thô (dễ báo giả khi model bị dính tên hãng/xuất
    # xứ, thêm prefix nhà sản xuất, hay tách/gộp khác nhau). Thay vào đó tách
    # mỗi model thành tập mảnh chữ-số rồi xét quan hệ tập hợp: chỉ coi là
    # LỆCH khi mỗi bên có mảnh RIÊNG mà không bên nào là tập con bên kia
    # (tức hai định danh khác gốc thực sự). Các trường hợp một bên chỉ dài
    # hơn (dính hãng, prefix) sẽ là tập cha/con -> không lệch.
    frags1 = _model_fragments(sig1.model_codes)
    frags2 = _model_fragments(sig2.model_codes)
    if frags1 and frags2 and not (frags1 <= frags2 or frags2 <= frags1):
        only1 = frags1 - frags2
        only2 = frags2 - frags1
        only1 -= MODEL_NOISE_FRAGMENTS
        only2 -= MODEL_NOISE_FRAGMENTS
        # Bỏ các mảnh mà một bên chỉ là tiền tố/hậu tố chuỗi của mảnh bên
        # kia (vd 'dnc' ⊂ 'ysdnc' do thêm prefix nhà sản xuất) -> không lệch.
        only1, only2 = _drop_affix_matches(only1, only2)
        if only1 and only2:
            return DIFF_SPEC, (
                f'model {sorted(only1)} ↔ {sorted(only2)}')

    return None


def _drop_affix_matches(only1, only2):
    """Loại khỏi hai tập những mảnh chữ mà một bên là tiền tố/hậu tố chuỗi
    của một mảnh bên kia (vd 'dnc' vs 'ysdnc', 'mr' vs 'mrb')."""
    matched1, matched2 = set(), set()
    for a in only1:
        for b in only2:
            if a.isalpha() and b.isalpha() and (
                a.endswith(b) or b.endswith(a)
                or a.startswith(b) or b.startswith(a)
            ):
                matched1.add(a)
                matched2.add(b)
    return only1 - matched1, only2 - matched2


def _model_fragments(model_codes):
    """Tập mảnh chữ-số của toàn bộ model codes (bỏ mảnh 1 ký tự nhiễu)."""
    frags = set()
    for model in model_codes:
        for piece in re.findall(r'[a-z]+|\d+', model):
            if len(piece) >= 2 or piece.isdigit():
                frags.add(piece)
    return frags


def _format_words(word_counter):
    return ' '.join(sorted(word_counter.elements()))


def _measure_and_color_tokens(sig):
    """Các token đã được giải thích bởi measures/màu/xuất xứ - loại khỏi túi
    từ khi so quan hệ cấu trúc để không báo 'thừa' oan các con số đo hay
    tên màu (chúng đã được so riêng ở _compare_attributes).

    KHÔNG loại token của model code: một cụm như '1-PORT' có thể vừa bị nhận
    là model vừa chứa từ 'port' có nghĩa; loại nó sẽ gây lệch giả bất đối
    xứng khi bên kia viết 'port' tách rời."""
    tokens = set()
    for label, values in sig.measures:
        tokens.add(label)
        for v in values:
            tokens.update(re.findall(r'[a-z0-9]+', v))
    tokens.update(sig.color_sequence)
    tokens.update(sig.origins)
    return tokens


def _structural_relation_bag(bag1, bag2, ignore1, ignore2):
    """
    So QUAN HỆ CẤU TRÚC trên TẬP TỪ (set of words), sau khi bỏ các token đã
    được giải thích bởi thông số/màu/xuất xứ/model.

    Dùng TẬP HỢP (không đếm bội) vì trong mô tả kỹ thuật, một từ lặp lại
    (vd 'port' trong '1-PORT ... 2 port', 'comm' trong 'ethernet comm ...
    serial comm') không có nghĩa là "nhiều hơn". Nhờ so trên tập từ, mọi
    khác biệt về dấu phân cách và cách gộp/tách trường đều bị bỏ qua.

    Trả về ('exact'|'contain'|'content', chi_tiết, side_đầy_đủ_hơn).
    """
    s1 = {w for w in bag1 if w not in ignore1}
    s2 = {w for w in bag2 if w not in ignore2}

    extra1 = s1 - s2
    extra2 = s2 - s1

    if not extra1 and not extra2:
        return 'exact', '', 0
    if extra1 and extra2:
        return 'content', (
            f'[{" ".join(sorted(extra1))}] ↔ [{" ".join(sorted(extra2))}]'), 0
    if extra2:
        return 'contain', ' '.join(sorted(extra2)), 2
    return 'contain', ' '.join(sorted(extra1)), 1


def compare_signatures(sig1, sig2):
    def result(label, detail='', missing='', fuller=''):
        return {'label': label, 'detail': detail,
                'missing': missing, 'fuller': fuller}

    if sig1.is_empty or sig2.is_empty:
        return result(DIFF_CONTENT, 'Thiếu mô tả')

    attribute_diff = _compare_attributes(sig1, sig2)

    ignore1 = _measure_and_color_tokens(sig1)
    ignore2 = _measure_and_color_tokens(sig2)

    rel_all, detail_all, fuller_all = _structural_relation_bag(
        sig1.word_bag, sig2.word_bag, ignore1, ignore2)

    rel_body, detail_body, _ = _structural_relation_bag(
        sig1.body_word_bag, sig2.body_word_bag, ignore1, ignore2)

    same_category = sig1.category == sig2.category

    # --- A. CÙNG loại vật tư ---
    if same_category:
        if attribute_diff is not None:
            label, detail = attribute_diff
            return result(label, detail)
        if rel_all == 'exact':
            return result(MATCH_EXACT, 'Mọi trường khớp sau chuẩn hóa')
        if rel_all == 'contain':
            return result(
                MATCH_CONTAIN,
                f'Mã {fuller_all} chứa trọn mã kia; bên ngắn thiếu: '
                f'[{detail_all}]',
                missing=detail_all, fuller=str(fuller_all))
        return result(DIFF_CONTENT, detail_all)

    # --- B. KHÁC loại vật tư ---
    if attribute_diff is None and rel_body in ('exact', 'contain'):
        note = 'khớp Description' if rel_body == 'exact' else 'khớp chứa nhau'
        return result(
            MATCH_ITEMNAME,
            f'Item Name: {sig1.category_text} ↔ {sig2.category_text}; '
            f'phần sau {note}')

    detail = f'{sig1.category_text.upper()} ↔ {sig2.category_text.upper()}'
    if attribute_diff is not None:
        label, extra = attribute_diff
        return result(DIFF_CATEGORY, f'{detail}; kèm {label.lower()}: {extra}')
    return result(DIFF_CATEGORY, detail)


# =============================================================================
# 5. UNION-FIND
# =============================================================================

class UnionFind:
    def __init__(self):
        self._parent = {}

    def find(self, item):
        self._parent.setdefault(item, item)
        while self._parent[item] != item:
            self._parent[item] = self._parent[self._parent[item]]
            item = self._parent[item]
        return item

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self._parent[rb] = ra

    def groups(self):
        result = {}
        for item in self._parent:
            result.setdefault(self.find(item), []).append(item)
        return result


# =============================================================================
# PHẦN 2 + 3: SO SÁNH DESCRIPTION
# =============================================================================

def classify_cause(label):
    if label in MATCH_LABELS:
        return 'Khớp'
    if label == DIFF_SPEC:
        return 'Lệch thông số (nghi lỗi gõ / khác quy cách)'
    if label in (DIFF_COLOR, DIFF_COLOR_ORDER):
        return 'Lệch màu sắc'
    if label in (DIFF_ORIGIN, DIFF_SUPPLIER):
        return 'Lệch xuất xứ / nhà cung cấp'
    if label == DIFF_CATEGORY:
        return 'Lệch loại vật tư (nghi lỗi ánh xạ)'
    return 'Lệch nội dung (cần rà soát)'


def classify_row(desc_a, desc_target):
    res = compare_signatures(build_signature(desc_a), build_signature(desc_target))
    return res['label'], res['detail'], classify_cause(res['label'])


def _rest_after_first_comma(description):
    """Trả phần SAU dấu phẩy đầu tiên (bỏ item name), giữ nguyên phần còn lại."""
    parts = str(description).split(',', 1)
    return parts[1].strip() if len(parts) > 1 else ''


def classify_code_a_vs_ksys(desc_a, desc_ksys):
    """
    Quy tắc riêng cho Code A vs Ksys.

    Theo thiết kế, Description Ksys = Description Code A đã BỎ item name (phần
    trước dấu phẩy đầu tiên). Vì vậy ta so PHẦN SAU dấu phẩy đầu của Code A
    với TOÀN BỘ Ksys:
      - Nếu khớp / khớp chứa nhau  -> 'Khớp' (đúng thiết kế).
      - Nếu phần sau đó KHÔNG khớp Ksys -> mới coi là lệch, và nêu rõ nguyên
        nhân (item name, thông số, màu, ...).
    """
    if desc_ksys is None or pd.isna(desc_ksys) or str(desc_ksys).strip() == '':
        return 'Không có Ksys', '', 'Bỏ qua (thiếu Ksys)'

    rest = _rest_after_first_comma(desc_a)
    if not rest:
        return 'Không có phần sau', '', 'Bỏ qua (Code A không có phần sau)'

    res = compare_signatures(build_signature(rest), build_signature(desc_ksys))

    if res['label'] in MATCH_LABELS:
        return 'Khớp (đúng thiết kế)', res['detail'], 'Khớp'

    # Phần sau dấu phẩy đầu của Code A không khớp Ksys -> lệch thực sự.
    return res['label'], res['detail'], classify_cause(res['label'])


def run_matching(df, col_a, col_target, prefix):
    results = df.apply(
        lambda row: classify_row(row[col_a], row[col_target]), axis=1)
    df[f'Trạng thái Desc ({prefix})'] = [r[0] for r in results]
    df[f'Chi tiết Desc ({prefix})'] = [r[1] for r in results]
    df[f'Nhóm nguyên nhân ({prefix})'] = [r[2] for r in results]
    return df


def run_matching_code_a_ksys(df):
    results = df.apply(
        lambda row: classify_code_a_vs_ksys(
            row['Description code A'], row['Description ksys']), axis=1)
    df['Trạng thái Desc (Code A vs Ksys)'] = [r[0] for r in results]
    df['Chi tiết Desc (Code A vs Ksys)'] = [r[1] for r in results]
    df['Nhóm nguyên nhân (Code A vs Ksys)'] = [r[2] for r in results]
    return df


# =============================================================================
# PHẦN 4: TRÙNG LẶP HỆ MÃ NEW
# =============================================================================

def build_new_item_table(df):
    columns = ['item new', 'Description new', 'item code A',
               'Description code A', 'item ksys']
    table = df[columns].copy()
    for column in ['item new', 'item code A', 'item ksys']:
        table[column] = table[column].fillna('').astype(str).str.strip()

    is_real = table['item new'].apply(lambda v: not is_placeholder_code(v))
    has_desc = (table['Description new'].notna()
                & table['Description new'].astype(str).str.strip().ne(''))

    return (table[is_real & has_desc]
            .drop_duplicates(keep='first').reset_index(drop=True))


def join_unique_values(series):
    values = {str(v).strip() for v in series if pd.notna(v) and str(v).strip()}
    return ' | '.join(sorted(values))


def aggregate_by_new_item(new_item_table):
    aggregated = (
        new_item_table
        .groupby(['item new', 'Description new'], as_index=False)
        .agg({'item code A': join_unique_values,
              'Description code A': join_unique_values,
              'item ksys': join_unique_values}))

    aggregated['signature'] = aggregated['Description new'].map(build_signature)
    aggregated = aggregated[
        aggregated['signature'].map(lambda s: not s.is_empty)
    ].reset_index(drop=True)

    aggregated['Loại vật tư'] = aggregated['signature'].map(lambda s: s.category_text)
    aggregated['Thông số kỹ thuật'] = aggregated['signature'].map(lambda s: s.technical_text)
    aggregated['Màu sắc'] = aggregated['signature'].map(lambda s: s.color_text)
    aggregated['Xuất xứ'] = aggregated['signature'].map(lambda s: s.origin_text)
    return aggregated


def block_keys(signature):
    """
    Sinh NHIỀU khóa chặn cho một mô tả (inverted index).

    Một cặp chỉ cần chung ÍT NHẤT một khóa là được đưa ra so khớp. Nhờ vậy
    bản "khớp chứa nhau" dù có thừa vài thông số/ghi chú vẫn gặp được bản
    ngắn (chúng vẫn chung các thông số cốt lõi). Đây là kỹ thuật chuẩn
    trong entity resolution để không bỏ sót cặp mà vẫn tránh so tất cả với
    tất cả.

    Khóa gồm:
      - mỗi cặp thông số nhãn-giá trị (vd 'm=5', 'l=30')
      - mỗi mã model (vd 'ht-05', 'yspl3')
      - nếu không có cả hai: loại vật tư + màu (cho nhóm chỉ khác ghi chú)
    """
    keys = set()

    for label, values in signature.measures:
        for value in values:
            keys.add(f'M:{label}={value}')

    for model in signature.model_codes:
        keys.add(f'K:{model}')

    # Lưới an toàn: LUÔN thêm khóa theo loại vật tư + màu. Nhờ vậy bản mô tả
    # ngắn (không có thông số/model nào để index) vẫn gặp được bản đầy đủ
    # cùng loại - đúng tinh thần "khớp chứa nhau".
    color = '/'.join(signature.color_sequence)
    keys.add(f'N:{signature.category_text}#{color}')

    return keys


PAIR_COLUMNS = [
    'Nhóm trùng số', 'Phân loại',
    'item new 1', 'Description new 1',
    'item code A 1', 'Description code A 1', 'item ksys 1',
    'item new 2', 'Description new 2',
    'item code A 2', 'Description code A 2', 'item ksys 2',
    'Thông số kỹ thuật', 'Màu sắc', 'Xuất xứ',
    'Chi tiết', 'Mã có mô tả đầy đủ hơn', 'Dữ kiện còn thiếu',
]

GROUP_COLUMNS = [
    'Nhóm trùng số', 'Phân loại nhóm', 'item new', 'Loại vật tư',
    'Description new', 'item code A', 'Description code A', 'item ksys',
    'Thông số kỹ thuật', 'Màu sắc', 'Xuất xứ', 'UOM new',
]

LABEL_STRENGTH = {MATCH_EXACT: 3, MATCH_CONTAIN: 2, MATCH_ITEMNAME: 1}


def find_duplicate_pairs(aggregated):
    """
    Ghép cặp Item New <-> Item New qua inverted-index blocking, so từng cặp
    ứng viên đúng một lần (khử trùng lặp bằng tập đã xét).
    """
    rows = aggregated.to_dict('records')

    # inverted index: khóa -> danh sách chỉ số dòng
    index = {}
    for idx, row in enumerate(rows):
        for key in block_keys(row['signature']):
            index.setdefault(key, []).append(idx)

    pairs = []
    near_misses = []
    evaluated = set()

    for candidates in index.values():
        if len(candidates) < 2:
            continue

        for a in range(len(candidates)):
            for b in range(a + 1, len(candidates)):
                i, j = candidates[a], candidates[b]
                key = (i, j) if i < j else (j, i)
                if key in evaluated:
                    continue
                evaluated.add(key)

                row1, row2 = rows[key[0]], rows[key[1]]
                if row1['item new'] == row2['item new']:
                    continue

                res = compare_signatures(row1['signature'], row2['signature'])
                first, second = sorted((row1, row2), key=lambda r: r['item new'])

                if res['fuller'] == '1':
                    fuller_item = row1['item new']
                elif res['fuller'] == '2':
                    fuller_item = row2['item new']
                else:
                    fuller_item = ''

                if res['label'] not in MATCH_LABELS:
                    if res['label'] == DIFF_COLOR_ORDER:
                        near_misses.append({
                            'detail': res['detail'],
                            'first': first, 'second': second})
                    continue

                pairs.append({
                    'label': res['label'], 'detail': res['detail'],
                    'missing': res['missing'], 'fuller': fuller_item,
                    'first': first, 'second': second})

    # Sắp xếp để kết quả TẤT ĐỊNH giữa các lần chạy (index dùng set).
    pairs.sort(key=lambda p: (p['first']['item new'], p['second']['item new']))
    near_misses.sort(
        key=lambda p: (p['first']['item new'], p['second']['item new']))

    return pairs, near_misses


def build_duplicate_reports(pairs, uom_by_item=None):
    uom_by_item = uom_by_item or {}
    if not pairs:
        return (pd.DataFrame(columns=GROUP_COLUMNS),
                pd.DataFrame(columns=PAIR_COLUMNS))

    uf = UnionFind()
    for pair in pairs:
        uf.union(pair['first']['item new'], pair['second']['item new'])

    group_id = {}
    for number, (_, members) in enumerate(
            sorted(uf.groups().items(), key=lambda kv: min(kv[1])), start=1):
        for member in members:
            group_id[member] = number

    group_label = {}
    for pair in pairs:
        gid = group_id[pair['first']['item new']]
        if (gid not in group_label
                or LABEL_STRENGTH[pair['label']] < LABEL_STRENGTH[group_label[gid]]):
            group_label[gid] = pair['label']

    pair_records = []
    for pair in pairs:
        first, second = pair['first'], pair['second']
        pair_records.append({
            'Nhóm trùng số': group_id[first['item new']],
            'Phân loại': pair['label'],
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
            'Chi tiết': pair['detail'],
            'Mã có mô tả đầy đủ hơn': pair['fuller'],
            'Dữ kiện còn thiếu': pair['missing']})

    df_pairs = (pd.DataFrame(pair_records, columns=PAIR_COLUMNS)
                .drop_duplicates(keep='first')
                .sort_values(['Nhóm trùng số', 'item new 1', 'item new 2'])
                .reset_index(drop=True))

    member_rows = {}
    for pair in pairs:
        for side in ('first', 'second'):
            row = pair[side]
            member_rows[row['item new']] = row

    df_groups = pd.DataFrame([
        {'Nhóm trùng số': group_id[item_new],
         'Phân loại nhóm': group_label[group_id[item_new]],
         'item new': item_new,
         'Loại vật tư': row['Loại vật tư'],
         'Description new': row['Description new'],
         'item code A': row['item code A'],
         'Description code A': row['Description code A'],
         'item ksys': row['item ksys'],
         'Thông số kỹ thuật': row['Thông số kỹ thuật'],
         'Màu sắc': row['Màu sắc'],
         'Xuất xứ': row['Xuất xứ'],
         'UOM new': uom_by_item.get(item_new, '')}
        for item_new, row in member_rows.items()
    ], columns=GROUP_COLUMNS)

    df_groups = (df_groups.sort_values(['Nhóm trùng số', 'item new'])
                 .reset_index(drop=True))
    return df_groups, df_pairs


MULTI_CODE_A_SUMMARY_COLUMNS = [
    'item new', 'Số lượng Item Code A', 'Danh sách Item Code A',
    'Danh sách Description New', 'Danh sách Item Ksys']

CODE_A_VERDICT_COLUMNS = [
    'item new', 'Description new', 'item code A 1', 'Description code A 1',
    'item code A 2', 'Description code A 2', 'Kết luận', 'Chi tiết']


def find_items_with_multiple_code_a(new_item_table):
    relation = new_item_table[new_item_table['item code A'].ne('')]
    count = (relation
             .drop_duplicates(subset=['item new', 'item code A'], keep='first')
             .groupby('item new')['item code A'].nunique())
    flagged = count[count > 1].index

    df_detail = (relation[relation['item new'].isin(flagged)]
                 .drop_duplicates(keep='first')
                 [['item new', 'Description new', 'item code A',
                   'Description code A', 'item ksys']]
                 .sort_values(['item new', 'item code A'])
                 .reset_index(drop=True))

    if df_detail.empty:
        return pd.DataFrame(columns=MULTI_CODE_A_SUMMARY_COLUMNS), df_detail

    df_summary = (df_detail.groupby('item new', as_index=False)
                  .agg(**{'Số lượng Item Code A': ('item code A', 'nunique'),
                          'Danh sách Item Code A': ('item code A', join_unique_values),
                          'Danh sách Description New': ('Description new', join_unique_values),
                          'Danh sách Item Ksys': ('item ksys', join_unique_values)})
                  .sort_values(['Số lượng Item Code A', 'item new'],
                               ascending=[False, True])
                  .reset_index(drop=True))
    return df_summary, df_detail


def check_one_new_many_desc(df_multi_code_a_detail):
    records = []
    for item_new, block in df_multi_code_a_detail.groupby('item new'):
        rows = (block.drop_duplicates(subset=['item code A'], keep='first')
                .to_dict('records'))
        for i in range(len(rows)):
            for j in range(i + 1, len(rows)):
                first, second = sorted((rows[i], rows[j]),
                                       key=lambda r: r['item code A'])
                res = compare_signatures(
                    build_signature(first['Description code A']),
                    build_signature(second['Description code A']))
                is_ok = res['label'] in MATCH_LABELS
                records.append({
                    'item new': item_new,
                    'Description new': first['Description new'],
                    'item code A 1': first['item code A'],
                    'Description code A 1': first['Description code A'],
                    'item code A 2': second['item code A'],
                    'Description code A 2': second['Description code A'],
                    'Kết luận': (f'Gộp hợp lý ({res["label"]})' if is_ok
                                 else f'LỖI ÁNH XẠ - {res["label"]}'),
                    'Chi tiết': res['detail']})

    return (pd.DataFrame(records, columns=CODE_A_VERDICT_COLUMNS)
            .sort_values(['Kết luận', 'item new', 'item code A 1'])
            .reset_index(drop=True))


REPEATED_NAME_COLUMNS = [
    'Hệ mã', 'item', 'Item Name (trước phẩy)', 'Cụm lặp (sau phẩy)',
    'Description']


def _first_two_parts(description):
    parts = str(description).split(',')
    head = parts[0].strip() if parts else ''
    nxt = parts[1].strip() if len(parts) > 1 else ''
    return head, nxt


def _norm_compact(text):
    return re.sub(r'[^a-z0-9]', '', remove_diacritics(str(text).lower()))


def find_repeated_item_name(new_item_table):
    """Item mà TÊN VẬT TƯ lặp: trước dấu phẩy đầu == cụm ngay sau."""
    records = []
    seen = set()

    def scan(desc, code, label):
        if not str(code).strip() or is_placeholder_code(code):
            return
        head, nxt = _first_two_parts(desc)
        if not head or not nxt:
            return
        nh = _norm_compact(head)
        if nh and nh == _norm_compact(nxt):
            key = (label, str(code), str(desc))
            if key in seen:
                return
            seen.add(key)
            records.append({
                'Hệ mã': label, 'item': str(code),
                'Item Name (trước phẩy)': head,
                'Cụm lặp (sau phẩy)': nxt, 'Description': desc})

    for _, row in new_item_table.iterrows():
        scan(row['Description new'], row['item new'], 'Code New')
        scan(row['Description code A'], row['item code A'], 'Code A')

    return (pd.DataFrame(records, columns=REPEATED_NAME_COLUMNS)
            .sort_values(['Hệ mã', 'Item Name (trước phẩy)', 'item'])
            .reset_index(drop=True))


# =============================================================================
# PHẦN 5: CẢNH BÁO DỮ LIỆU
# =============================================================================

def build_data_quality_report(df_full, df_dropped, new_item_table,
                              df_groups=None, near_misses=None):
    records = []

    def add(loai, ma, mo_ta):
        records.append({'Loại cảnh báo': loai, 'Mã': ma, 'Chi tiết': mo_ta})

    for _, row in df_dropped.iterrows():
        add('Thiếu Description New (không kiểm tra được)',
            str(row['item code A']),
            f"Code A: {row['Description code A']} | item new: {row['item new']}")

    placeholder = df_full[df_full['item new'].apply(is_placeholder_code)]
    for _, row in placeholder.iterrows():
        add('Item New là mã placeholder / N/A', str(row['item code A']),
            f"item new = '{row['item new']}' | {row['Description code A']}")

    desc_count = new_item_table.groupby('item new')['Description new'].nunique()
    for item_new in desc_count[desc_count > 1].index:
        add('Item New có nhiều Description New khác nhau', item_new,
            join_unique_values(new_item_table.loc[
                new_item_table['item new'] == item_new, 'Description new']))

    code_a_map = (new_item_table[new_item_table['item code A'].ne('')]
                  .groupby('item code A')['item new'].nunique())
    for code_a in code_a_map[code_a_map > 1].index:
        add('Một Item Code A ánh xạ nhiều Item New', code_a,
            join_unique_values(new_item_table.loc[
                new_item_table['item code A'] == code_a, 'item new']))

    ksys_map = (new_item_table[new_item_table['item ksys'].ne('')]
                .groupby('item new')['item ksys'].nunique())
    for item_new in ksys_map[ksys_map > 1].index:
        add('Item New ứng với nhiều Item Ksys', item_new,
            join_unique_values(new_item_table.loc[
                new_item_table['item new'] == item_new, 'item ksys']))

    if df_groups is not None and not df_groups.empty:
        for gid, blk in df_groups.groupby('Nhóm trùng số'):
            if blk['UOM new'].nunique() > 1:
                add('Nhóm nghi trùng nhưng lệch UOM', f'Nhóm {gid}',
                    ' | '.join(f"{r['item new']}={r['UOM new']}"
                               for _, r in blk.iterrows()))

    for item in (near_misses or []):
        add('Loại sát nút - chỉ chênh thứ tự màu',
            f"{item['first']['item new']} vs {item['second']['item new']}",
            f"{item['detail']} || {item['first']['Description new']} || "
            f"{item['second']['Description new']}")

    return pd.DataFrame(records, columns=['Loại cảnh báo', 'Mã', 'Chi tiết'])


# =============================================================================
# ĐIỀU PHỐI
# =============================================================================

def main():
    df_full = load_dataset()
    df, df_dropped = drop_invalid_description_new(df_full)

    df = compare_uom(df)
    df = run_matching_code_a_ksys(df)
    df = run_matching(df, 'Description code A', 'Description new', 'Code A vs New')

    df_uom_diff_ksys = df[~df['uom_ksys_match']].copy()
    df_uom_diff_new = df[~df['uom_new_match']].copy()
    df_desc_diff_new = (
        df[~df['Trạng thái Desc (Code A vs New)'].isin(MATCH_LABELS)]
        .drop_duplicates(keep='first').reset_index(drop=True))

    # Code A vs Ksys: chỉ giữ các dòng LỆCH THỰC SỰ (bỏ 'Khớp' và các dòng
    # thiếu Ksys). Đây là kiểm tra xác nhận hệ Ksys có đúng thiết kế không.
    ksys_ok_labels = {'Khớp (đúng thiết kế)', 'Không có Ksys',
                      'Không có phần sau'}
    df_desc_diff_ksys = (
        df[~df['Trạng thái Desc (Code A vs Ksys)'].isin(ksys_ok_labels)]
        [['item code A', 'item ksys', 'item new',
          'Description code A', 'Description ksys',
          'Trạng thái Desc (Code A vs Ksys)',
          'Chi tiết Desc (Code A vs Ksys)',
          'Nhóm nguyên nhân (Code A vs Ksys)']]
        .drop_duplicates(keep='first').reset_index(drop=True))

    new_item_table = build_new_item_table(df)
    aggregated = aggregate_by_new_item(new_item_table)

    uom_by_item = (
        df[df['item new'].apply(lambda v: not is_placeholder_code(v))]
        .groupby('item new')['uom new'].apply(join_unique_values).to_dict())

    pairs, near_misses = find_duplicate_pairs(aggregated)
    df_dup_groups, df_dup_pairs = build_duplicate_reports(pairs, uom_by_item)

    df_multi_code_a_summary, df_multi_code_a_detail = (
        find_items_with_multiple_code_a(new_item_table))
    df_one_new_many_desc = check_one_new_many_desc(df_multi_code_a_detail)
    df_repeated_name = find_repeated_item_name(new_item_table)

    df_data_quality = build_data_quality_report(
        df_full, df_dropped, new_item_table,
        df_groups=df_dup_groups, near_misses=near_misses)

    sheets = {
        'Lệch Desc (Code A vs New)': df_desc_diff_new,
        'Lệch Desc (Code A vs Ksys)': df_desc_diff_ksys,
        'Lệch UOM (Code A vs Ksys)': df_uom_diff_ksys,
        'Lệch UOM (Code A vs New)': df_uom_diff_new,
        'Toàn bộ Data Checked': df,
        'Item New trùng Description': df_dup_groups,
        'Chi tiết Item New trùng': df_dup_pairs,
        'Item New nhiều Code A': df_multi_code_a_summary,
        'Lỗi 1 item new nhiều desc': df_one_new_many_desc,
        'Item bị lặp Item Name': df_repeated_name,
        'Cảnh báo dữ liệu': df_data_quality,
    }

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"--> Xuất thành công: '{OUTPUT_FILE}'")
    print(f"    + Lệch UOM  : Ksys ({len(df_uom_diff_ksys)}) "
          f"| New ({len(df_uom_diff_new)})")
    print(f"    + Lệch Desc (Code A vs New): {len(df_desc_diff_new)} dòng")
    print(f"    + Lệch Desc (Code A vs Ksys): {len(df_desc_diff_ksys)} dòng "
          f"(còn lại khớp đúng thiết kế)")
    print("      Phân bố nhóm nguyên nhân:")
    for cause, cnt in df['Nhóm nguyên nhân (Code A vs New)'].value_counts().items():
        print(f"        - {cause}: {cnt}")
    print(f"    + [4A] Item New nghi trùng: "
          f"{df_dup_groups['Nhóm trùng số'].nunique()} nhóm / "
          f"{df_dup_groups['item new'].nunique()} mã / {len(df_dup_pairs)} cặp")
    if not df_dup_pairs.empty:
        for label, cnt in df_dup_pairs['Phân loại'].value_counts().items():
            print(f"        - {label}: {cnt} cặp")
    print(f"    + [4B] Item New nhiều Code A: {len(df_multi_code_a_summary)} mã")
    if not df_one_new_many_desc.empty:
        for verdict, cnt in (df_one_new_many_desc['Kết luận']
                             .str.split(' - ').str[0].value_counts().items()):
            print(f"        - {verdict}: {cnt} cặp")
    print(f"    + [4C] Item lặp Item Name: {len(df_repeated_name)} dòng")
    print(f"    + [5]  Cảnh báo dữ liệu   : {len(df_data_quality)} dòng")
    return locals()


if __name__ == '__main__':
    main()