# -*- coding: utf-8 -*-
"""
=============================================================================
 KIỂM TRA ĐỐI CHIẾU DANH MỤC VẬT TƯ: CODE A  <->  KSYS  <->  CODE NEW
=============================================================================
 Phần 1 : So sánh UOM
 Phần 2 : So sánh Description (dùng chung engine chữ ký của Phần 4)
 Phần 3 : Xuất báo cáo lệch UOM / Description
 Phần 4 : Kiểm tra trùng lặp trong hệ mã New
 Phần 5 : Cảnh báo chất lượng dữ liệu.
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

INPUT_FILE = 'CODE A check _Ms Hoa new 3.xlsx'
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
    # Bỏ dòng có TRƯỜNG CUỐI (sau dấu phẩy cuối) là 'vr' - dòng biến thể,
    # không cần đối chiếu.
    ends_with_vr = cleaned.str.split(',').str[-1].str.strip().eq('vr')
    keep = has_value & ~(is_invalid_value | has_no_letter | ends_with_vr)
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
    'bk': 'black', 'bl': 'black',
    'br': 'brown',
    'r': 'red',
    'og': 'orange', 'org': 'orange', 'orn': 'orange',
    'ye': 'yellow', 'y': 'yellow', 'yl': 'yellow',
    'gn': 'green', 
    'b': 'blue',
    'vio': 'violet', 'pur': 'violet',
    'gray': 'grey', 'gr': 'grey',
    'w': 'white',
    'pk': 'pink', 'pnk': 'pink',
    # Xuất xứ
    'vnm': 'vietnam', 'cn': 'china', 'chn': 'china',
    'kr': 'korea', 'kor': 'korea', 'jp': 'japan', 'jpn': 'japan',
    'tw': 'taiwan', 'twn': 'taiwan', 'deu': 'germany', 'vina': 'vietnam',
    # Thông số (mm² viết nhiều kiểu -> 1 dạng chuẩn)
    'sq': 'sqmm', 'sqmm': 'sqmm', 'mmsq': 'sqmm', 'sqm': 'sqmm',
    'fr': 'flameretardant', 'hpvc': 'hpvc',
}

COLOR_SET = {
    'black', 'red', 'blue', 'yellow', 'white',
    'green', 'orange', 'grey', 'brown', 'violet', 'pink', 'greenyellow',
}

ORIGIN_SET = {
    'vietnam', 'china', 'korea', 'japan', 'taiwan', 'germany',
    'italy', 'indonesia', 'turkey', 'malaysia',
    'india', 'usa', 'france', 'thailand', 'singapore',
}

FILLER_WORDS = {
    'color', 'colour', 'type', 'the', 'of', 'for', 'with', 'and',
    'w',  
    'assembly',  
}

MEASURE_LABEL_BLACKLIST = {
    'mm', 'cm', 'pi', 'ea', 'pcs', 'set', 'cl',
}

MODEL_NOISE_FRAGMENTS = {
    'ac', 'dc', 'v', 'ys',
    'hyundai', 'hengzhu', 'china', 'korea', 'vietnam', 'japan',
    'se', 'tse', 'vina',
}

TYPE_DISCRIMINATOR_WORDS = {
    'bolt', 'screw', 'nut', 'washer', 'rivet', 'stud',
    'socket', 'head', 'cap', 'countersunk', 'pan',
    'combined', 'tapping',
    'hex', 'hexagon', 'slotted', 'phillips',
}

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
    (re.compile(r'\b(?:aux|coil)[\s\.\-]*voltage\b'), 'voltage'),
    (re.compile(r"\bass(?:'?y|embly)?\b"), 'assembly'),
]

CLASS5_PATTERN = re.compile(r'\bclass\s*[-:=]?\s*5(?:\.0)?\b')
TOKEN_PATTERN = re.compile(r'\d+\.\d+|\d+|[a-z]+')
PARENTHETICAL_PATTERN = re.compile(r'\([^)]*\)')
DIGIT_PATTERN = re.compile(r'\d')
LABELED_FIELD = re.compile(r'([a-z][a-z ]{1,20}):\s*([^,()]+)')
LABEL_BEFORE_VALUE = re.compile(r'\b([a-z]{1,3})\s?(\d+(?:\.\d+)?)\b')
VALUE_BEFORE_LABEL = re.compile(r'\b(\d+(?:\.\d+)?)\s?([a-z]{1,4})\b')
CHANNEL_RATIO = re.compile(r'\b([a-z])/(\d+)\b')
RATIO_UNIT = re.compile(r'\b(\d+(?:\.\d+)?/\d+(?:\.\d+)?)([a-z]{1,3})\b')
MODEL_CODE_PATTERN = re.compile(r'[a-z0-9]+(?:[\-/][a-z0-9]+)*')


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


def normalize_voltage(text):
    """
    Chuẩn hóa điện áp về dạng TÁCH RỜI '<số>v <ac|dc>' để mọi cách viết
    """
    # số + V + (AC|DC) dính  ->  'số v ac [dc]'
    text = re.sub(
        r'\b(\d+(?:[./\-]\d+)*)\s*v\s*(ac|dc)(?:\s*/\s*(ac|dc))?\b',
        lambda m: f'{m.group(1)}v {m.group(2)}'
                  + (f' {m.group(3)}' if m.group(3) else ''),
        text)
    # (AC|DC) prefix + số + V  ->  'số v ac [dc]'
    text = re.sub(
        r'\b(ac|dc)(?:\s*/\s*(ac|dc))?[\s\-]*(\d+(?:[./\-]\d+)*)\s*v\b',
        lambda m: f'{m.group(3)}v {m.group(1)}'
                  + (f' {m.group(2)}' if m.group(2) else ''),
        text)
    return text


def normalize_text(text):
    s = remove_diacritics(str(text).lower())
    for pattern, replacement in PHRASE_RULES:
        s = pattern.sub(replacement, s)
    s = CLASS5_PATTERN.sub(' ', s)
    s = normalize_voltage(s)
    # Chuẩn hóa tiết diện: '10sq', '10 sq', '10sqmm', '10 mm2' -> '10 sqmm'
    s = re.sub(r'(\d)\s*(?:sq\s*mm|sqmm|mm\s*sq|sq|mm2)\b', r'\1 sqmm', s)
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
    for channel, count in CHANNEL_RATIO.findall(text):
        measures.setdefault(f'ch_{channel}', set()).add(count)

    for ratio, unit in RATIO_UNIT.findall(text):
        measures.setdefault(f'ratio_{unit}', set()).add(ratio)
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
    color_clusters: frozenset = frozenset()
    color_singles: frozenset = frozenset()
    origins: tuple = ()
    suppliers: frozenset = frozenset()
    measures: tuple = ()
    model_codes: frozenset = frozenset()
    word_bag: Counter = dc_field(default_factory=Counter)
    body_word_bag: Counter = dc_field(default_factory=Counter)
    labeled_fields: tuple = ()
    type_words: frozenset = frozenset()
    numbers: tuple = ()

    @property
    def is_empty(self):
        return not self.fields

    @property
    def measure_map(self):
        return dict(self.measures)

    @property
    def labeled_map(self):
        return dict(self.labeled_fields)

    @property
    def source_map(self):
        return dict(self.field_source)

    @property
    def colors(self):
        result = set(self.color_sequence) | set(self.color_singles)
        for cluster in self.color_clusters:
            result.update(cluster)
        return frozenset(result)

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
    GIỮ nguyên toàn bộ nội dung trong ngoặc.
    """
    return normalized_part


# Cụm "sạch" để lấy màu: chỉ gồm chữ và dấu ngăn màu (/ + - & khoảng trắng),
# KHÔNG chứa chữ số. 'GREEN/YELLOW' hợp lệ; 'DL11-GY' bị loại vì có số '11'.
COLOR_CLUSTER_PATTERN = re.compile(r'[a-z]+(?:[\s/&+\-][a-z]+)*')


def _extract_color_info(text):
    """
    Bóc màu từ cả nội dung ngoài ngoặc lẫn trong ngoặc.
    """
    cleaned = MODEL_CODE_PATTERN.sub(
        lambda mt: ' ' if (
            DIGIT_PATTERN.search(mt.group()) and re.search(r'[a-z]', mt.group())
        ) else mt.group(),
        text,
    )

    chain_abbrev = {
        'b': 'blue', 'w': 'white', 'r': 'red', 'y': 'yellow',
        'bk': 'black', 'blk': 'black', 'bn': 'brown', 'br': 'brown',
        'g': 'green', 'gn': 'green', 'gy': 'grey', 'gr': 'grey',
        'o': 'orange', 'v': 'violet',
    }

    clusters = []
    singles = []
    all_colors = []

    for run in COLOR_CLUSTER_PATTERN.findall(cleaned):
        has_strong_separator = bool(re.search(r'[/&+\-]', run))
        pieces = re.split(r'[\s/&+\-]+', run)
        if not pieces:
            continue

        # Chỉ trong chuỗi màu có dấu ngăn mới mở rộng B/W/R/Y một ký tự.
        mapped = [
            (chain_abbrev.get(p, WORD_MAP.get(p, p))
             if has_strong_separator else WORD_MAP.get(p, p))
            for p in pieces
        ]
        seq = [c for c in mapped if c in COLOR_SET]
        if not seq:
            continue

        # Một chuỗi có dấu ngăn chỉ là cụm màu nếu TẤT CẢ mảnh đều nhận diện
        # được là màu. Điều này ngăn 'MEANWELL/CHINA' thành cụm màu giả.
        if has_strong_separator and len(pieces) >= 2:
            if len(seq) != len(pieces):
                # Chuỗi có một phần không phải màu (T/B, W/PVC, AC/DC...)
                # không phải dữ kiện màu; bỏ cả run để tránh nhận nhầm B/W.
                continue
            cluster = tuple(dict.fromkeys(seq))
            if len(cluster) >= 2:
                clusters.append(cluster)
            else:
                singles.extend(cluster)
            all_colors.extend(cluster)
            continue

        # Không có dấu ngăn mạnh: màu cách khoảng trắng là các màu đơn.
        singles.extend(seq)
        all_colors.extend(seq)

    return (
        frozenset(clusters),
        frozenset(singles),
        tuple(dict.fromkeys(all_colors)),
    )


def _canon_number(text):
    """Chuẩn hóa số để so: '6.0'->'6', '1.50'->'1.5'."""
    if '.' in text:
        text = text.rstrip('0').rstrip('.')
    return text or '0'


def build_signature(description):
    if description is None or pd.isna(description):
        return ItemSignature()

    normalized = normalize_text(description)
    raw_parts = str(description).split(',')
    normalized_parts = normalized.split(',')

    field_token_lists = []
    field_source = OrderedDict()

    for index, part in enumerate(normalized_parts):
        # Giữ nội dung trong ngoặc như một phần dữ liệu cần so.
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

    # Màu được bóc từ TOÀN BỘ mô tả, gồm cả nội dung trong ngoặc. 
    color_clusters, color_singles, color_sequence = _extract_color_info(
        normalized)
    outside = normalized

    outside_tokens = tokenize_field(outside)
    origins = tuple(
        dict.fromkeys(t for t in outside_tokens if t in ORIGIN_SET))

    # So sánh màu sắc, hai hay ba lần lặp lại sẽ tính là chú giải
    all_tokens = [t for field in field_token_lists for t in field]
    word_bag = Counter(t for t in all_tokens if t not in FILLER_WORDS)

    body_tokens = [t for field in field_token_lists[1:] for t in field]
    body_word_bag = Counter(t for t in body_tokens if t not in FILLER_WORDS)

    # Trường có nhãn 'Nhãn: giá trị': bóc thành ràng buộc 
    # 'Words: English and Vietnamese' KHÁC 'Words: English'.
    labeled = {}
    for label, value in LABELED_FIELD.findall(normalize_text(str(description))):
        key = ' '.join(label.split())
        value_words = frozenset(
            w for w in re.findall(r'[a-z0-9]+', value)
            if w not in FILLER_WORDS)
        if value_words:
            labeled[key] = value_words
    labeled_fields = tuple(sorted(labeled.items()))

    # "Từ định loại": trước và sau dấu phẩy đầu tiên
    if len(field_token_lists) >= 2:
        type_words = frozenset(
            t for t in field_token_lists[1] if t not in FILLER_WORDS)
    else:
        type_words = frozenset()

    # Loại vật tư (category) cũng bỏ từ thừa
    category = tuple(t for t in field_token_lists[0] if t not in FILLER_WORDS)
    if not category:
        category = tuple(field_token_lists[0])

    # Bội số (multiset) TẤT CẢ con số: giữ số lần xuất hiện để phân biệt
    numbers = tuple(sorted(
        _canon_number(n) for n in re.findall(r'\d+(?:\.\d+)?', normalized)))

    return ItemSignature(
        raw=str(description),
        category=category,
        fields=Counter(' '.join(t) for t in field_token_lists),
        field_source=tuple(field_source.items()),
        color_sequence=color_sequence,
        color_clusters=color_clusters,
        color_singles=color_singles,
        origins=origins,
        suppliers=_extract_suppliers(normalized),
        measures=tuple(sorted(extract_measures(normalized).items())),
        model_codes=frozenset(
            t for t in MODEL_CODE_PATTERN.findall(normalized)
            if DIGIT_PATTERN.search(t) and re.search(r'[a-z]', t)),
        word_bag=word_bag,
        body_word_bag=body_word_bag,
        labeled_fields=labeled_fields,
        type_words=type_words,
        numbers=numbers,
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
    # --- Màu sắc ---
    # Chỉ khóa THỨ TỰ trong phạm vi một CỤM MÀU; màu đơn lẻ, cả cụm di chuyển tự do
    colors1 = sig1.colors
    colors2 = sig2.colors
    # Chỉ so màu khi CẢ HAI bên đều lộ màu.
    both_have_colors = bool(colors1) and bool(colors2)
    if both_have_colors:
        if colors1 != colors2:
            return DIFF_COLOR, f'{sig1.color_text} ↔ {sig2.color_text}'
        # (b) chỉ là ràng buộc THỨ TỰ màu trước sau, không khóa vị trí trong hàng
        def preserves_cluster_order(cluster, observed_sequence):
            positions = {color: i for i, color in enumerate(observed_sequence)}
            return all(
                positions[cluster[i]] < positions[cluster[i + 1]]
                for i in range(len(cluster) - 1)
            )

        all_constraints = sig1.color_clusters | sig2.color_clusters
        sequence1 = sig1.color_sequence or tuple(sorted(colors1))
        sequence2 = sig2.color_sequence or tuple(sorted(colors2))
        invalid = [
            cluster for cluster in all_constraints
            if not (preserves_cluster_order(cluster, sequence1)
                    and preserves_cluster_order(cluster, sequence2))
        ]
        if invalid:
            def show(sequence):
                return '/'.join(sequence).upper()
            return DIFF_COLOR_ORDER, (
                f'{show(sequence1)} ↔ {show(sequence2)}')

    if sig1.suppliers and sig2.suppliers and sig1.suppliers != sig2.suppliers:
        return DIFF_SUPPLIER, (
            f'{"/".join(sorted(sig1.suppliers))} ↔ '
            f'{"/".join(sorted(sig2.suppliers))}')

    if sig1.origins and sig2.origins and set(sig1.origins) != set(sig2.origins):
        return DIFF_ORIGIN, f'{sig1.origin_text} ↔ {sig2.origin_text}'

    # Trường có nhãn 'Nhãn: giá trị'
    lab1, lab2 = sig1.labeled_map, sig2.labeled_map
    label_conflicts = sorted(
        f'{key}: {" ".join(sorted(lab1[key]))} ↔ {" ".join(sorted(lab2[key]))}'
        for key in lab1.keys() & lab2.keys()
        if lab1[key] != lab2[key])
    if label_conflicts:
        return DIFF_SPEC, '; '.join(label_conflicts)

    map1, map2 = sig1.measure_map, sig2.measure_map
    conflicts = sorted(
        f'{label.upper()}={"/".join(sorted(map1[label]))}'
        f'↔{"/".join(sorted(map2[label]))}'
        for label in map1.keys() & map2.keys()
        if map1[label] != map2[label])
    if conflicts:
        return DIFF_SPEC, '; '.join(conflicts)

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
    từ khi so quan hệ cấu trúc. CHỈ loại GIÁ TRỊ SỐ của measure"""
    tokens = set()
    for label, values in sig.measures:
        for v in values:
            # chỉ bỏ mảnh SỐ trong giá trị; giữ mọi mảnh chữ
            for frag in re.findall(r'[a-z0-9]+', v):
                if frag.isdigit() or re.search(r'\d', frag):
                    tokens.add(frag)
    tokens.update(sig.color_sequence)
    tokens.update(sig.origins)
    return tokens


def _pair_off_variants(extra1, extra2):
    """
    Loại khỏi hai tập những từ là BIẾN THỂ CHUỖI của nhau:
    """
    matched1, matched2 = set(), set()
    for a in sorted(extra1):
        for b in sorted(extra2):
            if b in matched2:
                continue
            long, short = (a, b) if len(a) >= len(b) else (b, a)
            if len(short) < 2:
                continue
            is_variant = (
                short in long                      # chứa chuỗi con / affix
                or (len(a) >= 3 and len(b) >= 3
                    and _one_edit_apart(a, b))     # lỗi gõ 1 ký tự
            )
            if is_variant:
                matched1.add(a)
                matched2.add(b)
                break
    return extra1 - matched1, extra2 - matched2


def _one_edit_apart(a, b):
    """True nếu a và b chỉ khác nhau 1 phép sửa (thêm/bớt/thay 1 ký tự)."""
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) == len(b):
        diffs = sum(1 for x, y in zip(a, b) if x != y)
        return diffs == 1
    # khác 1 độ dài: kiểm tra chèn/xóa 1 ký tự
    long, short = (a, b) if len(a) > len(b) else (b, a)
    i = j = 0
    skipped = False
    while i < len(long) and j < len(short):
        if long[i] != short[j]:
            if skipped:
                return False
            skipped = True
            i += 1
        else:
            i += 1
            j += 1
    return True


def _words_only(bag, ignore):
    """Chỉ giữ TỪ CHỮ (bỏ token chứa số và nhãn đo 1 ký tự). Số được so
    riêng qua tập số toàn mô tả."""
    result = set()
    for w in bag:
        if w in ignore:
            continue
        if DIGIT_PATTERN.search(w):
            continue
        if len(w) == 1 and w.isalpha():
            continue
        result.add(w)
    return result


def _structural_relation_bag(bag1, bag2, ignore1, ignore2):
    """
    So QUAN HỆ CẤU TRÚC trên TẬP TỪ CHỮ (số được so RIÊNG ở phần thông số).
    """
    s1 = _words_only(bag1, ignore1)
    s2 = _words_only(bag2, ignore2)

    extra1 = s1 - s2
    extra2 = s2 - s1

    # Ghép cặp các từ riêng là BIẾN THỂ CHUỖI của nhau 
    extra1, extra2 = _pair_off_variants(extra1, extra2)

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

    # Kiểm tra "từ định loại" (2 trường đầu): nếu MỖI bên có từ định loại
    # riêng mà không bên nào là tập con -> khác LOẠI CHI TIẾT.
    type_diff = None
    tw1, tw2 = sig1.type_words, sig2.type_words
    if tw1 and tw2 and tw1 != tw2:
        only_t1 = {w for w in tw1 - tw2 if not DIGIT_PATTERN.search(w)}
        only_t2 = {w for w in tw2 - tw1 if not DIGIT_PATTERN.search(w)}
        # Chỉ coi là KHÁC LOẠI khi phần khác biệt của F2 chứa DANH TỪ ĐỊNH LOẠI.
        disc1 = only_t1 & TYPE_DISCRIMINATOR_WORDS
        disc2 = only_t2 & TYPE_DISCRIMINATOR_WORDS
        if disc1 or disc2:
            left = ' '.join(sorted(only_t1)) or '(không)'
            right = ' '.join(sorted(only_t2)) or '(không)'
            type_diff = f'{left} ↔ {right}'

    ignore1 = _measure_and_color_tokens(sig1)
    ignore2 = _measure_and_color_tokens(sig2)

    # So THÔNG SỐ theo BỘI SỐ (multiset) các con số (nguyên tắc hai đầu):
    #   - Mỗi bên có số RIÊNG (cả hai chiều khác rỗng) -> LỆCH THÔNG SỐ.
    #   - Chỉ một bên thừa số -> góp vào quan hệ "chứa nhau".
    num1, num2 = Counter(sig1.numbers), Counter(sig2.numbers)
    num_only1 = sorted((num1 - num2).elements())
    num_only2 = sorted((num2 - num1).elements())
    number_diff = bool(num_only1) and bool(num_only2)
    number_extra = None
    if num_only1 and not num_only2:
        number_extra = 1
    elif num_only2 and not num_only1:
        number_extra = 2

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
        # Khác loại chi tiết (2 trường đầu mỗi bên có từ riêng) -> lệch loại.
        if type_diff is not None:
            return result(DIFF_CATEGORY, f'Khác loại chi tiết: {type_diff}')
        # Mỗi bên có số RIÊNG -> lệch thông số .
        if number_diff:
            return result(
                DIFF_SPEC,
                f'số {num_only1} ↔ {num_only2}')
        if rel_all == 'exact':
            # Từ chữ khớp hết; nếu một bên thừa số -> chứa nhau, không thì khớp.
            if number_extra:
                extra = num_only1 if number_extra == 1 else num_only2
                return result(
                    MATCH_CONTAIN,
                    f'Mã {number_extra} có thêm số {extra}',
                    missing=' '.join(extra), fuller=str(number_extra))
            return result(MATCH_EXACT, 'Mọi trường khớp sau chuẩn hóa')
        if rel_all == 'contain':
            return result(
                MATCH_CONTAIN,
                f'Mã {fuller_all} chứa trọn mã kia; bên ngắn thiếu: '
                f'[{detail_all}]',
                missing=detail_all, fuller=str(fuller_all))
        return result(DIFF_CONTENT, detail_all)

    # --- B. KHÁC loại vật tư ---
    # Nếu item name của hai bên đều là DANH TỪ ĐỊNH LOẠI (bolt/nut/screw...)
    # thì đây là hai sản phẩm KHÁC HẲN.
    cat1_disc = set(sig1.category) & TYPE_DISCRIMINATOR_WORDS
    cat2_disc = set(sig2.category) & TYPE_DISCRIMINATOR_WORDS
    both_are_types = bool(cat1_disc) and bool(cat2_disc) and (
        set(sig1.category) != set(sig2.category))

    if (not both_are_types and not number_diff and attribute_diff is None
            and rel_body in ('exact', 'contain')):
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


VR_SUFFIX = re.compile(r'\bvr\b\s*$', re.IGNORECASE)


def _has_vr_suffix(description):
    """Mô tả có đuôi 'VR' (biến thể/variant) - bỏ qua khi báo lệch loại vật tư."""
    return bool(VR_SUFFIX.search(str(description).strip()))


def classify_row(desc_a, desc_target):
    res = compare_signatures(build_signature(desc_a), build_signature(desc_target))
    label = res['label']

    # Lệch loại vật tư: kiểm tra kỹ. Nếu một trong hai mô tả có đuôi 'VR'(dòng biến thể), KHÔNG báo lệch loại 
    if label == DIFF_CATEGORY and (
            _has_vr_suffix(desc_a) or _has_vr_suffix(desc_target)):
        return (DIFF_CONTENT,
                res['detail'] + ' [bỏ qua lệch loại: dòng đuôi VR]',
                'Lệch nội dung (cần rà soát)')

    return label, res['detail'], classify_cause(label)


def _rest_after_first_comma(description):
    """Trả phần SAU dấu phẩy đầu tiên (bỏ item name), giữ nguyên phần còn lại."""
    parts = str(description).split(',', 1)
    return parts[1].strip() if len(parts) > 1 else ''


def classify_code_a_vs_ksys(desc_a, desc_ksys):
    """
    Quy tắc riêng cho Code A vs Ksys.

    Theo thiết kế, Description Ksys = Description Code A đã BỎ item name (phần
    trước dấu phẩy đầu tiên). Vì vậy ta so PHẦN SAU dấu phẩy đầu của Code A
    với TOÀ  BỘ Ksys:
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

    aggregated['Loại vật tư'] = aggregated['signature'].map(
        lambda s: s.category_text)
    aggregated['Thông số kỹ thuật'] = aggregated['signature'].map(
        lambda s: s.technical_text)
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
      - nếu không có cả hai: loại vật tư + màu
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
                  .agg(**{
                      'Số lượng Item Code A': ('item code A', 'nunique'),
                      'Danh sách Item Code A': (
                          'item code A', join_unique_values),
                      'Danh sách Description New': (
                          'Description new', join_unique_values),
                      'Danh sách Item Ksys': ('item ksys', join_unique_values)})
                  .sort_values(['Số lượng Item Code A', 'item new'],
                               ascending=[False, True])
                  .reset_index(drop=True))
    return df_summary, df_detail


# =============================================================================
# KIỂM TRA QUAN HỆ 1–1: ITEM NEW <-> DESCRIPTION NEW
# =============================================================================

DUPLICATE_DESCRIPTION_COLUMNS = [
    'Nhóm Description New', 'Description New (khóa tra cứu)',
    'Phân loại khớp', 'Số lượng Item New',
    'Danh sách Item New trùng', 'Description New khớp/chứa'
]


def build_itemnew_duplicate_description_report(pairs):
    """
    Sheet 'Item New trùng Description'.

    KHÓA NHÓM là Description New
    - Một Description New (hoặc một nhóm Description New khớp/chứa nhau)
      được dùng bởi từ hai Item New trở lên là lỗi 1 Description -> nhiều Item.
    - Chỉ lưu các nhóm có ít nhất hai Item New KHÁC NHAU..
 Với case 'khớp chứa', hai Description raw không giống hệt nhau. Khi đó
    chọn mô tả ngắn/gọn nhất làm 'khóa tra cứu' và vẫn liệt kê đầy đủ tất cả
    Description New trong cùng nhóm để người dùng truy vết.
    """
    if not pairs:
        return pd.DataFrame(columns=DUPLICATE_DESCRIPTION_COLUMNS)

    # Node là raw Description New. Union các description được kết luận khớp.
    uf = UnionFind()
    desc_to_items = {}
    desc_to_labels = {}

    for pair in pairs:
        d1 = str(pair['first']['Description new']).strip()
        d2 = str(pair['second']['Description new']).strip()
        i1 = str(pair['first']['item new']).strip()
        i2 = str(pair['second']['item new']).strip()
        if not d1 or not d2 or not i1 or not i2:
            continue
        uf.union(d1, d2)
        desc_to_items.setdefault(d1, set()).add(i1)
        desc_to_items.setdefault(d2, set()).add(i2)
        desc_to_labels.setdefault((d1, d2), set()).add(pair['label'])
        desc_to_labels.setdefault((d2, d1), set()).add(pair['label'])

    records = []
    groups = uf.groups()
    # Chỉ các node thật sự nằm trong cặp khớp mới có mặt trong `groups`.
    ordered_groups = sorted(
        (sorted(members) for members in groups.values()),
        key=lambda members: (min(len(d) for d in members), min(members)))

    group_no = 0
    for descriptions in ordered_groups:
        item_news = set()
        labels = set()
        for description in descriptions:
            item_news.update(desc_to_items.get(description, set()))
            for other in descriptions:
                if description != other:
                    labels.update(desc_to_labels.get((description, other), set()))

        # Đây là sheet kiểm tra 1 Description -> nhiều Item New.
        if len(item_news) < 2:
            continue

        group_no += 1
        lookup_key = min(descriptions, key=lambda d: (len(d), d))
        records.append({
            'Nhóm Description New': group_no,
            'Description New (khóa tra cứu)': lookup_key,
            'Phân loại khớp': ' | '.join(sorted(labels)),
            'Số lượng Item New': len(item_news),
            'Danh sách Item New trùng': ' | '.join(sorted(item_news)),
            'Description New khớp/chứa': ' | '.join(sorted(descriptions)),
        })

    return (pd.DataFrame(records, columns=DUPLICATE_DESCRIPTION_COLUMNS)
            .sort_values(['Nhóm Description New',
                          'Description New (khóa tra cứu)'])
            .reset_index(drop=True))


ONE_NEW_MANY_DESC_COLUMNS = [
    'item new (khóa tra cứu)', 'Số lượng Description New',
    'Danh sách Description New', 'Kết quả kiểm tra nội dung',
    'Cặp Description xung đột', 'Chi tiết'
]


def check_one_new_many_desc(new_item_table):
    """
    Sheet 'Lỗi 1 item new nhiều desc'.

    KHÓA NHÓM là Item New:
    - Truy vấn tất cả Description New gán cho đúng một Item New.
    - Chỉ đưa vào sheet khi có từ hai Description New raw khác nhau.
    - So sánh các Description New trong từng Item New để xác định có xung
      đột nội dung thực sự hay chỉ khác cách viết/khớp chứa..
    """
    records = []
    table = (new_item_table[['item new', 'Description new']]
             .drop_duplicates()
             .copy())

    for item_new, block in table.groupby('item new', sort=True):
        descs = sorted({
            str(d).strip() for d in block['Description new'].tolist()
            if str(d).strip()
        })
        if len(descs) < 2:
            continue

        conflicts = []
        matched_pairs = 0
        details = []
        for i in range(len(descs)):
            for j in range(i + 1, len(descs)):
                d1, d2 = descs[i], descs[j]
                res = compare_signatures(build_signature(d1), build_signature(d2))
                pair_name = f'[{d1}] ↔ [{d2}]'
                if res['label'] in MATCH_LABELS:
                    matched_pairs += 1
                else:
                    conflicts.append(pair_name)
                    details.append(
                        f'{pair_name}: {res["label"]}'
                        + (f' ({res["detail"]})' if res['detail'] else ''))

        if conflicts:
            outcome = 'LỖI - Xung đột nội dung Description New'
            conflict_text = ' | '.join(conflicts)
            detail_text = ' | '.join(details)
        else:
            # Raw Description khác nhau vẫn phải được ghi nhận để kiểm tra
            # quy tắc 1 Item New chỉ có một Description New.
            outcome = ('Cảnh báo - Nhiều Description New nhưng '
                       'nội dung khớp/khớp chứa')
            conflict_text = ''
            detail_text = (f'{matched_pairs} cặp Description New đều khớp '
                           'hoặc khớp chứa.')

        records.append({
            'item new (khóa tra cứu)': str(item_new),
            'Số lượng Description New': len(descs),
            'Danh sách Description New': ' | '.join(descs),
            'Kết quả kiểm tra nội dung': outcome,
            'Cặp Description xung đột': conflict_text,
            'Chi tiết': detail_text,
        })

    return (pd.DataFrame(records, columns=ONE_NEW_MANY_DESC_COLUMNS)
            .sort_values(['Kết quả kiểm tra nội dung',
                          'item new (khóa tra cứu)'])
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
                              df_groups=None, df_pairs=None):
    records = []

    def add(loai, ma, mo_ta):
        records.append({'Loại cảnh báo': loai, 'Mã': ma, 'Chi tiết': mo_ta})

    # Item New có nhiều Description New khác nhau (nghi 1 mã gán nhiều nghĩa).
    desc_count = new_item_table.groupby('item new')['Description new'].nunique()
    for item_new in desc_count[desc_count > 1].index:
        add('Item New có nhiều Description New khác nhau', item_new,
            join_unique_values(new_item_table.loc[
                new_item_table['item new'] == item_new, 'Description new']))

    # Một Item Code A ánh xạ nhiều Item New (nghi tách mã sai).
    code_a_map = (new_item_table[new_item_table['item code A'].ne('')]
                  .groupby('item code A')['item new'].nunique())
    for code_a in code_a_map[code_a_map > 1].index:
        add('Một Item Code A ánh xạ nhiều Item New', code_a,
            join_unique_values(new_item_table.loc[
                new_item_table['item code A'] == code_a, 'item new']))

    # Nhóm các Item New KHỚP DESCRIPTION với nhau nhưng lệch UOM. Chỉ xét các
    # cặp đã được kết luận khớp (df_pairs), không phải mọi thành viên nhóm.
    if df_pairs is not None and not df_pairs.empty and df_groups is not None:
        uom_of = dict(zip(df_groups['item new'], df_groups['UOM new']))
        seen = set()
        for _, pair in df_pairs.iterrows():
            i1, i2 = pair['item new 1'], pair['item new 2']
            u1, u2 = uom_of.get(i1, ''), uom_of.get(i2, '')
            key = tuple(sorted((i1, i2)))
            if u1 and u2 and u1 != u2 and key not in seen:
                seen.add(key)
                add('Item New khớp Description nhưng lệch UOM',
                    f'{i1} vs {i2}', f'{i1}={u1} | {i2}={u2}')

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

    # Hai kiểm tra 1–1 độc lập, chỉ dựa trên Item New và Description New:
    # (1) Description New làm khóa -> tìm nhiều Item New dùng chung/khớp chứa.
    # (2) Item New làm khóa -> kiểm tra các Description New gán cho mã đó.
    df_itemnew_duplicate_desc = build_itemnew_duplicate_description_report(pairs)
    df_one_new_many_desc = check_one_new_many_desc(new_item_table)

    df_repeated_name = find_repeated_item_name(new_item_table)

    df_data_quality = build_data_quality_report(
        df_full, df_dropped, new_item_table,
        df_groups=df_dup_groups, df_pairs=df_dup_pairs)

    sheets = {
        'Lệch Desc (Code A vs New)': df_desc_diff_new,
        'Lệch Desc (Code A vs Ksys)': df_desc_diff_ksys,
        'Lệch UOM (Code A vs Ksys)': df_uom_diff_ksys,
        'Lệch UOM (Code A vs New)': df_uom_diff_new,
        'Toàn bộ Data Checked': df,
        'Item New trùng Description': df_itemnew_duplicate_desc,
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
    print(f"    + [4A] Description New trùng/khớp chứa dùng bởi nhiều Item New: "
          f"{len(df_itemnew_duplicate_desc)} nhóm / "
          f"{len(df_dup_pairs)} cặp đối chiếu")
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