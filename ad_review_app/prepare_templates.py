"""
광고물 심의 검토서 템플릿 준비 스크립트
기존 HWPX 파일에서 플레이스홀더 템플릿을 생성합니다.
앱 배포 전 개발자가 한 번 실행하는 스크립트입니다.
"""

import zipfile
import re
import os
import sys
from io import BytesIO
from xml.etree import ElementTree as ET

# ─────────────────────────────────────────────
# 경로 설정 (명령줄 인수 또는 기본값 사용)
# 사용법:
#   python prepare_templates.py [원본폴더경로]
#   원본폴더: 소심의 전(신규) 와 연장 검토서 폴더가 있는 상위 디렉터리
#
# 예시 (Windows):
#   python prepare_templates.py "C:\Users\user\Desktop\원본문서"
# 예시 (Linux):
#   python3 prepare_templates.py "/home/jinkui/dev/ad-review-extension"
# ─────────────────────────────────────────────

if len(sys.argv) >= 2:
    # 명령줄 인수로 원본 폴더 경로 지정
    BASE_DIR = sys.argv[1]
else:
    # 기본값: 스크립트 위치 기준 상위 폴더
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SHINYU_DIR   = os.path.join(BASE_DIR, '소심의 전(신규)')
YEONJANG_DIR = os.path.join(BASE_DIR, '연장 검토서')

# 출력 폴더: 스크립트와 같은 디렉터리의 templates/
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates')

os.makedirs(OUT_DIR, exist_ok=True)

# 경로 존재 확인
if not os.path.isdir(SHINYU_DIR):
    print(f'[오류] "소심의 전(신규)" 폴더를 찾을 수 없습니다.')
    print(f'       찾은 경로: {SHINYU_DIR}')
    print(f'       사용법: python prepare_templates.py "원본문서가_있는_폴더"')
    sys.exit(1)
if not os.path.isdir(YEONJANG_DIR):
    print(f'[오류] "연장 검토서" 폴더를 찾을 수 없습니다.')
    print(f'       찾은 경로: {YEONJANG_DIR}')
    sys.exit(1)

print(f'원본 폴더: {BASE_DIR}')
print(f'출력 폴더: {OUT_DIR}')
print()


# ─────────────────────────────────────────────
# HWPX 파일 I/O 유틸리티
# ─────────────────────────────────────────────

def read_hwpx(path):
    files = {}
    infos = {}
    with zipfile.ZipFile(path, 'r') as zf:
        for info in zf.infolist():
            files[info.filename] = zf.read(info.filename)
            infos[info.filename] = info
    return files, infos


# ─────────────────────────────────────────────
# 헤더 스타일 병합 유틸리티
# ─────────────────────────────────────────────

def _extract_style_elem(header_xml, tag, id_val):
    """header_xml 에서 <hh:TAG id="id_val"...>...</hh:TAG> 요소 추출"""
    pattern = f'<hh:{tag} id="{id_val}"'
    start = header_xml.find(pattern)
    if start == -1:
        return None
    gt_pos = header_xml.find('>', start)
    if header_xml[gt_pos - 1] == '/':          # 자기 닫힘 태그
        return header_xml[start:gt_pos + 1]
    close_tag = f'</hh:{tag}>'
    close_pos = header_xml.find(close_tag, gt_pos)
    if close_pos == -1:
        return None
    return header_xml[start:close_pos + len(close_tag)]


def _inject_into_container(header_xml, container_tag, new_elems):
    """컨테이너 닫힘 태그 앞에 요소를 삽입하고 itemCnt 업데이트"""
    if not new_elems:
        return header_xml
    close = f'</{container_tag}>'
    close_pos = header_xml.rfind(close)
    if close_pos == -1:
        return header_xml
    result = header_xml[:close_pos] + ''.join(new_elems) + header_xml[close_pos:]
    # itemCnt 업데이트
    result = re.sub(
        rf'(<{container_tag}[^>]+itemCnt=")(\d+)(")',
        lambda m: m.group(1) + str(int(m.group(2)) + len(new_elems)) + m.group(3),
        result, count=1
    )
    return result


def merge_missing_styles(base_header, src_header, table_xml):
    """table_xml 이 참조하지만 base_header 에 없는 스타일 정의를 src_header 에서 가져와 삽입"""
    result = base_header
    for attr, xml_tag, container_tag in [
        ('borderFillIDRef', 'borderFill', 'hh:borderFills'),
        ('charPrIDRef',     'charPr',     'hh:charProperties'),
        ('paraPrIDRef',     'paraPr',     'hh:paraProperties'),
    ]:
        refs     = set(re.findall(rf'{attr}="(\d+)"', table_xml))
        existing = set(re.findall(rf'<hh:{xml_tag} id="(\d+)"', result))
        missing  = sorted(refs - existing, key=int)
        new_elems = []
        for id_val in missing:
            elem = _extract_style_elem(src_header, xml_tag, id_val)
            if elem:
                new_elems.append(elem)
        result = _inject_into_container(result, container_tag, new_elems)
    return result


def merge_styles_with_remap(base_header, src_header, table_xml):
    """src 표가 참조하는 스타일을 base_header에 안전하게 이식한다.

    base와 src가 같은 ID를 서로 다른 정의로 재사용하는 경우가 많아서,
    src 쪽 스타일을 새 ID로 재매핑한 뒤 table_xml과 header.xml을 함께 갱신한다.
    """
    result_header = base_header
    result_table = table_xml
    remap_by_attr = {}

    for attr, xml_tag, container_tag in [
        ('borderFillIDRef', 'borderFill', 'hh:borderFills'),
        ('charPrIDRef',     'charPr',     'hh:charProperties'),
        ('paraPrIDRef',     'paraPr',     'hh:paraProperties'),
    ]:
        refs = sorted(set(re.findall(rf'{attr}="(\d+)"', result_table)), key=int)
        existing_ids = set(re.findall(rf'<hh:{xml_tag} id="(\d+)"', result_header))
        max_id = max(map(int, existing_ids)) if existing_ids else -1
        id_map = {}
        new_elems = []

        for id_val in refs:
            src_elem = _extract_style_elem(src_header, xml_tag, id_val)
            if not src_elem:
                continue

            base_elem = _extract_style_elem(result_header, xml_tag, id_val)
            if base_elem == src_elem:
                continue

            max_id += 1
            new_id = str(max_id)
            id_map[id_val] = new_id

            remapped_elem = src_elem.replace(f'id="{id_val}"', f'id="{new_id}"', 1)

            if xml_tag == 'charPr':
                border_map = remap_by_attr.get('borderFillIDRef', {})
                for old_ref, new_ref in border_map.items():
                    remapped_elem = remapped_elem.replace(
                        f'borderFillIDRef="{old_ref}"',
                        f'borderFillIDRef="{new_ref}"'
                    )

            new_elems.append(remapped_elem)

        for old_id, new_id in id_map.items():
            result_table = result_table.replace(
                f'{attr}="{old_id}"',
                f'{attr}="{new_id}"'
            )

        remap_by_attr[attr] = id_map
        result_header = _inject_into_container(result_header, container_tag, new_elems)

    return result_header, result_table


def write_hwpx(files, infos, path):
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w') as zout:
        for name, data in files.items():
            if name in infos:
                src = infos[name]
                zi = zipfile.ZipInfo(name, src.date_time)
                zi.compress_type = src.compress_type
                zi.create_system = src.create_system
                zi.create_version = src.create_version
                zi.extract_version = src.extract_version
                zi.external_attr = src.external_attr
                zi.internal_attr = src.internal_attr
            else:
                zi = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.create_system = 11  # NTFS (Hancom 기본값)
            zi.flag_bits = 0
            zi.extra = b''
            zi.comment = b''
            zout.writestr(zi, data)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())


def validate_hwpx_structure(path):
    with zipfile.ZipFile(path, 'r') as zf:
        names = zf.namelist()
        required_entries = [
            'mimetype',
            'version.xml',
            'Contents/header.xml',
            'Contents/section0.xml',
            'Contents/content.hpf',
            'META-INF/container.xml',
            'META-INF/manifest.xml',
        ]
        missing_entries = [name for name in required_entries if name not in names]
        if missing_entries:
            raise ValueError(f'{path}: HWPX 필수 항목 누락 - {missing_entries}')
        if names[:1] != ['mimetype']:
            raise ValueError(f'{path}: mimetype 엔트리가 ZIP 첫 항목이 아님')
        if zf.getinfo('mimetype').compress_type != zipfile.ZIP_STORED:
            raise ValueError(f'{path}: mimetype 엔트리가 비압축이 아님')
        header_xml = zf.read('Contents/header.xml').decode('utf-8')
        section_xml = zf.read('Contents/section0.xml').decode('utf-8')
        ET.fromstring(header_xml)
        ET.fromstring(section_xml)

    # 일부 실제 HWPX는 스타일 참조가 header.xml 외부 구조와 연결되어 있어
    # 정규식 기반 ID 검사는 오탐이 발생한다. 여기서는 ZIP/XML 구조만 검증한다.


def get_xml(files):
    return files['Contents/section0.xml'].decode('utf-8')


def set_xml(files, xml):
    files = dict(files)
    files['Contents/section0.xml'] = xml.encode('utf-8')
    return files


# ─────────────────────────────────────────────
# XML 조작 유틸리티
# ─────────────────────────────────────────────

def find_nth_table(xml, n=0):
    """n번째 hp:tbl 요소의 (시작, 끝) 위치 반환"""
    count = 0
    pos = 0
    while True:
        idx = xml.find('<hp:tbl ', pos)
        if idx == -1:
            return None
        if count == n:
            depth = 0
            p = idx
            while p < len(xml):
                if xml[p:p+8] == '<hp:tbl ':
                    depth += 1
                    end = xml.find('>', p)
                    p = end + 1 if end != -1 else p + 8
                elif xml[p:p+9] == '</hp:tbl>':
                    depth -= 1
                    p += 9
                    if depth == 0:
                        return (idx, p)
                else:
                    p += 1
            return None
        count += 1
        pos = xml.find('>', idx) + 1


def remove_row_containing(table_xml, text):
    """text를 포함하는 <hp:tr>...</hp:tr> 행 제거 후 rowCnt 업데이트"""
    result = []
    pos = 0
    while True:
        tr_start = table_xml.find('<hp:tr>', pos)
        if tr_start == -1:
            result.append(table_xml[pos:])
            break
        result.append(table_xml[pos:tr_start])
        tr_end = table_xml.find('</hp:tr>', tr_start) + len('</hp:tr>')
        row_content = table_xml[tr_start:tr_end]
        if text not in row_content:
            result.append(row_content)
        pos = tr_end

    result_str = ''.join(result)
    # rowCnt 업데이트
    total_rows = result_str.count('<hp:tr>')
    result_str = re.sub(r'rowCnt="\d+"', f'rowCnt="{total_rows}"', result_str, count=1)
    return result_str


def replace_tags(xml, replacements):
    """<hp:t> 태그 내용 치환"""
    for old_val, new_val in replacements.items():
        old_tag = f'<hp:t>{old_val}</hp:t>'
        new_tag = f'<hp:t>{new_val}</hp:t>'
        xml = xml.replace(old_tag, new_tag)
    return xml


def replace_tag_occurrences(xml, old_val, new_values):
    """같은 텍스트가 여러 번 나올 때 출현 순서대로 치환"""
    old_tag = f'<hp:t>{old_val}</hp:t>'
    for new_val in new_values:
        xml = xml.replace(old_tag, f'<hp:t>{new_val}</hp:t>', 1)
    return xml


def replace_first_in_table(xml, table_idx, old_val, new_val):
    """특정 테이블 내에서 첫 번째 출현만 치환 (수량 '1' 처리용)"""
    pos = find_nth_table(xml, table_idx)
    if pos is None:
        return xml
    t_xml = xml[pos[0]:pos[1]]
    old_tag = f'<hp:t>{old_val}</hp:t>'
    new_tag = f'<hp:t>{new_val}</hp:t>'
    t_xml_new = t_xml.replace(old_tag, new_tag, 1)
    return xml[:pos[0]] + t_xml_new + xml[pos[1]:]


def replace_regex_in_tags(xml, pattern, replacement):
    """정규식으로 hp:t 내용 치환"""
    return re.sub(pattern, replacement, xml, flags=re.DOTALL)


# ─────────────────────────────────────────────
# 신규 입간판 템플릿 생성
# ─────────────────────────────────────────────

def create_shinyu_ipganpan():
    """신규 입간판 HWPX → 플레이스홀더 템플릿"""
    src = f'{SHINYU_DIR}/심의 검토서(입간판).hwpx'
    files, infos = read_hwpx(src)
    xml = get_xml(files)

    # 광고물 내역 변수 치환
    xml = replace_tags(xml, {
        '김태욱':                    '__광고주__',
        '도산대로 111, 1층 (신사동)':  '__설치장소__',
        '약, 아이디약국':             '__표시내용__',
        '0.515*0.99':               '__규격__',   # 광고물내역 + 점검내역 두 곳 치환
    })

    # 수량 '1' - 첫 번째 테이블(광고물내역)에서만 치환
    xml = replace_first_in_table(xml, 0, '1', '__수량__')

    # 점검 내역 신청내용 치환
    xml = replace_tags(xml, {
        '상업지역': '__지역__',
        '자사광고': '__광고유형__',
        '비조명':   '__조명__',
        '1개':      '__수량__',
    })

    # 작성자
    xml = replace_regex_in_tags(xml,
        r'<hp:t>\[작성자 : [^\]]+\]</hp:t>',
        '<hp:t>[작성자 : __작성자__]</hp:t>')

    out = f'{OUT_DIR}/신규_입간판.hwpx'
    write_hwpx(set_xml(files, xml), infos, out)
    validate_hwpx_structure(out)
    print(f'* {out}')


# ─────────────────────────────────────────────
# 신규 비-입간판 템플릿 생성 (연장 HWPX 기반)
# 구조: 입간판 신규 base 위에 연장의 점검내역(안전점검 제거) 이식
# ─────────────────────────────────────────────

def extract_check_table_from_yeonjang(yeonjang_xml, with_placeholders: dict = None):
    """연장 XML에서 점검내역 테이블(2번째, index=1) 추출 후
    안전점검 행 제거 + 플레이스홀더 적용"""
    pos = find_nth_table(yeonjang_xml, 1)
    check_table = yeonjang_xml[pos[0]:pos[1]]
    check_table = remove_row_containing(check_table, '안전점검')
    if with_placeholders:
        check_table = replace_tags(check_table, with_placeholders)
    return check_table


def build_shinyu_from_base(base_xml, new_check_table_xml, signtypes_label):
    """신규 입간판 base XML에서:
    1. 점검내역 테이블(index=1)을 새 것으로 교체
    2. 광고물내역의 종류 값 변경
    3. 검토의견 일반화
    """
    # 점검내역 교체
    pos = find_nth_table(base_xml, 1)
    xml = base_xml[:pos[0]] + new_check_table_xml + base_xml[pos[1]:]

    # 광고물내역의 종류 셀: 첫 번째 테이블 안의 '입간판' → 새 종류
    # (입간판 템플릿의 광고물내역 종류 데이터 셀에만 적용)
    pos0 = find_nth_table(xml, 0)
    t0 = xml[pos0[0]:pos0[1]]
    # 광고물내역 데이터 행의 첫 번째 셀이 종류 → 'rowAddr="1"' 포함 셀
    # 간단히: 데이터 행(두 번째 tr)에서 첫 번째 <hp:t>입간판</hp:t> 치환
    # 헤더행에는 '입간판'이 없으므로 테이블 내 첫 번째 '입간판' = 데이터셀
    t0_new = t0.replace('<hp:t>입간판</hp:t>', f'<hp:t>{signtypes_label}</hp:t>', 1)
    xml = xml[:pos0[0]] + t0_new + xml[pos0[1]:]

    # 검토의견 업데이트 (신규용 일반 문구)
    # 입간판 기준 문구 → 범용 문구
    xml = xml.replace(
        '건물 부지 내 설치하는 입간판으로, 옥외광고물 관련 법령에 따른 기준에 적합하며, 표시 내용 또한 적합함.',
        '옥외광고물 관련 법령에 따른 기준에 적합하며, 표시 내용 또한 적합함.'
    )

    return xml


def create_shinyu_type(yeonjang_source, output_filename, signtypes_label,
                        check_replacements: dict):
    """신규 비-입간판 타입 템플릿 생성"""
    # 입간판 신규 base 로드 + 기본 플레이스홀더 적용
    ipganpan_files, ipganpan_infos = read_hwpx(f'{SHINYU_DIR}/심의 검토서(입간판).hwpx')
    base_xml = get_xml(ipganpan_files)

    # 광고물내역 플레이스홀더
    base_xml = replace_tags(base_xml, {
        '김태욱':                    '__광고주__',
        '도산대로 111, 1층 (신사동)':  '__설치장소__',
        '약, 아이디약국':             '__표시내용__',
        '0.515*0.99':               '__규격__',
    })
    base_xml = replace_first_in_table(base_xml, 0, '1', '__수량__')
    base_xml = replace_regex_in_tags(base_xml,
        r'<hp:t>\[작성자 : [^\]]+\]</hp:t>',
        '<hp:t>[작성자 : __작성자__]</hp:t>')

    # 연장 HWPX에서 점검내역 추출 (안전점검 제거 + 플레이스홀더)
    if os.path.isabs(yeonjang_source):
        yeonjang_path = yeonjang_source
    else:
        yeonjang_path = f'{YEONJANG_DIR}/{yeonjang_source}'

    yeonjang_files, _ = read_hwpx(yeonjang_path)
    yeonjang_xml = get_xml(yeonjang_files)
    new_check = extract_check_table_from_yeonjang(yeonjang_xml, check_replacements)

    # 점검 테이블 스타일은 base와 ID 충돌이 많아서 재매핑 후 병합한다.
    yeonjang_header = yeonjang_files['Contents/header.xml'].decode('utf-8')
    merged_header, new_check = merge_styles_with_remap(
        ipganpan_files['Contents/header.xml'].decode('utf-8'),
        yeonjang_header,
        new_check
    )
    merged_files = dict(ipganpan_files)
    merged_files['Contents/header.xml'] = merged_header.encode('utf-8')

    out = f'{OUT_DIR}/{output_filename}'
    result_xml = build_shinyu_from_base(base_xml, new_check, signtypes_label)
    write_hwpx(set_xml(merged_files, result_xml), ipganpan_infos, out)
    validate_hwpx_structure(out)
    print(f'* {out}')


def create_template_from_reference(reference_path, output_filename, replacements,
                                   first_table_replacements=None, regex_replacements=None):
    files, infos = read_hwpx(reference_path)
    xml = get_xml(files)

    if first_table_replacements:
        for old_val, new_val in first_table_replacements:
            xml = replace_first_in_table(xml, 0, old_val, new_val)

    xml = replace_tags(xml, replacements)

    if regex_replacements:
        for pattern, replacement in regex_replacements:
            xml = replace_regex_in_tags(xml, pattern, replacement)

    out = f'{OUT_DIR}/{output_filename}'
    write_hwpx(set_xml(files, xml), infos, out)
    validate_hwpx_structure(out)
    print(f'* {out} (reference)')


def create_shinyu_byeokmyeon_template():
    output_filename = '신규_벽면이용간판.hwpx'
    create_shinyu_type(
        yeonjang_source='검토서 및 허가증(벽면 연장).hwpx',
        output_filename=output_filename,
        signtypes_label='벽면이용간판',
        check_replacements={
            '주거지역': '__지역__',
            '자사광고': '__광고유형__',
            '7층':      '__위치_층__',
            '1개':      '__수량__',
            '3.26*0.9': '__규격__',
            'LED':      '__조명__',
        }
    )

    out_path = os.path.join(OUT_DIR, output_filename)
    files, infos = read_hwpx(out_path)
    header_xml = files['Contents/header.xml'].decode('utf-8')
    header_xml = re.sub(
        r'(<hh:charPr[^>]+)borderFillIDRef="5"',
        r'\1borderFillIDRef="1"',
        header_xml
    )
    header_xml = re.sub(
        r'(<hh:border[^>]+)borderFillIDRef="5"',
        r'\1borderFillIDRef="2"',
        header_xml
    )
    files['Contents/header.xml'] = header_xml.encode('utf-8')
    write_hwpx(files, infos, out_path)
    validate_hwpx_structure(out_path)


# ─────────────────────────────────────────────
# 연장 템플릿 생성
# ─────────────────────────────────────────────

def create_yeonjang_template(src_filename, out_filename, replacements,
                              split_pairs=None):
    """연장 HWPX → 플레이스홀더 템플릿
    split_pairs: [(old1, old2, new)] - 두 hp:t 노드를 하나로 병합하는 케이스
    """
    src = f'{YEONJANG_DIR}/{src_filename}'
    files, infos = read_hwpx(src)
    xml = get_xml(files)

    # 분리된 텍스트 노드 병합 (예: '대치서울' + '정형외과의원')
    if split_pairs:
        for old1, old2, new_val in split_pairs:
            xml = re.sub(
                rf'<hp:t>{re.escape(old1)}</hp:t>(.*?)<hp:t>{re.escape(old2)}</hp:t>',
                f'<hp:t>{new_val}</hp:t>',
                xml, count=1, flags=re.DOTALL
            )

    # 일반 치환
    xml = replace_tags(xml, replacements)

    # 수량 '1' - 광고물내역(첫 번째 테이블) 내 첫 치환
    xml = replace_first_in_table(xml, 0, '1', '__수량__')

    # 작성자
    xml = replace_regex_in_tags(xml,
        r'<hp:t>\[작성자 : [^\]]+\]</hp:t>',
        '<hp:t>[작성자 : __작성자__]</hp:t>')

    out = f'{OUT_DIR}/{out_filename}'
    write_hwpx(set_xml(files, xml), infos, out)
    validate_hwpx_structure(out)
    print(f'* {out}')


# ─────────────────────────────────────────────
# 각 템플릿 생성 실행
# ─────────────────────────────────────────────

def main():
    print('=== 광고물 심의 검토서 템플릿 생성 ===\n')

    # ── 신규 입간판 ──────────────────────────
    create_shinyu_ipganpan()

    # ── 신규 10층 이하 상단간판 ──────────────
    create_shinyu_type(
        yeonjang_source='검토서 및 허가증(10층 이하 상단 연장).hwpx',
        output_filename='신규_10층이하상단.hwpx',
        signtypes_label='벽면이용간판(가로형상단)',
        check_replacements={
            '주거지역': '__지역__',
            '자사광고': '__광고유형__',
            '9층':      '__위치_층__',
            '1개':      '__수량__',
            '10*1.1':   '__규격__',
            'LED':      '__조명__',
        }
    )

    # ── 신규 11층 이상 상단간판 ──────────────
    create_shinyu_type(
        yeonjang_source='검토서 및 허가증(11층 이상 상단 연장).hwpx',
        output_filename='신규_11층이상상단.hwpx',
        signtypes_label='벽면이용간판(가로형상단)',
        check_replacements={
            '상업지역': '__지역__',
            '자사광고': '__광고유형__',
            '20층':     '__위치_층__',
            '1개':      '__수량__',
            '17.5*3':   '__규격__',
            'LED':      '__조명__',
        }
    )

    # ── 신규 돌출간판 ────────────────────────
    create_shinyu_type(
        yeonjang_source='검토서 및 허가증(돌출 연장).hwpx',
        output_filename='신규_돌출간판.hwpx',
        signtypes_label='돌출간판',
        check_replacements={
            '주거지역': '__지역__',
            '자사광고': '__광고유형__',
            '3층':      '__위치_층__',
            '1개':      '__수량__',
            '0.7*3':    '__규격__',
            'LED':      '__조명__',
        }
    )

    # ── 신규 벽면이용간판 ────────────────────
    create_shinyu_byeokmyeon_template()

    # ── 연장 10층 이하 상단 ──────────────────
    create_yeonjang_template(
        src_filename='검토서 및 허가증(10층 이하 상단 연장).hwpx',
        out_filename='연장_10층이하상단.hwpx',
        replacements={
            '2019-3220174-09-2-00061':   '__신고번호__',
            '언주로 864, 9층 상단(신사동)': '__설치장소__',
            'Seoul Auction':             '__표시내용__',
            '10*1.1':                    '__규격__',
            '2022. 6. 3.':              '__변경전시작__',
            '2025. 6. 2.':              '__변경전종료__',
            '2025. 6. 3.':              '__변경후시작__',
            '2028. 6. 2.':              '__변경후종료__',
            '주거지역':                  '__지역__',
            '자사광고':                  '__광고유형__',
            '9층':                       '__위치_층__',
            '1개':                       '__수량__',
            'LED':                       '__조명__',
            '○ 안전점검 합격(2025. 6. 4.)': '○ 안전점검 합격(__안전점검일__)',
        }
    )

    # ── 연장 11층 이상 상단 ──────────────────
    create_yeonjang_template(
        src_filename='검토서 및 허가증(11층 이상 상단 연장).hwpx',
        out_filename='연장_11층이상상단.hwpx',
        replacements={
            '2022-3220174-09-1-00062':    '__신고번호__',
            '학동로 343, 20층 상단(논현동)':  '__설치장소__',
            '(로고)FADU':                 '__표시내용__',
            '17.5*3':                     '__규격__',
            '2022. 5. 2.':               '__변경전시작__',
            '2025. 5. 1.':               '__변경전종료__',
            '2025. 5. 2.':               '__변경후시작__',
            '2028. 5. 1.':               '__변경후종료__',
            '상업지역':                    '__지역__',
            '자사광고':                    '__광고유형__',
            '20층':                        '__위치_층__',
            '1개':                         '__수량__',
            'LED':                         '__조명__',
            '○ 안전점검 합격(2025. 5. 21.)': '○ 안전점검 합격(__안전점검일__)',
        }
    )

    # ── 연장 돌출 ────────────────────────────
    create_yeonjang_template(
        src_filename='검토서 및 허가증(돌출 연장).hwpx',
        out_filename='연장_돌출간판.hwpx',
        replacements={
            '2022-3220174-09-1-00093':  '__신고번호__',
            '남부순환로 2947, 3층(대치동)': '__설치장소__',
            '0.7*3':                    '__규격__',
            '2022. 5. 20.':            '__변경전시작__',
            '2025. 5. 19.':            '__변경전종료__',
            '2025. 5. 20.':            '__변경후시작__',
            '2028. 5. 19.':            '__변경후종료__',
            '주거지역':                 '__지역__',
            '자사광고':                 '__광고유형__',
            '3층':                      '__위치_층__',
            '1개':                      '__수량__',
            'LED':                      '__조명__',
            # 안전점검 날짜는 별도 hp:t 노드로 분리됨
            '2025. 6. 20.':            '__안전점검일__',
        },
        split_pairs=[
            # '대치서울' + '정형외과의원' → 하나의 플레이스홀더
            ('대치서울', '정형외과의원', '__표시내용__'),
        ]
    )

    # ── 연장 벽면 ────────────────────────────
    create_yeonjang_template(
        src_filename='검토서 및 허가증(벽면 연장).hwpx',
        out_filename='연장_벽면이용간판.hwpx',
        replacements={
            '2019-3220174-09-2-00151':  '__신고번호__',
            '선릉로 635, 7층(논현동)':   '__설치장소__',
            '(로고)LG U+':              '__표시내용__',
            '3.26*0.9':                 '__규격__',
            '2022. 11. 29. ~ ':         '__변경전시작__ ~ ',   # 물결표 포함 노드
            '2025. 11. 28.':            '__변경전종료__',
            '2025. 11. 29.':            '__변경후시작__',
            '2028. 11. 28.':            '__변경후종료__',
            '주거지역':                  '__지역__',
            '자사광고':                  '__광고유형__',
            '7층':                       '__위치_층__',
            '1개':                       '__수량__',
            'LED':                       '__조명__',
            '○ 안전점검 합격(2026. 1. 9.)': '○ 안전점검 합격(__안전점검일__)',
        }
    )

    print('\n완료! 모든 템플릿이 생성되었습니다.')
    print(f'위치: {OUT_DIR}')



if __name__ == '__main__':
    main()
