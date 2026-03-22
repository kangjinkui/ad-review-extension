"""
광고물 심의 검토서 자동 생성 프로그램
간판 종류 및 신규/연장 구분에 따라 한글(HWPX) 파일을 자동으로 생성합니다.
"""

import os
import sys
import zipfile
import re
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
from io import BytesIO
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET

HWP_NS = {
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hhs': 'http://www.hancom.co.kr/hwpml/2011/history',
    'hm': 'http://www.hancom.co.kr/hwpml/2011/master-page',
    'hpf': 'http://www.hancom.co.kr/schema/2011/hpf',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'opf': 'http://www.idpf.org/2007/opf/',
    'ooxmlchart': 'http://www.hancom.co.kr/hwpml/2016/ooxmlchart',
    'hwpunitchar': 'http://www.hancom.co.kr/hwpml/2016/HwpUnitChar',
    'epub': 'http://www.idpf.org/2007/ops',
    'config': 'urn:oasis:names:tc:opendocument:xmlns:config:1.0',
}
for prefix, uri in HWP_NS.items():
    ET.register_namespace(prefix, uri)


# ─────────────────────────────────────────────
# 리소스 경로 (PyInstaller 호환)
# ─────────────────────────────────────────────

def resource_path(relative_path):
    """PyInstaller exe 안에 번들된 리소스 경로 반환"""
    if hasattr(sys, '_MEIPASS'):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


TEMPLATE_DIR = resource_path('templates')
APP_STATE_PATH = os.path.join(
    os.environ.get('LOCALAPPDATA', os.path.dirname(os.path.abspath(__file__))),
    'ad-review-builder',
    'state.json'
)
REQUIRED_HWPX_ENTRIES = [
    'mimetype',
    'version.xml',
    'Contents/header.xml',
    'Contents/section0.xml',
    'Contents/content.hpf',
    'META-INF/container.xml',
    'META-INF/manifest.xml',
]

# 간판 종류 → 템플릿 파일 매핑
TEMPLATE_MAP = {
    ('신규', '입간판'):          '신규_입간판.hwpx',
    ('신규', '10층 이하 상단간판'): '신규_10층이하상단.hwpx',
    ('신규', '11층 이상 상단간판'): '신규_11층이상상단.hwpx',
    ('신규', '돌출간판'):         '신규_돌출간판.hwpx',
    ('신규', '벽면이용간판'):      '신규_벽면이용간판.hwpx',
    ('연장', '10층 이하 상단'):    '연장_10층이하상단.hwpx',
    ('연장', '11층 이상 상단'):    '연장_11층이상상단.hwpx',
    ('연장', '돌출간판'):          '연장_돌출간판.hwpx',
    ('연장', '벽면이용간판'):      '연장_벽면이용간판.hwpx',
    ('내용변경', '공공시설물'):    '내용변경_공공시설물.hwpx',
}

SHINYU_TYPES = ['입간판', '10층 이하 상단간판', '11층 이상 상단간판', '돌출간판', '벽면이용간판']
YEONJANG_TYPES = ['10층 이하 상단', '11층 이상 상단', '돌출간판', '벽면이용간판']
CONTENT_CHANGE_TYPE = '공공시설물'

JIYEOK_OPTIONS = ['상업지역', '주거지역', '준주거지역', '공업지역', '전용주거지역']
GWANGGO_OPTIONS = ['자사광고', '타사광고']
JOMYEONG_OPTIONS = ['LED', '비조명', '형광등', '백열등', '기타']


# ─────────────────────────────────────────────
# HWPX 생성 엔진
# ─────────────────────────────────────────────

def _sanitize_xml_text(value: str) -> str:
    # XML 1.0에서 허용되지 않는 제어 문자를 제거한다.
    return ''.join(
        ch for ch in value
        if ch in '\t\n\r' or ord(ch) >= 0x20
    )


def _build_clean_zipinfo(item: zipfile.ZipInfo) -> zipfile.ZipInfo:
    zi = zipfile.ZipInfo(item.filename, item.date_time)
    zi.compress_type = item.compress_type
    zi.create_system = item.create_system
    zi.create_version = item.create_version
    zi.extract_version = item.extract_version
    zi.external_attr = item.external_attr
    zi.internal_attr = item.internal_attr
    zi.flag_bits = 0
    zi.extra = b''
    zi.comment = b''
    return zi


def _validate_hwpx_bytes(hwpx_bytes: bytes):
    with zipfile.ZipFile(BytesIO(hwpx_bytes), 'r') as zf:
        names = zf.namelist()
        missing_entries = [name for name in REQUIRED_HWPX_ENTRIES if name not in names]
        if missing_entries:
            raise ValueError(f'HWPX 필수 항목이 누락되었습니다: {missing_entries}')
        if names[:1] != ['mimetype']:
            raise ValueError('HWPX ZIP 첫 번째 엔트리가 mimetype이 아닙니다.')
        mimetype_info = zf.getinfo('mimetype')
        if mimetype_info.compress_type != zipfile.ZIP_STORED:
            raise ValueError('HWPX mimetype 엔트리가 비압축 상태가 아닙니다.')
        header_xml = zf.read('Contents/header.xml').decode('utf-8')
        section_xml = zf.read('Contents/section0.xml').decode('utf-8')
        ET.fromstring(header_xml)
        ET.fromstring(section_xml)


def fill_template(template_name: str, values: dict) -> bytes:
    """템플릿 HWPX에 values 딕셔너리의 값을 채워 bytes로 반환"""
    tpl_path = os.path.join(TEMPLATE_DIR, template_name)
    if not os.path.exists(tpl_path):
        raise FileNotFoundError(f'템플릿 파일을 찾을 수 없습니다: {tpl_path}')

    buf = BytesIO()
    with zipfile.ZipFile(tpl_path, 'r') as zin, \
         zipfile.ZipFile(buf, 'w') as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'Contents/section0.xml':
                xml = data.decode('utf-8')
                if template_name.startswith('연장_'):
                    xml = xml.replace(
                        '<hp:t>표시내용</hp:t>',
                        '<hp:t>상호명/표시내용</hp:t>',
                        1
                    )
                    xml = _replace_multiline_placeholder_paragraphs(
                        xml, '__신고번호__', values.get('신고번호', '')
                    )
                    xml = _replace_multiline_placeholder_paragraphs(
                        xml, '__표시내용__', values.get('표시내용', '')
                    )
                elif template_name == '내용변경_공공시설물.hwpx':
                    xml = _fill_content_change_template(xml, values)
                for key, val in values.items():
                    xml = xml.replace(f'__{key}__', escape(_sanitize_xml_text(val)))
                data = xml.encode('utf-8')
            # Hancom은 ZIP 메타데이터 차이에 민감하므로 깨끗한 헤더만 재구성한다.
            zi = _build_clean_zipinfo(item)
            zout.writestr(zi, data)

    hwpx_bytes = buf.getvalue()
    _validate_hwpx_bytes(hwpx_bytes)
    return hwpx_bytes


def generate_file(mode: str, sign_type: str, values: dict,
                  output_dir: str, folder_name: str) -> str:
    """HWPX 파일 생성 후 경로 반환"""
    tpl_name = TEMPLATE_MAP.get((mode, sign_type))
    if not tpl_name:
        raise ValueError(f'지원하지 않는 조합: {mode} / {sign_type}')

    hwpx_bytes = fill_template(tpl_name, values)

    # 폴더 생성
    target_dir = os.path.join(output_dir, folder_name)
    os.makedirs(target_dir, exist_ok=True)

    # 파일명: 심의검토서(간판종류)_업소명.hwpx
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', folder_name)
    if mode == '신규':
        prefix = '심의검토서'
        filename = f'{prefix}({sign_type})_{safe_name}.hwpx'
    elif mode == '내용변경':
        prefix = '공공시설물내용변경검토서'
        filename = f'{prefix}_{safe_name}.hwpx'
    else:
        prefix = '검토서및허가증'
        filename = f'{prefix}({sign_type})_{safe_name}.hwpx'
    out_path = os.path.join(target_dir, filename)

    with open(out_path, 'wb') as f:
        f.write(hwpx_bytes)

    return out_path


def load_app_state() -> dict:
    try:
        with open(APP_STATE_PATH, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass
    return {}


def save_app_state(state: dict):
    try:
        os.makedirs(os.path.dirname(APP_STATE_PATH), exist_ok=True)
        with open(APP_STATE_PATH, 'w', encoding='utf-8') as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_last_output_dir() -> str:
    value = load_app_state().get('last_output_dir', '')
    if isinstance(value, str) and value.strip():
        return value.strip()
    return ''


def set_last_output_dir(path: str):
    path = (path or '').strip()
    if not path:
        return
    state = load_app_state()
    state['last_output_dir'] = path
    save_app_state(state)


# ─────────────────────────────────────────────
# GUI 메인 클래스
# ─────────────────────────────────────────────

class AdReviewApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('광고물 심의 검토서 자동 생성')
        self.resizable(False, False)
        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f'+{x}+{y}')

    def _build_ui(self):
        # 탭 노트북
        nb = ttk.Notebook(self)
        nb.pack(fill='both', expand=True, padx=10, pady=10)

        self.tab_shinyu = ShinYuTab(nb)
        self.tab_yeonjang = YeonjangTab(nb)
        self.tab_content_change = ContentChangeTab(nb)

        nb.add(self.tab_shinyu,   text='  신규 심의 검토서  ')
        nb.add(self.tab_yeonjang, text='  연장 검토서 및 허가증  ')
        nb.add(self.tab_content_change, text='  내용변경  ')


# ─────────────────────────────────────────────
# 공통 탭 기반 클래스
# ─────────────────────────────────────────────

class BaseTab(ttk.Frame):
    PADDING = {'padx': 8, 'pady': 4}

    def __init__(self, parent, mode):
        super().__init__(parent, padding=10)
        self.mode = mode
        self._vars = {}
        self._text_widgets = {}
        self._build()

    def _lf(self, parent, text):
        """LabelFrame 생성"""
        lf = ttk.LabelFrame(parent, text=text, padding=6)
        return lf

    def _row(self, parent, label, widget_factory, row, col=0, span=1, **kw):
        """라벨 + 입력위젯 행 추가"""
        ttk.Label(parent, text=label).grid(
            row=row, column=col*2, sticky='e', **self.PADDING)
        w = widget_factory(parent, **kw)
        w.grid(row=row, column=col*2+1, sticky='ew', **self.PADDING)
        return w

    def _entry(self, parent, var=None, width=30, **kw):
        if var is None:
            var = tk.StringVar()
        e = ttk.Entry(parent, textvariable=var, width=width, **kw)
        return e, var

    def _combo(self, parent, options, var=None, width=18, **kw):
        if var is None:
            var = tk.StringVar(value=options[0])
        c = ttk.Combobox(parent, textvariable=var, values=options,
                          state='readonly', width=width, **kw)
        return c, var

    def _add_entry_row(self, parent, label, row, key, default='', width=32):
        var = tk.StringVar(value=default)
        self._vars[key] = var
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='e', **self.PADDING)
        entry = ttk.Entry(parent, textvariable=var, width=width)
        _bind_korean_ime(entry)
        entry.grid(row=row, column=1, columnspan=3, sticky='ew', **self.PADDING)

    def _add_combo_row(self, parent, label, row, key, options, col=0):
        var = tk.StringVar(value=options[0])
        self._vars[key] = var
        ttk.Label(parent, text=label).grid(
            row=row, column=col*2, sticky='e', **self.PADDING)
        ttk.Combobox(parent, textvariable=var, values=options,
                     state='readonly', width=18).grid(
            row=row, column=col*2+1, sticky='ew', **self.PADDING)

    def _add_text_row(self, parent, label, row, key, height=3, width=40):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='ne', **self.PADDING)
        text = tk.Text(parent, width=width, height=height, font=('Malgun Gothic', 9))
        _bind_korean_ime(text)
        text.grid(row=row, column=1, columnspan=3, sticky='ew', **self.PADDING)
        self._text_widgets[key] = text

    def _get(self, key):
        if key in self._text_widgets:
            return self._text_widgets[key].get('1.0', 'end-1c').strip()
        return self._vars[key].get().strip()

    def _build(self):
        raise NotImplementedError

    def _select_folder(self, var):
        folder = filedialog.askdirectory(title='출력 폴더 선택')
        if folder:
            var.set(folder)
            set_last_output_dir(folder)

    def _on_generate(self):
        raise NotImplementedError

    def _create_output_dir_var(self):
        return tk.StringVar(value=get_last_output_dir())

    def _get_or_ask_dir(self):
        d = self.out_dir_var.get().strip()
        if not d:
            d = filedialog.askdirectory(title='출력 폴더 선택')
            if d:
                self.out_dir_var.set(d)
        if d:
            set_last_output_dir(d)
        return d

    def _show_success(self, path):
        if messagebox.askyesno(
            '생성 완료',
            f'파일이 생성되었습니다!\n\n{path}\n\n폴더를 열겠습니까?'
        ):
            folder = os.path.dirname(path)
            os.startfile(folder) if sys.platform == 'win32' else \
                os.system(f'xdg-open "{folder}"')


# ─────────────────────────────────────────────
# 신규 탭
# ─────────────────────────────────────────────

class ShinYuTab(BaseTab):
    def __init__(self, parent):
        super().__init__(parent, mode='신규')

    def _build(self):
        # ── 간판 종류 ──────────────────────────
        lf_type = self._lf(self, '간판 종류')
        lf_type.pack(fill='x', pady=(0, 6))
        self.sign_type_var = tk.StringVar(value=SHINYU_TYPES[0])
        for i, t in enumerate(SHINYU_TYPES):
            rb = ttk.Radiobutton(lf_type, text=t, variable=self.sign_type_var,
                                  value=t, command=self._on_type_change)
            rb.grid(row=0, column=i, padx=10, pady=4)

        # ── 광고물 내역 ────────────────────────
        lf_info = self._lf(self, '광고물 내역')
        lf_info.pack(fill='x', pady=(0, 6))
        lf_info.columnconfigure(1, weight=1)
        lf_info.columnconfigure(3, weight=1)

        self._add_entry_row(lf_info, '광고주 *', 0, '광고주')
        self._add_entry_row(lf_info, '설치 장소 *', 1, '설치장소')
        self._add_entry_row(lf_info, '표시 내용 *', 2, '표시내용')

        # 규격 + 수량 한 행에
        ttk.Label(lf_info, text='규격 (M) *').grid(row=3, column=0, sticky='e', **self.PADDING)
        self._vars['규격'] = tk.StringVar()
        ent_size = ttk.Entry(lf_info, textvariable=self._vars['규격'], width=16)
        _bind_korean_ime(ent_size)
        ent_size.grid(row=3, column=1, sticky='ew', **self.PADDING)
        ttk.Label(lf_info, text='수량 *').grid(row=3, column=2, sticky='e', **self.PADDING)
        self._vars['수량'] = tk.StringVar(value='1')
        ent_count = ttk.Entry(lf_info, textvariable=self._vars['수량'], width=6)
        _bind_korean_ime(ent_count)
        ent_count.grid(row=3, column=3, sticky='w', **self.PADDING)

        # ── 점검 내역 ──────────────────────────
        lf_check = self._lf(self, '점검 내역 신청내용')
        lf_check.pack(fill='x', pady=(0, 6))

        self._add_combo_row(lf_check, '지역 *',     0, '지역',     JIYEOK_OPTIONS)
        self._add_combo_row(lf_check, '광고유형 *',  0, '광고유형', GWANGGO_OPTIONS, col=1)
        self._add_combo_row(lf_check, '조명 *',     1, '조명',     JOMYEONG_OPTIONS)

        # 위치(층수) - 입간판 제외
        self.lbl_wichi = ttk.Label(lf_check, text='위치(층수) *')
        self._vars['위치_층'] = tk.StringVar()
        self.ent_wichi = ttk.Entry(lf_check, textvariable=self._vars['위치_층'], width=12)
        _bind_korean_ime(self.ent_wichi)
        self.lbl_wichi.grid(row=1, column=2, sticky='e', **self.PADDING)
        self.ent_wichi.grid(row=1, column=3, sticky='w', **self.PADDING)

        # ── 작성자 ────────────────────────────
        lf_author = self._lf(self, '작성자')
        lf_author.pack(fill='x', pady=(0, 6))
        self._add_entry_row(lf_author, '작성자', 0, '작성자', default='', width=20)

        # ── 출력 설정 ──────────────────────────
        lf_out = self._lf(self, '출력 설정')
        lf_out.pack(fill='x', pady=(0, 8))
        lf_out.columnconfigure(1, weight=1)

        ttk.Label(lf_out, text='출력 폴더 *').grid(row=0, column=0, sticky='e', **self.PADDING)
        self.out_dir_var = self._create_output_dir_var()
        out_entry = ttk.Entry(lf_out, textvariable=self.out_dir_var, width=38)
        _bind_korean_ime(out_entry)
        out_entry.grid(row=0, column=1, sticky='ew', **self.PADDING)
        ttk.Button(lf_out, text='찾아보기',
                   command=lambda: self._select_folder(self.out_dir_var)).grid(
            row=0, column=2, padx=4, pady=4)

        # ── 생성 버튼 ──────────────────────────
        ttk.Button(self, text='📄  파일 생성', style='Accent.TButton',
                   command=self._on_generate).pack(pady=6, ipadx=20, ipady=6)

        self._on_type_change()

    def _on_type_change(self):
        """입간판 선택 시 위치(층수) 필드 숨김"""
        show = self.sign_type_var.get() != '입간판'
        if show:
            self.lbl_wichi.grid()
            self.ent_wichi.grid()
        else:
            self.lbl_wichi.grid_remove()
            self.ent_wichi.grid_remove()

    def _on_generate(self):
        sign_type = self.sign_type_var.get()
        out_dir = self._get_or_ask_dir()
        if not out_dir:
            return

        # 필수값 검증
        required = ['광고주', '설치장소', '표시내용', '규격', '수량']
        for k in required:
            if not self._get(k):
                messagebox.showwarning('입력 오류', f'{k}을(를) 입력해주세요.')
                return
        if sign_type != '입간판' and not self._get('위치_층'):
            messagebox.showwarning('입력 오류', '위치(층수)를 입력해주세요. (예: 9층)')
            return

        values = {
            '광고주':    self._get('광고주'),
            '설치장소':  self._get('설치장소'),
            '표시내용':  self._get('표시내용'),
            '규격':      self._get('규격'),
            '수량':      self._get('수량'),
            '지역':      self._get('지역'),
            '광고유형':  self._get('광고유형'),
            '조명':      self._get('조명'),
            '작성자':    self._get('작성자'),
            '위치_층':   self._get('위치_층'),
        }

        folder_name = self._get('광고주') or self._get('표시내용')
        try:
            out = generate_file('신규', sign_type, values, out_dir, folder_name)
            self._show_success(out)
        except Exception as e:
            messagebox.showerror('오류', f'파일 생성 중 오류가 발생했습니다.\n\n{e}')

# ─────────────────────────────────────────────
# 연장 탭
# ─────────────────────────────────────────────

class YeonjangTab(BaseTab):
    def __init__(self, parent):
        super().__init__(parent, mode='연장')

    def _build(self):
        # ── 간판 종류 ──────────────────────────
        lf_type = self._lf(self, '간판 종류')
        lf_type.pack(fill='x', pady=(0, 6))
        self.sign_type_var = tk.StringVar(value=YEONJANG_TYPES[0])
        for i, t in enumerate(YEONJANG_TYPES):
            ttk.Radiobutton(lf_type, text=t, variable=self.sign_type_var,
                             value=t).grid(row=0, column=i, padx=10, pady=4)

        # ── 광고물 내역 ────────────────────────
        lf_info = self._lf(self, '광고물 내역')
        lf_info.pack(fill='x', pady=(0, 6))
        lf_info.columnconfigure(1, weight=1)
        lf_info.columnconfigure(3, weight=1)

        self._add_entry_row(lf_info, '신고번호 *', 0, '신고번호',
                            default='2022-3220174-09-1-00001')
        self._add_entry_row(lf_info, '상호명 *', 1, '상호명')
        self._add_entry_row(lf_info, '설치 장소 *', 2, '설치장소')
        self._add_entry_row(lf_info, '표시 내용 *', 3, '표시내용')

        ttk.Label(lf_info, text='규격 (M) *').grid(row=4, column=0, sticky='e', **self.PADDING)
        self._vars['규격'] = tk.StringVar()
        ent_size = ttk.Entry(lf_info, textvariable=self._vars['규격'], width=16)
        _bind_korean_ime(ent_size)
        ent_size.grid(row=4, column=1, sticky='ew', **self.PADDING)
        ttk.Label(lf_info, text='수량 *').grid(row=4, column=2, sticky='e', **self.PADDING)
        self._vars['수량'] = tk.StringVar(value='1')
        ent_count = ttk.Entry(lf_info, textvariable=self._vars['수량'], width=6)
        _bind_korean_ime(ent_count)
        ent_count.grid(row=4, column=3, sticky='w', **self.PADDING)

        # ── 표시기간 ───────────────────────────
        lf_period = self._lf(self, '표시기간')
        lf_period.pack(fill='x', pady=(0, 6))

        r = 0
        for key, label in [('변경전시작', '변경전 시작일 *'), ('변경전종료', '변경전 종료일 *')]:
            ttk.Label(lf_period, text=label).grid(row=r//2, column=(r%2)*2,
                                                   sticky='e', **self.PADDING)
            self._vars[key] = tk.StringVar()
            e = ttk.Entry(lf_period, textvariable=self._vars[key], width=18)
            _bind_korean_ime(e)
            e.grid(row=r//2, column=(r%2)*2+1, sticky='ew', **self.PADDING)
            if key == '변경전시작':
                e.bind('<KeyRelease>', self._on_period_input, add='+')
                e.bind('<FocusOut>', self._on_period_input, add='+')
            r += 1
        for key, label in [('변경후시작', '변경후 시작일 *'), ('변경후종료', '변경후 종료일 *')]:
            ttk.Label(lf_period, text=label).grid(row=r//2, column=(r%2)*2,
                                                   sticky='e', **self.PADDING)
            self._vars[key] = tk.StringVar()
            e = ttk.Entry(lf_period, textvariable=self._vars[key], width=18)
            _bind_korean_ime(e)
            e.grid(row=r//2, column=(r%2)*2+1, sticky='ew', **self.PADDING)
            r += 1
        ttk.Label(lf_period, text='예) 2023.01.10.', foreground='gray').grid(
            row=0, column=4, sticky='w'
        )

        # ── 점검 내역 ──────────────────────────
        lf_check = self._lf(self, '점검 내역 신청내용')
        lf_check.pack(fill='x', pady=(0, 6))

        self._add_combo_row(lf_check, '지역 *',     0, '지역',     JIYEOK_OPTIONS)
        self._add_combo_row(lf_check, '광고유형 *',  0, '광고유형', GWANGGO_OPTIONS, col=1)
        self._add_combo_row(lf_check, '조명 *',     1, '조명',     JOMYEONG_OPTIONS)

        ttk.Label(lf_check, text='위치(층수) *').grid(row=1, column=2, sticky='e', **self.PADDING)
        self._vars['위치_층'] = tk.StringVar()
        ent_floor = ttk.Entry(lf_check, textvariable=self._vars['위치_층'], width=12)
        _bind_korean_ime(ent_floor)
        ent_floor.grid(row=1, column=3, sticky='w', **self.PADDING)

        ttk.Label(lf_check, text='안전점검일 *').grid(row=2, column=0, sticky='e', **self.PADDING)
        self._vars['안전점검일'] = tk.StringVar()
        ent_safe = ttk.Entry(lf_check, textvariable=self._vars['안전점검일'], width=18)
        _bind_korean_ime(ent_safe)
        ent_safe.grid(row=2, column=1, sticky='ew', **self.PADDING)
        ttk.Label(lf_check, text='예) 2025. 6. 4.', foreground='gray').grid(
            row=2, column=2, sticky='w')

        # ── 작성자 ────────────────────────────
        lf_author = self._lf(self, '작성자')
        lf_author.pack(fill='x', pady=(0, 6))
        self._add_entry_row(lf_author, '작성자', 0, '작성자', default='', width=20)

        # ── 출력 설정 ──────────────────────────
        lf_out = self._lf(self, '출력 설정')
        lf_out.pack(fill='x', pady=(0, 8))
        lf_out.columnconfigure(1, weight=1)

        ttk.Label(lf_out, text='출력 폴더 *').grid(row=0, column=0, sticky='e', **self.PADDING)
        self.out_dir_var = self._create_output_dir_var()
        out_entry = ttk.Entry(lf_out, textvariable=self.out_dir_var, width=38)
        _bind_korean_ime(out_entry)
        out_entry.grid(row=0, column=1, sticky='ew', **self.PADDING)
        ttk.Button(lf_out, text='찾아보기',
                   command=lambda: self._select_folder(self.out_dir_var)).grid(
            row=0, column=2, padx=4, pady=4)

        # ── 생성 버튼 ──────────────────────────
        ttk.Button(self, text='📄  파일 생성', style='Accent.TButton',
                   command=self._on_generate).pack(pady=6, ipadx=20, ipady=6)

    def _on_period_input(self, _event=None):
        self._on_period_start_change()

    def _on_period_start_change(self):
        start_text = self._get('변경전시작')
        start_date = _parse_korean_date(start_text)
        if not start_date:
            return

        before_end = _add_years(start_date, 3) - timedelta(days=1)
        after_start = before_end + timedelta(days=1)
        after_end = _add_years(after_start, 3) - timedelta(days=1)

        self._vars['변경전종료'].set(_format_korean_date(before_end))
        self._vars['변경후시작'].set(_format_korean_date(after_start))
        self._vars['변경후종료'].set(_format_korean_date(after_end))

    def _on_generate(self):
        sign_type = self.sign_type_var.get()
        out_dir = self._get_or_ask_dir()
        if not out_dir:
            return

        required = ['신고번호', '상호명', '설치장소', '표시내용', '규격', '수량',
                    '변경전시작', '변경전종료', '변경후시작', '변경후종료',
                    '위치_층', '안전점검일']
        for k in required:
            if not self._get(k):
                label = k.replace('_', ' ')
                messagebox.showwarning('입력 오류', f'{label}을(를) 입력해주세요.')
                return

        values = {k: self._get(k) for k in [
            '신고번호', '상호명', '설치장소', '표시내용', '규격', '수량',
            '변경전시작', '변경전종료', '변경후시작', '변경후종료',
            '지역', '광고유형', '조명', '위치_층', '안전점검일', '작성자'
        ]}
        values['신고번호'] = _format_report_number(values['신고번호'])
        values['표시내용'] = f"{values['상호명']}/\n{values['표시내용']}"

        folder_name = self._get('상호명')
        try:
            out = generate_file('연장', sign_type, values, out_dir, folder_name)
            self._show_success(out)
        except Exception as e:
            messagebox.showerror('오류', f'파일 생성 중 오류가 발생했습니다.\n\n{e}')

# ─────────────────────────────────────────────
# 내용변경 탭
# ─────────────────────────────────────────────

class ContentChangeTab(BaseTab):
    def __init__(self, parent):
        super().__init__(parent, mode='내용변경')

    def _build(self):
        lf_type = self._lf(self, '서식')
        lf_type.pack(fill='x', pady=(0, 6))
        ttk.Label(
            lf_type,
            text='공공시설물(지상변압기함) 내용변경 검토서',
        ).grid(row=0, column=0, sticky='w', padx=8, pady=6)

        lf_info = self._lf(self, '광고물 내역')
        lf_info.pack(fill='x', pady=(0, 6))
        lf_info.columnconfigure(1, weight=1)
        lf_info.columnconfigure(3, weight=1)

        self._add_entry_row(lf_info, '광고주 *', 0, '광고주')
        self._add_text_row(lf_info, '표시 위치 *', 1, '표시위치', height=3)
        self._add_text_row(lf_info, '표시 내용 *', 2, '표시내용', height=4)

        ttk.Label(lf_info, text='규격 (M) *').grid(row=3, column=0, sticky='e', **self.PADDING)
        self._vars['규격'] = tk.StringVar()
        ent_size = ttk.Entry(lf_info, textvariable=self._vars['규격'], width=16)
        _bind_korean_ime(ent_size)
        ent_size.grid(row=3, column=1, sticky='ew', **self.PADDING)
        ttk.Label(lf_info, text='수량 *').grid(row=3, column=2, sticky='e', **self.PADDING)
        self._vars['수량'] = tk.StringVar(value='1')
        ent_count = ttk.Entry(lf_info, textvariable=self._vars['수량'], width=6)
        _bind_korean_ime(ent_count)
        ent_count.grid(row=3, column=3, sticky='w', **self.PADDING)

        lf_author = self._lf(self, '검토자')
        lf_author.pack(fill='x', pady=(0, 6))
        self._add_entry_row(lf_author, '검토자', 0, '검토자', default='', width=20)

        lf_out = self._lf(self, '출력 설정')
        lf_out.pack(fill='x', pady=(0, 8))
        lf_out.columnconfigure(1, weight=1)

        ttk.Label(lf_out, text='출력 폴더 *').grid(row=0, column=0, sticky='e', **self.PADDING)
        self.out_dir_var = self._create_output_dir_var()
        out_entry = ttk.Entry(lf_out, textvariable=self.out_dir_var, width=38)
        _bind_korean_ime(out_entry)
        out_entry.grid(row=0, column=1, sticky='ew', **self.PADDING)
        ttk.Button(lf_out, text='찾아보기',
                   command=lambda: self._select_folder(self.out_dir_var)).grid(
            row=0, column=2, padx=4, pady=4)

        ttk.Button(self, text='📄  파일 생성', style='Accent.TButton',
                   command=self._on_generate).pack(pady=6, ipadx=20, ipady=6)

    def _on_generate(self):
        out_dir = self._get_or_ask_dir()
        if not out_dir:
            return

        required = ['광고주', '표시위치', '표시내용', '규격', '수량']
        for key in required:
            if not self._get(key):
                messagebox.showwarning('입력 오류', f'{key}을(를) 입력해주세요.')
                return

        values = {
            '광고주': self._get('광고주'),
            '표시위치': self._get('표시위치'),
            '표시내용': self._get('표시내용'),
            '규격': self._get('규격'),
            '수량': self._get('수량'),
            '검토자': self._get('검토자'),
        }

        folder_name = self._get('광고주')
        try:
            out = generate_file('내용변경', CONTENT_CHANGE_TYPE, values, out_dir, folder_name)
            self._show_success(out)
        except Exception as e:
            messagebox.showerror('오류', f'파일 생성 중 오류가 발생했습니다.\n\n{e}')

# ─────────────────────────────────────────────
# 유틸리티
# ─────────────────────────────────────────────

def _bind_korean_ime(widget):
    widget.bind('<FocusIn>', _activate_korean_ime, add='+')


def _activate_korean_ime(_event=None):
    if sys.platform != 'win32':
        return
    try:
        import ctypes
        user32 = ctypes.WinDLL('user32', use_last_error=True)
        hkl = user32.LoadKeyboardLayoutW('00000412', 1)
        if hkl:
            user32.ActivateKeyboardLayout(hkl, 0)
    except Exception:
        pass


def _parse_korean_date(value: str):
    cleaned = re.sub(r'[^0-9]', '', value)
    if len(cleaned) != 8:
        return None
    try:
        return datetime.strptime(cleaned, '%Y%m%d').date()
    except ValueError:
        return None


def _format_korean_date(value):
    return value.strftime('%Y.%m.%d.')


def _add_years(value, years: int):
    try:
        return value.replace(year=value.year + years)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + years)


def _format_report_number(value: str):
    value = value.strip()
    if not value or '\n' in value:
        return value
    parts = value.split('-')
    if len(parts) >= 2:
        return '-'.join(parts[:-1]) + '-\n' + parts[-1]
    midpoint = len(value) // 2
    return value[:midpoint] + '\n' + value[midpoint:]


def _replace_multiline_placeholder_paragraphs(xml: str, placeholder: str, value: str) -> str:
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if len(lines) <= 1:
        return xml

    pattern = re.compile(
        rf'<hp:p (?P<p_attrs>[^>]*)>(?P<prefix_runs>.*?)<hp:run charPrIDRef="(?P<char_id>\d+)"><hp:t>{re.escape(placeholder)}</hp:t></hp:run>'
        rf'<hp:linesegarray><hp:lineseg (?P<line_attrs>[^>]*)/></hp:linesegarray></hp:p>'
        ,
        flags=re.DOTALL
    )
    match = pattern.search(xml)
    if not match:
        return xml

    p_attrs = match.group('p_attrs')
    prefix_runs = match.group('prefix_runs')
    char_id = match.group('char_id')
    line_attrs = match.group('line_attrs')
    horz_match = re.search(r'horzsize="([^"]+)"', line_attrs)
    flags_match = re.search(r'flags="([^"]+)"', line_attrs)
    baseline_match = re.search(r'baseline="([^"]+)"', line_attrs)
    spacing_match = re.search(r'spacing="([^"]+)"', line_attrs)
    vertsize_match = re.search(r'vertsize="([^"]+)"', line_attrs)
    textheight_match = re.search(r'textheight="([^"]+)"', line_attrs)

    horzsize = horz_match.group(1) if horz_match else '8044'
    flags = flags_match.group(1) if flags_match else '393216'
    baseline = baseline_match.group(1) if baseline_match else '935'
    spacing = spacing_match.group(1) if spacing_match else '276'
    vertsize = vertsize_match.group(1) if vertsize_match else '1100'
    textheight = textheight_match.group(1) if textheight_match else '1100'

    paragraphs = []
    for idx, line in enumerate(lines):
        vertpos = str(idx * 1376)
        prefix = prefix_runs if idx == 0 else ''
        paragraphs.append(
            f'<hp:p {p_attrs}>{prefix}<hp:run charPrIDRef="{char_id}"><hp:t>{escape(_sanitize_xml_text(line))}</hp:t></hp:run>'
            f'<hp:linesegarray><hp:lineseg textpos="0" vertpos="{vertpos}" vertsize="{vertsize}" '
            f'textheight="{textheight}" baseline="{baseline}" spacing="{spacing}" horzpos="0" '
            f'horzsize="{horzsize}" flags="{flags}"/></hp:linesegarray></hp:p>'
        )

    return xml[:match.start()] + ''.join(paragraphs) + xml[match.end():]


def _fill_content_change_template(xml: str, values: dict) -> str:
    root = ET.fromstring(xml)
    _replace_table_cell_text(root, 1, 1, values.get('광고주', ''))
    _replace_table_cell_text(root, 2, 1, values.get('표시위치', ''))
    _replace_table_cell_text(root, 3, 1, values.get('표시내용', ''))
    _replace_table_cell_text(root, 4, 1, values.get('규격', ''))
    _replace_table_cell_text(root, 5, 1, values.get('수량', ''))
    _set_content_change_opinion(
        root,
        '해당 광고물은 옥외광고물 관련 법령에 따른 기준에 적합하며, 표시 내용 또한 적합함.'
    )

    reviewer = _sanitize_xml_text(values.get('검토자', '').strip())
    for text_node in root.findall('.//hp:t', HWP_NS):
        if text_node.text == '[검토자 : ]':
            text_node.text = f'[검토자 : {reviewer}]'
            break

    return ET.tostring(root, encoding='unicode')


def _replace_table_cell_text(root, col: int, row: int, value: str):
    target_cell = None
    for cell in root.findall('.//hp:tc', HWP_NS):
        addr = cell.find('hp:cellAddr', HWP_NS)
        if addr is None:
            continue
        if addr.get('colAddr') == str(col) and addr.get('rowAddr') == str(row):
            target_cell = cell
            break
    if target_cell is None:
        return

    sublist = target_cell.find('hp:subList', HWP_NS)
    if sublist is None:
        return
    first_paragraph = sublist.find('hp:p', HWP_NS)
    if first_paragraph is None:
        return

    first_run = first_paragraph.find('hp:run', HWP_NS)
    line_seg = first_paragraph.find('hp:linesegarray/hp:lineseg', HWP_NS)
    if first_run is None or line_seg is None:
        return

    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if not lines:
        lines = ['']

    try:
        line_step = int(line_seg.get('vertsize', '1100')) + int(line_seg.get('spacing', '276'))
    except ValueError:
        line_step = 1376

    line_attrs = dict(line_seg.attrib)
    base_vertpos = int(line_attrs.get('vertpos', '0'))
    paragraph_attrs = dict(first_paragraph.attrib)
    char_id = first_run.get('charPrIDRef', '26')

    for child in list(sublist):
        sublist.remove(child)

    for idx, line in enumerate(lines):
        para_attrs = dict(paragraph_attrs)
        if idx > 0:
            para_attrs['id'] = '0'
        paragraph = ET.SubElement(sublist, _hp_tag('p'), para_attrs)
        run = ET.SubElement(paragraph, _hp_tag('run'), {'charPrIDRef': char_id})
        if line:
            text = ET.SubElement(run, _hp_tag('t'))
            text.text = _sanitize_xml_text(line)

        linesegarray = ET.SubElement(paragraph, _hp_tag('linesegarray'))
        current_line_attrs = dict(line_attrs)
        current_line_attrs['vertpos'] = str(base_vertpos + idx * line_step)
        ET.SubElement(linesegarray, _hp_tag('lineseg'), current_line_attrs)


def _set_content_change_opinion(root, text: str):
    paragraphs = root.findall('.//hp:p', HWP_NS)
    for idx, paragraph in enumerate(paragraphs):
        text_nodes = paragraph.findall('.//hp:t', HWP_NS)
        if ''.join(node.text or '' for node in text_nodes).strip() != '□ 검토 의견':
            continue
        if idx + 1 >= len(paragraphs):
            return

        target = paragraphs[idx + 1]
        run = target.find('hp:run', HWP_NS)
        if run is None:
            return

        for child in list(run):
            run.remove(child)
        opinion = ET.SubElement(run, _hp_tag('t'))
        opinion.text = _sanitize_xml_text(text)
        return


def _hp_tag(local_name: str) -> str:
    return f'{{{HWP_NS["hp"]}}}{local_name}'


def apply_style(root):
    style = ttk.Style(root)
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure('TLabelframe.Label', font=('Malgun Gothic', 9, 'bold'))
    style.configure('TButton', font=('Malgun Gothic', 9))
    style.configure('TLabel', font=('Malgun Gothic', 9))
    style.configure('TEntry', font=('Malgun Gothic', 9))
    style.configure('TCombobox', font=('Malgun Gothic', 9))
    style.configure('Accent.TButton',
                    font=('Malgun Gothic', 11, 'bold'),
                    foreground='white',
                    background='#1a73e8')


# ─────────────────────────────────────────────
# 엔트리포인트
# ─────────────────────────────────────────────

def main():
    app = AdReviewApp()
    apply_style(app)
    app.mainloop()


if __name__ == '__main__':
    main()
