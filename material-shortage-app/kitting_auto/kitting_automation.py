#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sMES 키팅 자재 자동 다운로드 & 자재부족현황 자동 업로드
========================================================
실행 전 준비:
  pip install pywinauto pyautogui pillow playwright psutil keyboard
  playwright install chromium
"""

import os, sys, time, ctypes, subprocess
from datetime import datetime
from pathlib import Path

# ──────────────────────────────────────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────────────────────────────────────
SMES_EXE     = Path(r"C:\Program Files (x86)\I2R\sMES\sMES.exe")

# 스크립트 위치 기준 상대 경로 (어떤 PC에서 실행해도 동작)
_HERE        = Path(__file__).parent          # kitting_auto 폴더
_BASE        = _HERE.parent                   # 프로젝트 루트

DOWNLOAD_DIR   = _BASE / "kitting 자재"
INVENTORY_DIR  = _BASE / "재고현황"
HTML_FILE      = _BASE / "자재부족현황.html"
LOG_FILE       = _HERE / "kitting_log.txt"

SMES_ID      = "SSAT045"
SMES_PW      = "rlatndus1!"

WEB_EMAIL    = "kulhyang0117@gmail.com"
WEB_PW       = "jxy0830!"

TODAY        = datetime.now().strftime("%Y-%m-%d")
TODAY_KR     = datetime.now().strftime("%Y%m%d")

STEP_DELAY   = 1.0
LOAD_DELAY   = 4.0
EXCEL_DELAY  = 5.0

# ── 외부 좌표 설정 파일 (kitting_config.json) ─────────────────────────────
# 앱 설정 > 자동화 좌표 설정 > "kitting_config.json 저장" 버튼으로 생성
import json as _json
_CFG_FILE = _HERE / "kitting_config.json"
_ext_cfg  = {}
if _CFG_FILE.exists():
    try:
        _ext_cfg = _json.loads(_CFG_FILE.read_text(encoding='utf-8'))
        print(f"[설정] kitting_config.json 로드: {_ext_cfg}")
    except Exception as _e:
        print(f"[설정] kitting_config.json 읽기 실패: {_e}")

COORD_EXCEL_X      = int(_ext_cfg.get('excel_btn_x',   308))
COORD_EXCEL_Y      = int(_ext_cfg.get('excel_btn_y',   533))
COORD_MENU_OFFSET  = int(_ext_cfg.get('menu_x_offset', 230))
COORD_ROW_Y_OFFSET = int(_ext_cfg.get('row_y_offset',  183))
COORD_ROW_HEIGHT   = int(_ext_cfg.get('row_height',     33))

# ── 경로/계정 오버라이드 (kitting_config.json 우선) ───────────────────────
if 'smes_exe'     in _ext_cfg and _ext_cfg['smes_exe']:
    SMES_EXE = Path(_ext_cfg['smes_exe'])
if 'download_dir' in _ext_cfg and _ext_cfg['download_dir']:
    DOWNLOAD_DIR = Path(_ext_cfg['download_dir'])
if 'inv_dir'      in _ext_cfg and _ext_cfg['inv_dir']:
    INVENTORY_DIR = Path(_ext_cfg['inv_dir'])
if 'prev_kit_dir' in _ext_cfg and _ext_cfg['prev_kit_dir']:
    _PREV_KIT_DIR_CFG = Path(_ext_cfg['prev_kit_dir'])
else:
    _PREV_KIT_DIR_CFG = None
if 'web_email'    in _ext_cfg and _ext_cfg['web_email']:
    WEB_EMAIL = _ext_cfg['web_email']
if 'web_pw'       in _ext_cfg and _ext_cfg['web_pw']:
    WEB_PW = _ext_cfg['web_pw']
if 'smes_id'      in _ext_cfg and _ext_cfg['smes_id']:
    SMES_ID = _ext_cfg['smes_id']
if 'smes_pw'      in _ext_cfg and _ext_cfg['smes_pw']:
    SMES_PW = _ext_cfg['smes_pw']

# 첫 저장 시 폴더 클리어 여부 (세션당 1회만)
_kitting_folder_cleared = False

# ──────────────────────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────────────────────
_log_file = None

def _init_log():
    global _log_file
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    _log_file = open(LOG_FILE, "a", encoding="utf-8")
    log(f"{'='*60}")
    log(f"로그 파일: {LOG_FILE}")

def log(msg):
    line = f"[{datetime.now():%H:%M:%S}] {msg}"
    try:
        print(line, flush=True)
    except UnicodeEncodeError:
        print(line.encode('utf-8', errors='replace').decode('ascii', errors='replace'), flush=True)
    if _log_file:
        _log_file.write(line + "\n")
        _log_file.flush()

def is_admin():
    try: return ctypes.windll.shell32.IsUserAnAdmin()
    except: return False

def elevate():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable,
        " ".join(f'"{a}"' for a in sys.argv),
        None, 1
    )
    sys.exit()

# ──────────────────────────────────────────────────────────────────────────────
# Step 1 : sMES 실행
# ──────────────────────────────────────────────────────────────────────────────
def launch_smes():
    log("▶ Step 1: sMES 실행 중...")
    if not SMES_EXE.exists():
        raise FileNotFoundError(f"sMES.exe 없음: {SMES_EXE}")
    subprocess.Popen([str(SMES_EXE)])
    log(f"  {LOAD_DELAY}초 로딩 대기...")
    time.sleep(LOAD_DELAY)

# ──────────────────────────────────────────────────────────────────────────────
# sMES 프로세스/창 찾기
# ──────────────────────────────────────────────────────────────────────────────
def _get_smes_pid():
    """실행 중인 sMES 프로세스 PID 반환"""
    import psutil
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            name = (proc.info['name'] or '').lower()
            exe  = (proc.info['exe']  or '').lower()
            if 'smes' in name or 'smes' in exe:
                return proc.info['pid']
        except Exception:
            pass
    return None

def _get_smes_window():
    """
    sMES의 최상위 창을 반환.
    win32 백엔드로 시도 → uia 백엔드로 재시도.
    """
    import pyautogui
    from pywinauto import Application, Desktop

    pid = _get_smes_pid()
    if not pid:
        return None, None

    for backend in ('win32', 'uia'):
        try:
            app = Application(backend=backend).connect(process=pid, timeout=5)
            # 최상위 창 중 가장 큰 창 선택
            wins = app.windows()
            if not wins:
                continue
            # 창 크기가 있는 것 우선
            wins_with_size = [(w, w.rectangle()) for w in wins]
            wins_with_size.sort(key=lambda x: x[1].width() * x[1].height(), reverse=True)
            win = wins_with_size[0][0]
            log(f"  창 연결 성공 (backend={backend}): '{win.window_text()}'")
            return app, win
        except Exception as e:
            log(f"  backend={backend} 연결 실패: {e}")

    return None, None

def _dump_window_info(win):
    """디버그용: 창의 자식 컨트롤 목록 출력"""
    try:
        log("  === 창 컨트롤 목록 (디버그) ===")
        for ctrl in win.children():
            try:
                log(f"    [{ctrl.element_info.control_type}] title='{ctrl.window_text()}' class='{ctrl.class_name()}'")
            except Exception:
                pass
    except Exception:
        pass

def _paste_text(text):
    """
    클립보드로 텍스트 붙여넣기 — 특수문자(!@#$ 등) 오입력 방지.
    pywin32의 win32clipboard 사용 (pywinauto 설치 시 함께 설치됨).
    """
    import win32clipboard
    import pyautogui

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()

    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.2)

# ──────────────────────────────────────────────────────────────────────────────
# Step 2 : sMES 로그인  (IMMES MEMBERS LOGIN 창)
# ──────────────────────────────────────────────────────────────────────────────
def _find_login_dialog(timeout=20):
    """
    sMES 프로세스의 모든 창을 순회해
    'WindowsForms10.EDIT.*' 컨트롤이 2개 이상인 창을 로그인 다이얼로그로 반환.
    타이틀바 없는 borderless form 도 탐지 가능.
    """
    from pywinauto import Application

    pid = _get_smes_pid()
    if not pid:
        raise RuntimeError("sMES 프로세스를 찾을 수 없습니다.")

    app = Application(backend='win32').connect(process=pid, timeout=10)
    log(f"  sMES PID={pid} 연결 완료")

    deadline = time.time() + timeout
    while time.time() < deadline:
        for w in app.windows():
            try:
                edits = w.children(class_name_re="WindowsForms10.EDIT.*")
                if len(edits) >= 2:
                    log(f"  로그인 폼 발견: title='{w.window_text()}' "
                        f"/ Edit {len(edits)}개")
                    return app, w
            except Exception:
                pass
        time.sleep(0.5)

    # 마지막 수단: 가장 작은 창 (로그인 다이얼로그는 메인보다 작음)
    wins = app.windows()
    if wins:
        smallest = min(wins,
                       key=lambda w: w.rectangle().width() * w.rectangle().height())
        log(f"  폼 자동 탐지 실패 → 가장 작은 창 사용: '{smallest.window_text()}'")
        return app, smallest

    raise RuntimeError("로그인 폼을 찾을 수 없습니다.")


def _click_field_and_input(field_ctrl, text, label):
    """
    C1TextBox 필드를 pyautogui 마우스 클릭 → 전체선택 삭제 → 클립보드 붙여넣기.
    좌표는 컨트롤 rectangle() 중앙 사용.
    """
    import pyautogui, win32clipboard

    rect = field_ctrl.rectangle()
    cx = (rect.left + rect.right) // 2
    cy = (rect.top + rect.bottom) // 2

    log(f"  {label} 필드 클릭: ({cx}, {cy})")

    # 1) 마우스로 정확히 클릭
    pyautogui.click(cx, cy)
    time.sleep(0.4)

    # 2) 전체 선택 후 삭제
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.2)
    pyautogui.press('delete')
    time.sleep(0.2)

    # 3) 클립보드로 붙여넣기
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.4)

    log(f"  {label} 입력 완료")


def login_smes(_unused_win=None):
    """
    ① PID 로 sMES 연결
    ② Edit 컨트롤 2개 있는 창 = 로그인 폼
    ③ Y좌표 정렬 → 위=ID, 아래=PW
    ④ pyautogui 마우스 클릭 + 클립보드 붙여넣기 (C1TextBox 호환)
    ⑤ Login 버튼 클릭
    """
    import pyautogui

    log("  로그인 폼 탐색 중 (PID 방식)...")
    app, dlg = _find_login_dialog(timeout=20)

    # 창을 맨 앞으로 활성화
    try:
        dlg.set_focus()
    except Exception:
        pass
    try:
        import ctypes
        ctypes.windll.user32.SetForegroundWindow(dlg.handle)
    except Exception:
        pass
    time.sleep(0.5)

    # ── Edit 컨트롤 수집 ────────────────────────────────────────────────────
    all_edits = dlg.children(class_name_re="WindowsForms10.EDIT.*")
    log(f"  Edit 컨트롤 전체 {len(all_edits)}개")

    BUTTON_TEXTS = {'exit', 'login', '로그인', '취소', 'cancel', 'ok', '확인'}
    id_field = None
    pw_field = None

    # 1순위: 텍스트로 직접 탐지 — 크기 필터 없이 전체에서 검색
    for e in all_edits:
        try:
            txt = e.window_text().strip()
            log(f"    필드 '{txt}' 위치: {e.rectangle()}")
            if txt.lower() in BUTTON_TEXTS:
                continue
            if txt == SMES_ID:
                id_field = e
            elif txt.upper() == 'PASSWORD':
                pw_field = e
        except Exception:
            pass

    # 2순위: 크기 필터 후 Y 정렬 fallback
    if id_field is None or pw_field is None:
        small_edits = [
            e for e in all_edits
            if e.rectangle().width() < 500 and e.rectangle().height() < 60
        ]
        log(f"  크기 필터 필드 {len(small_edits)}개")
        input_only = [
            e for e in small_edits
            if e.window_text().strip().lower() not in BUTTON_TEXTS
        ]
        sorted_edits = sorted(input_only, key=lambda c: c.rectangle().top)
        if id_field is None and len(sorted_edits) >= 1:
            id_field = sorted_edits[0]
        if pw_field is None and len(sorted_edits) >= 2:
            pw_field = sorted_edits[1]

    # 3순위: 크기 필터 없이 Y 정렬 (최후 수단)
    if id_field is None or pw_field is None:
        input_only_all = [
            e for e in all_edits
            if e.window_text().strip().lower() not in BUTTON_TEXTS
        ]
        sorted_all = sorted(input_only_all, key=lambda c: c.rectangle().top)
        if id_field is None and len(sorted_all) >= 1:
            id_field = sorted_all[0]
        if pw_field is None and len(sorted_all) >= 2:
            pw_field = sorted_all[1]

    if id_field is None or pw_field is None:
        raise RuntimeError(f"ID/PW 입력 필드를 찾지 못했습니다. Edit 전체 수: {len(all_edits)}")

    log(f"  ID 필드 위치: {id_field.rectangle()}")
    log(f"  PW 필드 위치: {pw_field.rectangle()}")

    # ── ① ID 입력 ────────────────────────────────────────────────────────
    _click_field_and_input(id_field, SMES_ID, "① ID")
    time.sleep(0.3)

    # ── ② PW 입력 ────────────────────────────────────────────────────────
    _click_field_and_input(pw_field, SMES_PW, "② PW")
    time.sleep(0.3)

    # ── ③ Login 버튼 클릭 ────────────────────────────────────────────────
    login_btn = None
    for btn in dlg.children(class_name_re="WindowsForms10.BUTTON.*"):
        try:
            if btn.window_text().strip().lower() in ('login', '로그인'):
                login_btn = btn
                break
        except Exception:
            pass

    if login_btn:
        rect = login_btn.rectangle()
        bx = (rect.left + rect.right) // 2
        by = (rect.top + rect.bottom) // 2
        log(f"  ③ Login 버튼 클릭: ({bx}, {by})")
        pyautogui.click(bx, by)
    else:
        log("  ③ Login 버튼 미발견 → Enter 키 입력")
        pyautogui.press('enter')
    time.sleep(0.5)

    # ── 로그인 성공 여부 검증 ────────────────────────────────────────────────
    # Login 버튼이 실제로 사라졌는지 확인 (창 핸들 깜빡임 오감지 방지)
    log("  로그인 결과 확인 중...")
    deadline = time.time() + 12
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            # 현재 창 상태 재취득 후 Login 버튼 존재 여부 확인
            _, cur_win = _get_smes_window()
            if cur_win is None:
                log("  ✅ 로그인 성공 (창 재구성)")
                time.sleep(LOAD_DELAY)
                return
            btns = cur_win.children(class_name_re="WindowsForms10.BUTTON.*")
            login_btn_visible = any(
                b.window_text().strip().lower() in ('login', '로그인')
                and b.is_visible()
                for b in btns
            )
            if not login_btn_visible:
                log("  ✅ 로그인 성공 (Login 버튼 사라짐)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            pass

        # 에러 팝업 메시지 확인
        try:
            for child in dlg.children():
                txt = child.window_text().strip()
                if txt and txt not in (SMES_ID, 'PASSWORD', '') and len(txt) < 100:
                    log(f"  ⚠️  화면 메시지: '{txt}'")
        except Exception:
            pass

    # 12초 후에도 Login 버튼이 남아있으면 실패 → 중단
    raise RuntimeError(
        "❌ 로그인 실패: Login 버튼이 사라지지 않았습니다.\n"
        "  → ID/PW 확인 후 다시 실행해주세요."
    )

# ──────────────────────────────────────────────────────────────────────────────
# Step 3~5 : 메뉴 이동 → 조회 → Excel 다운로드
# ──────────────────────────────────────────────────────────────────────────────
def navigate_and_download(app):
    """
    로그인 후 메인 창을 다시 가져와 메뉴/버튼 클릭.
    """
    import pyautogui
    from pywinauto import Desktop

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    downloaded_files = []

    # 메인 창 재취득 (로그인 후 새 창)
    time.sleep(LOAD_DELAY)
    _, main_win = _get_smes_window()
    if not main_win:
        log("  ⚠️  메인 창을 찾을 수 없습니다. 수동으로 진행해주세요.")
        input("  조회 완료 후 Enter → ")
        _, main_win = _get_smes_window()

    main_win.set_focus()
    time.sleep(0.5)

    # ── 생산관리 클릭 (다중 방법) ───────────────────────────────────────────
    log("  생산관리 클릭...")
    clicked = False

    # Method 1: pywinauto menu_select
    try:
        main_win.menu_select("생산관리")
        clicked = True
        log("  ✅ menu_select 성공")
    except Exception as e:
        log(f"  menu_select 실패: {e}")

    # Method 2: UIA 백엔드로 MenuItem 탐색
    if not clicked:
        try:
            from pywinauto import Application as _App
            _pid = main_win.process_id()
            _uia = _App(backend='uia').connect(process=_pid, timeout=5)
            for _w in _uia.windows():
                try:
                    for _d in _w.descendants(control_type="MenuItem"):
                        if "생산관리" in (_d.window_text() or ""):
                            _d.click_input()
                            clicked = True
                            log("  ✅ UIA MenuItem 성공")
                            break
                except Exception:
                    pass
                if clicked:
                    break
        except Exception as e:
            log(f"  UIA MenuItem 실패: {e}")

    # Method 3: 기존 child_window 탐색
    if not clicked:
        clicked = _try_click(main_win, ["생산관리"])
        if clicked:
            log("  ✅ child_window 탐색 성공")

    # Method 4: 창 좌표 기준 메뉴바 클릭 (4번째 항목)
    if not clicked:
        try:
            rect = main_win.rectangle()
            # 최대화 창은 rect.top이 -8 정도로 음수 → max()로 보정
            menu_y = max(rect.top + 50, 42)
            menu_x = rect.left + COORD_MENU_OFFSET   # System+기본정보관리+영업관리 이후 위치
            pyautogui.click(menu_x, menu_y)
            clicked = True
            log(f"  ✅ 좌표 클릭 성공: ({menu_x}, {menu_y})")
        except Exception as e:
            log(f"  좌표 클릭 실패: {e}")

    if not clicked:
        log("  ⚠️  수동으로 '생산관리'를 클릭해주세요.")
        input("  완료 후 Enter → ")

    time.sleep(STEP_DELAY)  # 드롭다운 열릴 때까지 대기

    # ── 방향키 ↓ 8번 → Enter (조립 자재 kitting 진입) ────────────────────
    log("  방향키 ↓ 8번 → Enter (조립 자재 kitting 진입)...")
    for _ in range(8):
        pyautogui.press('down')
        time.sleep(0.15)
    pyautogui.press('enter')
    time.sleep(LOAD_DELAY)

    # ── 생산일자 설정 ──────────────────────────────────────────────────────
    log(f"  생산일자 설정: {TODAY}")
    set_date(main_win)

    # ── 조회 클릭 ─────────────────────────────────────────────────────────
    log("  조회 버튼 클릭...")
    import pyautogui as _pag
    clicked_search = False
    # 1) child_window 탐색
    if _try_click(main_win, ["조회", "검색", "Search"]):
        clicked_search = True
        log("  ✅ 조회 버튼 클릭 (child_window)")
    # 2) F5 단축키 (sMES 조회 표준 단축키)
    if not clicked_search:
        _pag.press('f5')
        clicked_search = True
        log("  ✅ 조회 F5 키 입력")
    time.sleep(LOAD_DELAY)
    log("  ✅ 조회 완료")

    # 수동 조작 등으로 창 핸들이 스테일해질 수 있으므로 재취득
    _, main_win = _get_smes_window()
    if not main_win:
        raise RuntimeError("조회 후 메인 창을 찾을 수 없습니다.")

    # ── 품목별 다운로드 ────────────────────────────────────────────────────
    log("  품목별 Excel 다운로드 시작...")
    downloaded_files = download_all_items(main_win)

    return downloaded_files


def set_date(win):
    import pyautogui
    # DateTimePicker 또는 C1DateEdit 시도
    for ct in ["DateTimePicker", "Edit", "Custom"]:
        try:
            ctrl = win.child_window(control_type=ct, found_index=0)
            if ctrl.exists(timeout=1):
                ctrl.set_focus()
                time.sleep(0.2)
                ctrl.type_keys("^a{DEL}")
                time.sleep(0.1)
                ctrl.type_keys(TODAY_KR)
                time.sleep(0.3)
                return
        except Exception:
            pass
    # 키보드 직접
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(TODAY_KR, interval=0.05)
    log(f"  날짜 키보드 입력: {TODAY_KR}")


def _try_click(win, candidates):
    """텍스트/버튼/메뉴/트리 등 모든 방식으로 클릭 시도"""
    for name in candidates:
        for ct in ["Button", "MenuItem", "TreeItem", "ListItem", "Text", "Custom", None]:
            try:
                kwargs = {"title_re": f".*{name}.*"}
                if ct:
                    kwargs["control_type"] = ct
                ctrl = win.child_window(**kwargs)
                if ctrl.exists(timeout=1):
                    ctrl.click_input()
                    return True
            except Exception:
                pass
    return False


def _uia_selected_row_rect(win):
    """
    UIA로 현재 선택/포커스된 그리드 행의 bounding rectangle 반환.
    DataItem → ListItem → Custom 순으로 시도.
    성공 시 pywinauto RECT 반환, 실패 시 None.
    """
    try:
        from pywinauto import Application
        pid = win.process_id()
        uia_app = Application(backend='uia').connect(process=pid, timeout=3)
        for w in uia_app.windows():
            try:
                for ct in ('DataItem', 'ListItem', 'Custom'):
                    for row in w.descendants(control_type=ct):
                        try:
                            selected = False
                            try:
                                selected = row.is_selected()
                            except Exception:
                                pass
                            if not selected:
                                try:
                                    selected = row.has_keyboard_focus()
                                except Exception:
                                    pass
                            if selected:
                                rect = row.rectangle()
                                if rect.width() > 50 and rect.height() > 5:
                                    return rect
                        except Exception:
                            pass
            except Exception:
                pass
    except Exception:
        pass
    return None


def _find_selected_row_y(col_x, grid_top=150, grid_bottom=560, row_height=22):
    """
    그리드에서 현재 선택된(하이라이트) 행의 Y 좌표를 색상으로 탐지.
    Windows 기본 선택색(파란 계열) 기준으로 열 X를 세로 스캔.
    """
    import pyautogui

    h = grid_bottom - grid_top
    region = (col_x - 5, grid_top, 10, h)
    try:
        img = pyautogui.screenshot(region=region)
    except Exception:
        return None

    row_scores = {}
    for y in range(h):
        row_idx = y // row_height
        try:
            pixel = img.getpixel((5, y))
            r, g, b = pixel[0], pixel[1], pixel[2]
        except Exception:
            continue
        blue_score = b - (r + g) / 2
        row_scores.setdefault(row_idx, []).append(blue_score)

    if not row_scores:
        return None

    best_row, scores = max(row_scores.items(), key=lambda kv: sum(kv[1]) / len(kv[1]))
    avg = sum(scores) / len(scores)
    if avg < 20:
        return None

    return grid_top + best_row * row_height + row_height // 2


def download_all_items(win):
    import pyautogui

    # 그리드 첫 번째 행만 클릭 (선택 상태 만들기)
    # children()은 화면에 보이는 행만 반환하므로 전체 순회에 사용 불가
    for ct in ["DataItem", "ListItem", "TreeItem", "Custom"]:
        try:
            candidates = win.children(control_type=ct)
            if len(candidates) > 0:
                log(f"  그리드 첫 행 클릭 (type={ct}, 화면 표시 {len(candidates)}개)")
                candidates[0].click_input()
                time.sleep(0.5)
                break
        except Exception:
            pass

    # 키보드 Down 방식으로 전체 품목 순환 — 스크롤도 자동 처리됨
    log("  키보드 Down 방식으로 전체 품목 다운로드 시작...")
    downloaded_files = _download_by_keyboard(win)

    log(f"  ✅ {len(downloaded_files)}개 파일 다운로드 완료")
    return downloaded_files


def _click_excel_btn(win, row_y=None):
    """
    Excel 버튼 탐색 및 클릭.
    0순위: pnl_top_button 중앙 클릭 (조회/Excel/종료 3버튼 중 2번째)
    1순위: win32 child_window  2순위: UIA 부분텍스트/automation_id 스캔
    3순위: ToolBar 항목 스캔  4순위: win32 완전재귀탐색  5순위: 우클릭
    성공 시 True 반환.
    """
    import pyautogui, win32gui
    EXCEL_CONTAINS = ['excel', '엑셀', 'xls', 'export', '다운로드']

    # 0) Excel 버튼 절대 좌표 클릭 (kitting_config.json 또는 기본값)
    pyautogui.click(COORD_EXCEL_X, COORD_EXCEL_Y)
    log(f"    Excel 버튼 좌표 클릭: ({COORD_EXCEL_X}, {COORD_EXCEL_Y})")
    return True

    # 1) win32 backend
    if _try_click(win, ["Excel 다운로드", "Excel", "엑셀", "Export", "EXCEL", "다운로드"]):
        log("    Excel 버튼 클릭 (win32)")
        return True

    # 2) UIA descendants 전체 스캔 — 부분 텍스트 / automation_id / name 포함 검사
    _dump_controls = []   # 진단용 컨트롤 목록
    try:
        from pywinauto import Application
        pid = win.process_id()
        uia_app = Application(backend='uia').connect(process=pid, timeout=5)
        for w in uia_app.windows():
            try:
                for ctrl in w.descendants():
                    try:
                        txt  = (ctrl.window_text() or '').strip().lower()
                        name = ''
                        aid  = ''
                        ctype = ''
                        try:
                            name  = (ctrl.element_info.name or '').lower()
                            aid   = (ctrl.element_info.automation_id or '').lower()
                            ctype = (ctrl.element_info.control_type or '')
                        except Exception:
                            pass
                        combined = f"{txt} {name} {aid}"
                        # 진단 목록에 추가 (Button/MenuItem/Custom 계열만)
                        if ctype in ('Button', 'MenuItem', 'SplitButton', 'Custom') or any(k in combined for k in ['btn', 'button', 'menu', 'tool']):
                            _dump_controls.append(f"[{ctype}] txt='{txt}' name='{name}' aid='{aid}'")
                        if any(k in combined for k in EXCEL_CONTAINS):
                            ctrl.click_input()
                            log(f"    Excel 버튼 클릭 (UIA 부분일치: '{txt or name}')")
                            return True
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception as e:
        log(f"    UIA 스캔 실패: {e}")

    # 3) ToolBar / ToolStripButton 항목 스캔 (아이콘 전용 버튼 대응)
    try:
        from pywinauto import Application
        pid = win.process_id()
        uia_app = Application(backend='uia').connect(process=pid, timeout=5)
        for w in uia_app.windows():
            try:
                toolbars = w.descendants(control_type="ToolBar")
                for tb in toolbars:
                    try:
                        items = tb.descendants(control_type="Button")
                        for item in items:
                            try:
                                iname = (item.element_info.name or '').lower()
                                iaid  = (item.element_info.automation_id or '').lower()
                                itxt  = (item.window_text() or '').strip().lower()
                                combined = f"{iname} {iaid} {itxt}"
                                _dump_controls.append(f"[ToolBar/Button] txt='{itxt}' name='{iname}' aid='{iaid}'")
                                if any(k in combined for k in EXCEL_CONTAINS):
                                    item.click_input()
                                    log(f"    Excel 버튼 클릭 (ToolBar: '{iname or iaid}')")
                                    return True
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception as e:
        log(f"    ToolBar 스캔 실패: {e}")

    # 4) win32 EnumWindows + PID 필터 — 프로세스 전체 버튼 완전 재귀 탐색
    #    (콜백이 반드시 True/False 반환 → EnumChildWindows 정상 동작)
    try:
        import win32gui, win32con, win32process
        _excel_hwnd = [None]
        _pid = win.process_id()
        _all_btns = []   # 진단용: 발견된 모든 버튼 텍스트

        def _enum_cb(hwnd, _):
            """EnumChildWindows 콜백: True=계속, False=중단(찾음)"""
            _scan_hwnd(hwnd)
            return not bool(_excel_hwnd[0])

        def _scan_hwnd(hwnd):
            """hwnd 및 모든 자손을 재귀 탐색.
            C1Input.C1Button  → 클래스에 'BUTTON' 포함
            C1Command 리본버튼 → 클래스에 'WINDOW'/'FORMS' 포함, BUTTON 미포함
            두 유형 모두 탐색.
            """
            if _excel_hwnd[0]:
                return
            try:
                cls = win32gui.GetClassName(hwnd)
                txt = win32gui.GetWindowText(hwnd) or ''
                cls_up = cls.upper()
                # WinForms 계열 컨트롤은 모두 검사 (BUTTON + Window 형 모두)
                if 'WINDOWSFORMS10' in cls_up or 'BUTTON' in cls_up:
                    if txt:
                        _all_btns.append(f"[{cls[:30]}] '{txt}'")
                    if any(k in txt.lower() for k in EXCEL_CONTAINS):
                        _excel_hwnd[0] = hwnd
                        return
            except Exception:
                pass
            try:
                win32gui.EnumChildWindows(hwnd, _enum_cb, None)
            except Exception:
                pass

        def _enum_top_cb(hwnd, _):
            if _excel_hwnd[0]:
                return False
            try:
                _, h_pid = win32process.GetWindowThreadProcessId(hwnd)
                if h_pid == _pid:
                    _scan_hwnd(hwnd)
            except Exception:
                pass
            return not bool(_excel_hwnd[0])

        win32gui.EnumWindows(_enum_top_cb, None)

        if _excel_hwnd[0]:
            btn_txt = win32gui.GetWindowText(_excel_hwnd[0])
            btn_cls = win32gui.GetClassName(_excel_hwnd[0]).upper()
            if 'BUTTON' in btn_cls:
                # 표준 버튼: BM_CLICK
                win32gui.SendMessage(_excel_hwnd[0], win32con.BM_CLICK, 0, 0)
            else:
                # C1Command 리본 버튼: 좌표 기반 마우스 클릭
                rc = win32gui.GetWindowRect(_excel_hwnd[0])
                cx, cy = (rc[0] + rc[2]) // 2, (rc[1] + rc[3]) // 2
                pyautogui.click(cx, cy)
            log(f"    Excel 버튼 클릭 (win32 완전재귀탐색: '{btn_txt}')")
            return True
        else:
            log(f"    [win32 버튼 진단] 발견된 버튼 {len(_all_btns)}개: {_all_btns[:50]}")
    except Exception as e:
        log(f"    win32 완전재귀탐색 실패: {e}")

    # 5) 우클릭 컨텍스트 메뉴 (최후 수단 — 느림)
    if row_y is not None:
        try:
            from pywinauto import Desktop
            log(f"    우클릭 컨텍스트 메뉴 시도: (400, {row_y})")
            pyautogui.rightClick(400, row_y)
            time.sleep(0.8)
            desktop = Desktop(backend='uia')
            for _ in range(5):
                try:
                    menus = desktop.windows(control_type='Menu')
                    for menu in menus:
                        for item in menu.descendants(control_type='MenuItem'):
                            itxt = (item.window_text() or '').strip().lower()
                            iname = (item.element_info.name or '').lower()
                            if any(k in f"{itxt} {iname}" for k in EXCEL_CONTAINS):
                                item.click_input()
                                log(f"    Excel 컨텍스트 메뉴 클릭: '{itxt or iname}'")
                                return True
                except Exception:
                    pass
                time.sleep(0.3)
            pyautogui.press('escape')
            time.sleep(0.3)
        except Exception as e:
            log(f"    우클릭 시도 실패: {e}")

    if _dump_controls:
        log(f"    [진단] UIA 컨트롤 목록 ({len(_dump_controls)}개, 최대 60개 출력):")
        for line in _dump_controls[:60]:
            log(f"      {line}")
    else:
        log("    [진단] UIA에서 발견된 Button/MenuItem 컨트롤 없음")

    return False


def _click_excel_download(win, idx, item_name, row_y=None):
    """Excel 다운로드 버튼 클릭 및 저장"""
    import pyautogui
    from pywinauto import Desktop

    # Excel 버튼 클릭 — 저장 다이얼로그 미감지 시 1회 재시도
    save_path = str(DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}_{idx:03d}.xlsx")

    for _attempt in range(2):
        if not _click_excel_btn(win, row_y=row_y):
            log(f"    ⚠️  [{idx}] Excel 버튼 미발견 — 건너뜀")
            return None

        time.sleep(EXCEL_DELAY)

        # 저장 다이얼로그 처리
        if _handle_save_dialog(save_path):
            log(f"    ✅ 저장: {Path(save_path).name}")
            pyautogui.hotkey('ctrl', 'w')   # Excel 현재 창 닫기
            time.sleep(0.5)
            return save_path

        # 다이얼로그 없이 자동 저장된 경우 — Downloads 폴더에서 이동
        latest = _find_latest_download()
        if latest:
            import shutil
            dest = DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}_{idx:03d}.xlsx"
            shutil.move(str(latest), str(dest))
            log(f"    ✅ 이동: {dest.name}")
            pyautogui.hotkey('ctrl', 'w')   # Excel 현재 창 닫기
            time.sleep(0.5)
            return str(dest)

        if _attempt == 0:
            log(f"    저장 다이얼로그 미감지 — Excel 버튼 재클릭 시도...")
            time.sleep(1.0)

    log(f"    ⚠️  [{idx}] 파일 저장 실패 — 건너뜀")
    return None


def _download_by_keyboard(win):
    import pyautogui

    win.set_focus(); time.sleep(0.3)

    # ── UIA 'row N' 요소 직접 수집 (C1TrueDBGrid 행 노출, 최대 3회 재시도) ──
    _uia_rows = {}   # {row_idx: elem}
    for _attempt in range(3):
        try:
            from pywinauto import Application as _UA
            _ua = _UA(backend='uia').connect(process=win.process_id(), timeout=5)
            for _uw in _ua.windows():
                try:
                    for _elem in _uw.descendants(control_type='Custom'):
                        try:
                            _name = (_elem.element_info.name or '')
                            if _name.startswith('row ') and _name[4:].isdigit():
                                _idx = int(_name[4:])
                                if _idx not in _uia_rows:
                                    _uia_rows[_idx] = _elem
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception as _e:
            log(f"  UIA row 수집 실패 (시도 {_attempt+1}): {_e}")
        if _uia_rows:
            break
        log(f"  UIA row 미발견 — 1초 대기 후 재시도 ({_attempt+1}/3)...")
        time.sleep(1)

    downloaded = []

    if _uia_rows:
        # ── UIA row 직접 클릭 방식 ──────────────────────────────────────────
        log(f"  UIA row 요소 {len(_uia_rows)}개 발견 → 직접 클릭 방식")
        for i in sorted(_uia_rows.keys()):
            elem = _uia_rows[i]
            log(f"  [{i+1}] row {i} 클릭 → Excel 다운로드...")
            try:
                rect = elem.rectangle()
                row_x = (rect.left + rect.right) // 2
                row_y = (rect.top + rect.bottom) // 2
                pyautogui.click(row_x, row_y)
            except Exception:
                try:
                    elem.click_input()
                    rect = elem.rectangle()
                    row_x = (rect.left + rect.right) // 2
                    row_y = (rect.top + rect.bottom) // 2
                except Exception:
                    row_x, row_y = 800, 205
            time.sleep(0.3)

            saved = _click_excel_download(win, i + 1, f"kitting_{i+1:03d}", row_y=row_y)
            if saved:
                downloaded.append(saved)

            try:
                win.set_focus()
            except Exception:
                _, win = _get_smes_window()
                if not win:
                    log("  ⚠️  창 재취득 실패 — 루프 종료")
                    break
                try:
                    win.set_focus()
                except Exception:
                    pass
            time.sleep(0.3)

        log(f"  ✅ {len(downloaded)}개 파일 다운로드 완료")
        return downloaded

    # ── fallback: 키보드 Down 방식 ─────────────────────────────────────────
    log("  UIA row 미발견 → 키보드 Down 방식 fallback")

    try:
        _wrect = win.rectangle()
        ROW_X = _wrect.left + int((_wrect.right - _wrect.left) * 0.35)
    except Exception:
        ROW_X = 800
    ROW_HEIGHT = COORD_ROW_HEIGHT  # kitting_config.json 또는 기본값(33, 2560x1600 기준)

    try:
        _wrect = win.rectangle()
        _row_y = _wrect.top + COORD_ROW_Y_OFFSET
    except Exception:
        _row_y = COORD_ROW_Y_OFFSET
    pyautogui.click(ROW_X, _row_y)
    log(f"  첫 번째 행 좌표 클릭: ({ROW_X}, {_row_y})")
    time.sleep(0.5)
    last_sel_y = _row_y

    # 상태바에서 총 행 수 읽기 ("조회 결과 : N Rows")
    import re as _re
    _total_rows = None
    try:
        def _status_cb(hwnd, _):
            try:
                txt = win32gui.GetWindowText(hwnd)
                m = _re.search(r'(\d+)\s*Rows?', txt, _re.IGNORECASE)
                if m:
                    _total_rows_ref[0] = int(m.group(1))
            except Exception:
                pass
            return True
        _total_rows_ref = [None]
        win32gui.EnumChildWindows(win.handle, _status_cb, None)
        _total_rows = _total_rows_ref[0]
        if _total_rows is not None:
            log(f"  총 행 수 감지: {_total_rows}개")
    except Exception:
        pass

    screen_w, screen_h = pyautogui.size()
    region_top = 140
    region_h   = screen_h - region_top - 50
    grid_bottom = screen_h - 100
    region = (max(ROW_X - 200, 0), region_top, 500, region_h)

    def _advance_and_check():
        """현재 행 클릭 → ↓ → 화면 비교. True=이동됨, False=마지막 행"""
        nonlocal last_sel_y, win
        try:
            win.set_focus()
        except Exception:
            _, win = _get_smes_window()
            if not win:
                return None  # 창 없음
            try:
                win.set_focus()
            except Exception:
                pass
        time.sleep(0.4)
        pyautogui.click(ROW_X, last_sel_y)
        time.sleep(0.3)
        before = pyautogui.screenshot(region=region)
        pyautogui.press('down')
        last_sel_y = min(last_sel_y + ROW_HEIGHT, grid_bottom - ROW_HEIGHT)
        time.sleep(0.5)
        after = pyautogui.screenshot(region=region)
        return list(before.getdata()) != list(after.getdata())

    # 첫 번째 행 다운로드
    log(f"  [1] Excel 다운로드 시도...")
    saved = _click_excel_download(win, 1, "kitting_001", row_y=last_sel_y)
    if saved:
        downloaded.append(saved)

    for i in range(1, 500):
        # 총 행 수를 알면 초과 시 즉시 종료
        if _total_rows is not None and i >= _total_rows:
            log(f"  ✅ 전체 {_total_rows}개 행 완료")
            break

        # 다음 행으로 이동 — 이동 안 되면 마지막 행
        moved = _advance_and_check()
        if moved is None:
            log("  ⚠️  창 재취득 실패 — 루프 종료")
            break
        if not moved:
            log(f"  ✅ 마지막 행 도달 — 전체 {len(downloaded)}개 다운로드 완료")
            break

        log(f"  [{i+1}] Excel 다운로드 시도...")
        saved = _click_excel_download(win, i + 1, f"kitting_{i+1:03d}", row_y=last_sel_y)
        if saved:
            downloaded.append(saved)
    else:
        log(f"  ✅ 최대 반복 도달 — 전체 {len(downloaded)}개 다운로드 완료")

    return downloaded


def _handle_save_dialog(save_path):
    """Windows 저장 다이얼로그 자동 처리 — win32gui 탐지"""
    import pyautogui, pyperclip
    import win32gui, win32con
    global _kitting_folder_cleared

    target_dir = str(Path(save_path).parent)

    # 첫 저장 전 폴더 기존 파일 전체 삭제 (세션당 1회)
    if not _kitting_folder_cleared:
        _dir = Path(target_dir)
        for _f in list(_dir.glob("*.xlsx")) + list(_dir.glob("*.xls")):
            try:
                _f.unlink()
            except Exception:
                pass
        _kitting_folder_cleared = True
        log("    kitting 자재 폴더 기존 파일 삭제 완료")

    _SAVE_KEYS = ['저장', 'Save As', '다른 이름', 'xlsx', 'xls']

    def _find_save_hwnd():
        found = []
        def _cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                t = win32gui.GetWindowText(hwnd)
                if any(k in t for k in _SAVE_KEYS):
                    found.append(hwnd)
            return True
        win32gui.EnumWindows(_cb, None)
        return found[0] if found else None

    # 다이얼로그 탐지 대기 (최대 10초)
    dlg_hwnd = None
    for _ in range(20):
        dlg_hwnd = _find_save_hwnd()
        if dlg_hwnd:
            break
        time.sleep(0.5)

    if not dlg_hwnd:
        log("    저장 다이얼로그 미감지 (win32gui)")
        return False

    log(f"    저장 다이얼로그 감지: '{win32gui.GetWindowText(dlg_hwnd)}'")

    try:
        # 포커스 이동
        win32gui.ShowWindow(dlg_hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(dlg_hwnd)
        time.sleep(0.4)

        # 파일명 입력란에 전체 경로 붙여넣기
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyperclip.copy(save_path)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1.5)

        # 덮어쓰기 확인 팝업 (win32gui로 탐지)
        conf_hwnd = _find_save_hwnd()
        if conf_hwnd and conf_hwnd != dlg_hwnd:
            win32gui.SetForegroundWindow(conf_hwnd)
            time.sleep(0.2)
            pyautogui.press('enter')

        return True

    except Exception as e:
        log(f"    저장 다이얼로그 처리 오류: {e}")
        return False


def _find_latest_download():
    """Downloads 폴더 최근 xlsx 파일"""
    dl_dir = Path.home() / "Downloads"
    files = list(dl_dir.glob("*.xlsx")) + list(dl_dir.glob("*.xls"))
    if not files:
        return None
    f = max(files, key=lambda x: x.stat().st_mtime)
    # 5초 이내에 생성된 파일만
    if time.time() - f.stat().st_mtime < 30:
        return f
    return None


def _safe(name):
    import re
    return re.sub(r'[\\/:*?"<>|]', '_', name)


# ──────────────────────────────────────────────────────────────────────────────
# Step 5b : 현황조회 → 창고별 부품 현재고 조회 → Excel 다운로드
# ──────────────────────────────────────────────────────────────────────────────
def navigate_and_download_inventory(app):
    """
    현황조회 메뉴(생산관리 오른쪽 5번째) → ↓7 → Enter
    → Tab×3 → 수원부품창고 → Enter
    → Tab×7 → Enter
    → Tab×1 → Excel 저장 → 재고현황 폴더
    """
    import pyautogui

    log("▶ Step 5b: 재고현황 다운로드...")
    INVENTORY_DIR.mkdir(parents=True, exist_ok=True)

    # 메인 창 재취득
    time.sleep(LOAD_DELAY)
    _, main_win = _get_smes_window()
    if not main_win:
        log("  ⚠️  메인 창 없음. 수동으로 열어주세요.")
        input("  열린 후 Enter → ")
        _, main_win = _get_smes_window()

    main_win.set_focus()
    time.sleep(0.5)

    # ── 현황조회 클릭 (다중 방법) ────────────────────────────────────────────
    log("  현황조회 클릭...")
    clicked = False

    try:
        main_win.menu_select("현황조회")
        clicked = True
        log("  ✅ menu_select 성공")
    except Exception as e:
        log(f"  menu_select 실패: {e}")

    if not clicked:
        try:
            from pywinauto import Application as _App
            _pid = main_win.process_id()
            _uia = _App(backend='uia').connect(process=_pid, timeout=5)
            for _w in _uia.windows():
                try:
                    for _d in _w.descendants(control_type="MenuItem"):
                        if "현황조회" in (_d.window_text() or ""):
                            _d.click_input()
                            clicked = True
                            log("  ✅ UIA MenuItem 성공")
                            break
                except Exception:
                    pass
                if clicked:
                    break
        except Exception as e:
            log(f"  UIA MenuItem 실패: {e}")

    if not clicked:
        clicked = _try_click(main_win, ["현황조회"])
        if clicked:
            log("  ✅ child_window 탐색 성공")

    if not clicked:
        try:
            rect = main_win.rectangle()
            # 최대화 창은 rect.top이 -8 정도로 음수 → max()로 보정
            menu_y = max(rect.top + 50, 42)
            # 생산관리(230) 오른쪽 5번째 — 각 메뉴 약 80px 간격
            menu_x = rect.left + 630
            pyautogui.click(menu_x, menu_y)
            clicked = True
            log(f"  ✅ 좌표 클릭: ({menu_x}, {menu_y})")
        except Exception as e:
            log(f"  좌표 클릭 실패: {e}")

    if not clicked:
        log("  ⚠️  수동으로 '현황조회'를 클릭해주세요.")
        input("  완료 후 Enter → ")

    time.sleep(STEP_DELAY)

    # ── ↓ 7번 → Enter (창고별 부품 현재고 조회) ──────────────────────────────
    log("  ↓ 7번 → Enter (창고별 부품 현재고 조회)...")
    for _ in range(7):
        pyautogui.press('down')
        time.sleep(0.15)
    pyautogui.press('enter')
    time.sleep(LOAD_DELAY)

    # ── Tab×3 → ↓3 → Enter ───────────────────────────────────────────────────
    log("  Tab×3 → ↓3 → Enter...")
    for _ in range(3):
        pyautogui.press('tab')
        time.sleep(0.2)
    for _ in range(3):
        pyautogui.press('down')
        time.sleep(0.15)
    pyautogui.press('enter')
    time.sleep(LOAD_DELAY)

    # ── Tab×7 → Enter (조회 실행) ────────────────────────────────────────────
    log("  Tab×7 → Enter (조회)...")
    for _ in range(7):
        pyautogui.press('tab')
        time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(LOAD_DELAY)

    # ── Tab×1 → Enter → 재고현황 Excel 다운로드 ─────────────────────────────
    import win32gui
    file_name = f"재고현황_{datetime.now().strftime('%H%M%S')}"
    save_path = str(INVENTORY_DIR / f"{file_name}.xlsx")

    log("  Tab×1 → Enter (Excel 다운로드)...")
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(EXCEL_DELAY)

    # ── 저장 다이얼로그 처리 (kitting과 동일 방식) ────────────────────────────
    if _handle_save_dialog(save_path):
        log(f"  ✅ 저장 완료: {file_name}.xlsx")
        pyautogui.hotkey('ctrl', 'w')   # Excel 창 닫기
        time.sleep(0.8)
    else:
        # fallback: Downloads 폴더에서 이동
        import shutil as _shutil
        latest = _find_latest_download()
        if latest:
            _shutil.move(str(latest), save_path)
            log(f"  ✅ 이동 완료: {file_name}.xlsx")
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
        else:
            log("  ⚠️  파일 저장 실패")

    # ── Excel 창 닫기 ────────────────────────────────────────────────────────
    log("  Ctrl+W → Excel 창 닫기...")
    closed = False
    try:
        excel_hwnd = None
        def _find_excel(h, _):
            nonlocal excel_hwnd
            try:
                title = win32gui.GetWindowText(h)
                if any(k in title for k in ['.xlsx', '.xls', 'Excel', '엑셀', file_name]):
                    excel_hwnd = h
            except Exception:
                pass
        win32gui.EnumWindows(_find_excel, None)
        if excel_hwnd:
            win32gui.SetForegroundWindow(excel_hwnd)
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
            closed = True
            log("  ✅ Excel 창(hwnd) 포커스 후 Ctrl+W 완료")
    except Exception as e:
        log(f"  hwnd 방식 실패({e}) → 직접 Ctrl+W")

    if not closed:
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(0.5)

    log("  ✅ 재고현황 저장 및 닫기 완료")
    # 실제 저장된 최신 파일 반환
    inv_latest = sorted(INVENTORY_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
    if inv_latest:
        return str(inv_latest[0])
    return str(INVENTORY_DIR)


# ──────────────────────────────────────────────────────────────────────────────
# Step 6 : 자재부족현황 Playwright 업로드
# ──────────────────────────────────────────────────────────────────────────────
def automate_upload(downloaded_files):
    from playwright.sync_api import sync_playwright

    log("▶ Step 6: 자재부족현황.html 자동 업로드...")

    if not downloaded_files:
        downloaded_files = [str(f) for f in DOWNLOAD_DIR.glob("*.xlsx")]
        downloaded_files += [str(f) for f in DOWNLOAD_DIR.glob("*.xls")]

    if not downloaded_files:
        log("  ❌ 업로드할 파일 없음.")
        return

    with sync_playwright() as p:
        page = None
        is_cdp = False
        browser = None

        try:
            # ── 1순위: 이미 열린 브라우저에 CDP 연결 ─────────────────────────
            try:
                browser = p.chromium.connect_over_cdp("http://localhost:9222")
                is_cdp = True
                log("  기존 브라우저 CDP 연결 성공")

                for ctx in browser.contexts:
                    for pg in ctx.pages:
                        if "material-shortage" in pg.url:
                            page = pg
                            page.bring_to_front()
                            log("  기존 앱 탭 사용")
                            break
                    if page:
                        break

                if not page:
                    ctx = browser.contexts[0] if browser.contexts else browser.new_context()
                    page = ctx.new_page()
                    page.goto("https://material-shortage.vercel.app/", timeout=30000)
                    page.wait_for_load_state("networkidle", timeout=15000)

            except Exception:
                log("  CDP 연결 실패 → 새 브라우저 실행")
                browser = p.chromium.launch(
                    headless=False,
                    args=["--start-maximized", "--disable-web-security"]
                )
                page = browser.new_context(no_viewport=True).new_page()
                try:
                    page.goto("https://material-shortage.vercel.app/", timeout=30000)
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception as e2:
                    log(f"  페이지 로드 실패: {e2}")
                    input("  페이지 수동 로드 후 Enter → ")

            # ── 로그인 (이미 로그인 상태면 스킵) ──────────────────────────────
            # .btn-logout 버튼이 2개(설정/로그아웃)이므로 .first 사용
            if page.locator(".btn-logout").first.is_visible():
                log("  이미 로그인 상태")
            else:
                log("  로그인 중...")
                try:
                    page.locator("#login-email").fill(WEB_EMAIL)
                    page.locator("#login-pw").fill(WEB_PW)
                    page.locator("#auth-login .btn-auth").click()
                    page.wait_for_selector(".btn-logout", timeout=10000)
                    log("  로그인 완료")
                except Exception as e:
                    log(f"  로그인 실패: {e}")
                    input("  수동 로그인 후 Enter → ")

            # ── 키팅 로컬 상태 초기화 ──────────────────────────────────────────
            log("  키팅 로컬 상태 초기화...")
            try:
                page.evaluate("""
                    () => {
                        window.state.kitFiles = [];
                        if (window.dbPut) window.dbPut('kit_list', []);
                        if (window.renderKitChips) window.renderKitChips();
                        if (window.checkReady) window.checkReady();
                        localStorage.removeItem('ms_uptime_kit');
                        localStorage.removeItem('ms_uploader_kit');
                        const el = document.getElementById('uptime-kit');
                        if (el) el.textContent = '';
                    }
                """)
            except Exception as e:
                log(f"  초기화 오류: {e}")

            # ── 키팅된 자재 업로드 (kitting 자재 + 전일키팅) ────────────────────
            PREV_KIT_DIR = _PREV_KIT_DIR_CFG if _PREV_KIT_DIR_CFG else (_BASE / "전일키팅")
            all_kit = sorted(DOWNLOAD_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime)
            all_kit += sorted(DOWNLOAD_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime)
            if PREV_KIT_DIR.exists():
                all_kit += sorted(PREV_KIT_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime)
                all_kit += sorted(PREV_KIT_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime)
            valid = [str(f) for f in all_kit if f.exists()]
            log(f"  kitting 자재 + 전일키팅 전체 {len(valid)}개 파일 업로드 중...")
            if valid:
                page.locator("#file-kit").set_input_files(valid)

                try:
                    page.wait_for_selector('.kit-chip', timeout=15000)
                    log("  파일 처리 완료 — Supabase 저장 중...")
                except Exception:
                    time.sleep(5)

                try:
                    result = page.evaluate("""
                        async () => {
                            const files = (window.state && window.state.kitFiles) || [];
                            if (!files.length) return { ok: false, reason: 'state.kitFiles 비어있음' };
                            const results = [];
                            for (const f of files) {
                                const r = await window.uploadFileToStorage('kit/' + f.name, f.rawData);
                                results.push({ name: f.name, ok: r?.ok, msg: r?.msg || '' });
                            }
                            const ts = Date.now();
                            const uname = (window.currentUser && window.currentUser.displayName) || '';
                            await window.syncUploadLog('kit', ts, uname, files.map(f => f.name));
                            localStorage.setItem('ms_uptime_kit', String(ts));
                            if (uname) localStorage.setItem('ms_uploader_kit', uname);
                            return { ok: true, fileCount: files.length, names: files.map(f => f.name), results };
                        }
                    """, None)
                    if result and result.get('ok'):
                        log(f"  키팅 Supabase 저장 완료: {result.get('names')}")
                    else:
                        log(f"  키팅 Supabase 저장 결과: {result}")
                except Exception as e:
                    log(f"  키팅 Supabase 저장 오류: {e} — 10초 추가 대기")
                    time.sleep(10)

            # ── 재고현황 업로드 ────────────────────────────────────────────────
            inv_files = sorted(INVENTORY_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
            inv_files += sorted(INVENTORY_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime, reverse=True)
            if inv_files:
                latest_inv = str(inv_files[0])
                log(f"  재고현황 업로드: {inv_files[0].name}")
                page.locator("#file-inv").set_input_files(latest_inv)
                time.sleep(3)
                log("  재고현황 업로드 완료")
            else:
                log("  재고현황 폴더에 파일 없음")

            # ── 완료 팝업 ──────────────────────────────────────────────────────
            try:
                page.evaluate("""
                    const el = document.createElement('div');
                    el.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.55);z-index:99999;display:flex;align-items:center;justify-content:center';
                    el.innerHTML = '<div style="background:white;border-radius:16px;padding:36px 52px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.35)">' +
                        '<div style="font-size:52px;margin-bottom:12px">완료</div>' +
                        '<div style="font-size:22px;font-weight:700;color:#1e3a5f;margin-bottom:8px">실행완료 되었습니다.</div>' +
                        '<div style="font-size:13px;color:#666;margin-bottom:20px">키팅 파일이 자재부족현황에 업로드되었습니다.</div>' +
                        '<button onclick="this.closest(\'div[style]\').remove()" style="padding:10px 36px;background:#1e3a5f;color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">확인</button>' +
                        '</div>';
                    document.body.appendChild(el);
                """)
            except Exception:
                pass

            log("전체 자동화 완료!")
            try:
                import ctypes
                ctypes.windll.user32.MessageBoxW(0, "키팅 자동화가 완료되었습니다.", "✅ 자동화 완료", 0x40)
            except Exception:
                pass

        except Exception as e:
            log(f"  업로드 오류: {e}")
            import traceback as _tb; _tb.print_exc()
            input("  오류 확인 후 Enter → ")

        # CDP 연결은 브라우저를 닫지 않음 — 새 브라우저는 60초 후 종료
        if not is_cdp and browser:
            try:
                page.wait_for_timeout(60000)
            except Exception:
                pass
            try:
                browser.close()
            except Exception:
                pass


# ──────────────────────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────────────────────
def main():
    _init_log()
    log("=" * 60)
    log(f"sMES 키팅 자동화 시작  ({TODAY})")
    log("=" * 60)

    if not is_admin():
        log("관리자 권한 필요. 재실행 중...")
        time.sleep(1)
        elevate()
        return

    try:
        # Step 1: 실행
        launch_smes()

        # sMES 창 연결
        log("▶ Step 2: sMES 창 연결...")
        app, win = _get_smes_window()
        if not win:
            log("  sMES 창을 찾을 수 없습니다. 수동으로 sMES를 열어주세요.")
            input("  sMES 열린 후 Enter → ")
            app, win = _get_smes_window()

        # 디버그: 창 정보 출력
        _dump_window_info(win)

        # Step 2: 로그인
        log("▶ Step 3: 로그인...")
        login_smes(win)

        # Step 3~5: 키팅 자재 메뉴 이동 + 다운로드
        log("▶ Step 4~5: 키팅 자재 메뉴 이동 및 다운로드...")
        downloaded = navigate_and_download(app)

        # Step 5b: 재고현황 다운로드
        navigate_and_download_inventory(app)

        # Step 6: 자재부족현황 웹앱 업로드
        automate_upload(downloaded)

        # Step 7: sMES 창 닫기
        log("▶ Step 7: sMES 창 닫기...")
        try:
            import psutil, win32gui, win32con
            pid = _get_smes_pid()

            # 1차: pywinauto close()
            _, main_win = _get_smes_window()
            if main_win:
                try:
                    main_win.close()
                    time.sleep(1.5)
                except Exception:
                    pass

            # 2차: WM_CLOSE 메시지 전송 (닫기 확인창 자동 처리)
            if pid and psutil.pid_exists(pid):
                def _close_hwnd(h, _):
                    try:
                        if win32gui.IsWindowVisible(h):
                            wp = win32gui.GetWindowThreadProcessId(h)[1]
                            if wp == pid:
                                win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                    except Exception:
                        pass
                win32gui.EnumWindows(_close_hwnd, None)
                time.sleep(1.5)

                # 닫기 확인 팝업(예/아니오) 자동 Enter
                try:
                    import pyautogui
                    pyautogui.press('enter')
                    time.sleep(0.5)
                except Exception:
                    pass

            # 3차: 프로세스 강제 종료
            if pid and psutil.pid_exists(pid):
                psutil.Process(pid).terminate()
                time.sleep(1)
                log("  ✅ sMES 프로세스 종료 완료")
            else:
                log("  ✅ sMES 창 닫기 완료")
        except Exception as e:
            log(f"  ⚠️  sMES 닫기 실패: {e}")

    except Exception as e:
        log(f"❌ 오류: {e}")
        import traceback; traceback.print_exc()
        input("오류 확인 후 Enter → ")


if __name__ == "__main__":
    main()
