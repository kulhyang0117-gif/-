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
DOWNLOAD_DIR   = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\kitting 자재")
INVENTORY_DIR  = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\재고현황")
HTML_FILE      = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\자재부족현황.html")

SMES_ID      = "SSAT045"
SMES_PW      = "rlatndus1!"

WEB_EMAIL    = "kulhyang0117@gmail.com"
WEB_PW       = "jxy0830!"

TODAY        = datetime.now().strftime("%Y-%m-%d")
TODAY_KR     = datetime.now().strftime("%Y%m%d")

STEP_DELAY   = 1.0
LOAD_DELAY   = 4.0
EXCEL_DELAY  = 5.0

# ──────────────────────────────────────────────────────────────────────────────
# 유틸
# ──────────────────────────────────────────────────────────────────────────────
def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)

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


def login_smes(_unused_win=None):
    """
    ① PID 로 sMES 연결
    ② Edit 컨트롤 2개 있는 창 = 로그인 폼
    ③ Y좌표 정렬 → 위=ID, 아래=PW
    ④ set_edit_text() 로 입력 (키보드 우회 → 특수문자 안전)
    ⑤ Login 버튼 클릭
    """
    log("  로그인 폼 탐색 중 (PID 방식)...")
    app, dlg = _find_login_dialog(timeout=20)

    dlg.set_focus()
    time.sleep(0.3)

    # ── Edit 컨트롤 수집 & Y좌표 정렬 ──────────────────────────────────────
    edits = dlg.children(class_name_re="WindowsForms10.EDIT.*")
    edits = sorted(edits, key=lambda c: c.rectangle().top)
    log(f"  Edit 컨트롤 {len(edits)}개  (위→아래 순)")

    if len(edits) < 2:
        raise RuntimeError(f"Edit 컨트롤 {len(edits)}개 — 2개 필요")

    id_field = edits[0]   # 첫 번째(위) = ID 입력란
    pw_field = edits[1]   # 두 번째(아래) = PW 입력란

    # ── ID 필드 : 클릭 → 전체 삭제 → 타이핑 ───────────────────────────────
    id_field.set_focus()
    id_field.click_input()
    time.sleep(0.3)
    id_field.type_keys("^a{BACKSPACE}", with_spaces=True)
    time.sleep(0.2)
    id_field.type_keys(SMES_ID, with_spaces=True)
    log(f"  ① ID 입력 완료: {SMES_ID}")
    time.sleep(0.3)

    # ── PW 필드 : 클릭 → 전체 삭제 → 클립보드 붙여넣기 → Enter ────────────
    pw_field.set_focus()
    pw_field.click_input()
    time.sleep(0.3)
    pw_field.type_keys("^a{BACKSPACE}", with_spaces=True)
    time.sleep(0.2)
    _paste_text(SMES_PW)
    log(f"  ② PW 입력 완료")
    time.sleep(0.5)   # 붙여넣기 후 시스템이 인식할 시간 확보
    pw_field.type_keys("{ENTER}")

    # ── 로그인 성공 여부 검증 ────────────────────────────────────────────────
    # 로그인 폼(Edit 2개 있는 창)이 사라지면 성공, 남아있으면 실패
    log("  로그인 결과 확인 중...")
    deadline = time.time() + 10
    while time.time() < deadline:
        time.sleep(0.5)
        try:
            # 창이 사라졌는지 확인
            if not dlg.exists(timeout=0.5):
                log("  ✅ 로그인 성공 (로그인 창 닫힘)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            # exists() 자체가 예외 → 창 없어진 것
            log("  ✅ 로그인 성공 (로그인 창 닫힘)")
            time.sleep(LOAD_DELAY)
            return

        # 창이 아직 있으면 에러 메시지 확인
        try:
            for child in dlg.children():
                txt = child.window_text().strip()
                if txt and txt not in (SMES_ID, '') and len(txt) < 100:
                    log(f"  ⚠️  화면 메시지: '{txt}'")
        except Exception:
            pass

    # 10초 후에도 창이 남아있으면 실패
    raise RuntimeError(
        "로그인 실패: 로그인 창이 닫히지 않았습니다.\n"
        "  → ID/PW 를 확인하거나 수동으로 로그인 후 Enter를 눌러주세요."
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
            menu_y = rect.top + 7
            menu_x = rect.left + 230   # System+기본정보관리+영업관리 이후 위치
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
    if not _try_click(main_win, ["조회", "검색", "Search"]):
        log("  ⚠️  수동으로 '조회'를 클릭해주세요.")
        input("  조회 완료 후 Enter → ")
    time.sleep(LOAD_DELAY)
    log("  ✅ 조회 완료")

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


def _click_excel_download(win, idx, item_name):
    """Excel 다운로드 버튼 클릭 및 저장"""
    import pyautogui
    from pywinauto import Desktop

    # Excel 버튼 클릭 — 미발견 시 input() 차단 없이 자동 건너뜀
    if not _try_click(win, ["Excel 다운로드", "Excel", "엑셀", "엑셀 다운로드", "Export", "EXCEL"]):
        log(f"    ⚠️  [{idx}] Excel 버튼 미발견 — 건너뜀")
        return None

    time.sleep(EXCEL_DELAY)

    # 저장 다이얼로그 처리 — idx 포함으로 파일명 중복 방지
    save_path = str(DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}_{idx:03d}.xlsx")
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

    log(f"    ⚠️  [{idx}] 파일 저장 실패 — 건너뜀")
    return None


def _download_by_keyboard(win):
    import pyautogui

    ROW_X      = 1000   # 품목명 컬럼 X (절대 좌표)
    ROW_Y      = 165    # 첫 번째(선택) 행 Y (절대 좌표)
    ROW_HEIGHT = 22     # 행 높이 픽셀 — 실제 MES 그리드 행 높이에 맞게 조정

    win.set_focus(); time.sleep(0.3)
    pyautogui.click(ROW_X, ROW_Y)
    log(f"  첫 번째 행 클릭: ({ROW_X}, {ROW_Y})")
    time.sleep(0.5)

    downloaded = []
    region = (ROW_X - 200, 140, 500, 300)
    same_count = 0   # 연속으로 화면 변화 없는 횟수 카운트

    for i in range(500):   # 200 → 500으로 확장 (전 품목 완료 보장)
        log(f"  [{i+1}] Excel 다운로드 시도...")
        saved = _click_excel_download(win, i + 1, f"kitting_{i+1:03d}")
        if saved:
            downloaded.append(saved)

        # 포커스를 창으로만 복귀 — 고정 Y 클릭 금지 (행 선택 초기화 방지)
        win.set_focus()
        time.sleep(0.4)

        before = pyautogui.screenshot(region=region)

        # 현재 선택된 행에만 클릭 (색상 탐지 성공 시) — 실패 시 클릭 없이 Down만
        sel_y = _find_selected_row_y(ROW_X, grid_top=150, grid_bottom=560, row_height=ROW_HEIGHT)
        if sel_y is not None:
            log(f"    현재 선택 행 클릭: ({ROW_X}, {sel_y}) → ↓")
            pyautogui.click(ROW_X, sel_y)
            time.sleep(0.2)

        pyautogui.press('down')   # 다음 품목으로 이동
        time.sleep(0.5)
        after = pyautogui.screenshot(region=region)

        # 스크린샷이 동일 → 더 이상 내려갈 행 없음 확인 (3회 연속 동일 시 완료)
        if list(before.getdata()) == list(after.getdata()):
            same_count += 1
            log(f"    화면 변화 없음 ({same_count}/3)")
            if same_count >= 3:
                log(f"  ✅ 마지막 행 도달 — 전체 {len(downloaded)}개 다운로드 완료")
                break
        else:
            same_count = 0   # 화면 변화 있으면 카운트 초기화
    else:
        log(f"  ✅ 최대 반복 도달 — 전체 {len(downloaded)}개 다운로드 완료")

    return downloaded


def _handle_save_dialog(save_path):
    """Windows 저장 다이얼로그 자동 처리"""
    import pyautogui
    from pywinauto import Desktop

    desktop = Desktop(backend='uia')
    for _ in range(20):
        try:
            dlg = desktop.window(title_re=".*(저장|Save As|다른 이름).*")
            if dlg.exists(timeout=0.5):
                dlg.set_focus()
                time.sleep(0.3)
                try:
                    fn_edit = dlg.child_window(control_type="Edit", found_index=0)
                    fn_edit.set_edit_text(save_path)
                except Exception:
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.1)
                    pyautogui.write(save_path, interval=0.02)
                time.sleep(0.3)
                pyautogui.press('enter')
                time.sleep(1.5)
                # 덮어쓰기 확인
                try:
                    conf = desktop.window(title_re=".*(덮어|overwrite|Confirm).*")
                    if conf.exists(timeout=1):
                        pyautogui.press('enter')
                except Exception:
                    pass
                return True
        except Exception:
            pass
        time.sleep(0.5)
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
            menu_y = rect.top + 7
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

    # ── Tab×1 → Enter → Excel 다운로드 트리거 ───────────────────────────────
    log("  Tab×1 → Enter → Excel 다운로드...")
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(EXCEL_DELAY)

    # ── Save As 다이얼로그: 재고현황+현재시간 으로 저장 ───────────────────────
    file_name = f"재고현황{datetime.now().strftime('%H%M%S')}"
    save_path = str(INVENTORY_DIR / f"{file_name}.xlsx")
    log(f"  Save As 다이얼로그 처리: {file_name}.xlsx")

    if _handle_save_dialog(save_path):
        log(f"  ✅ 저장 완료: {file_name}.xlsx")
    else:
        # 자동 저장된 경우 Downloads에서 이동
        latest = _find_latest_download()
        if latest:
            import shutil
            dest = INVENTORY_DIR / f"{file_name}.xlsx"
            shutil.move(str(latest), str(dest))
            log(f"  ✅ 이동 완료: {dest.name}")
        else:
            log("  ⚠️  재고현황 파일 저장 실패")

    # ── Ctrl+W → Excel 현재 창 닫기 ─────────────────────────────────────────
    log("  Ctrl+W → Excel 창 닫기...")
    pyautogui.hotkey('ctrl', 'w')
    time.sleep(0.5)

    log("  ✅ 재고현황 저장 및 닫기 완료")
    return save_path


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

        # ── 1순위: 이미 열린 브라우저에 CDP 연결 ────────────────────────────
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            is_cdp = True
            log("  ✅ 기존 브라우저 CDP 연결 성공")

            # 이미 열린 앱 탭 찾기
            for ctx in browser.contexts:
                for pg in ctx.pages:
                    if "material-shortage" in pg.url:
                        page = pg
                        page.bring_to_front()
                        log("  ✅ 기존 앱 탭 사용")
                        break
                if page:
                    break

            # 앱 탭 없으면 새 탭 열기
            if not page:
                ctx = browser.contexts[0] if browser.contexts else browser.new_context()
                page = ctx.new_page()
                page.goto("https://material-shortage.vercel.app/", timeout=30000)
                page.wait_for_load_state("networkidle", timeout=15000)

        except Exception as e:
            log(f"  CDP 연결 실패({e}) → 새 브라우저 실행")
            browser = p.chromium.launch(
                headless=False,
                args=["--start-maximized", "--disable-web-security"]
            )
            page = browser.new_context(no_viewport=True).new_page()
            page.goto("https://material-shortage.vercel.app/", timeout=30000)
            page.wait_for_load_state("networkidle", timeout=15000)

        # 로그인 (이미 로그인 상태면 스킵)
        if page.locator(".btn-logout").is_visible():
            log("  ✅ 이미 로그인 상태")
        else:
            log("  로그인 중...")
            try:
                page.locator("#login-email").fill(WEB_EMAIL)
                page.locator("#login-pw").fill(WEB_PW)
                page.locator("#auth-login .btn-auth").click()
                page.wait_for_selector(".btn-logout", timeout=10000)
                log("  ✅ 로그인 완료")
            except Exception as e:
                log(f"  ⚠️  로그인 실패: {e}")
                input("  수동 로그인 후 Enter → ")

        # 키팅 초기화
        log("  키팅 초기화...")
        try:
            page.locator(".btn-kit-clear").click()
            time.sleep(1)
        except Exception:
            pass

        # 재고현황 최신 파일 업로드
        inv_files = sorted(INVENTORY_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
        inv_files += sorted(INVENTORY_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime, reverse=True)
        if inv_files:
            latest_inv = str(inv_files[0])
            log(f"  재고현황 업로드: {inv_files[0].name}")
            page.locator("#file-inv").set_input_files(latest_inv)
            time.sleep(3)
            log("  ✅ 재고현황 업로드 완료")
        else:
            log("  ⚠️  재고현황 폴더에 파일 없음")

        # 파일 업로드
        valid = [f for f in downloaded_files if Path(f).exists()]
        log(f"  {len(valid)}개 파일 업로드 중...")
        if valid:
            page.locator("#file-kit").set_input_files(valid)

            # 칩 렌더링 대기 (파일이 state에 로드됐을 때)
            try:
                page.wait_for_selector('.kit-chip', timeout=15000)
                log("  파일 처리 완료 — Supabase 저장 중...")
            except Exception:
                time.sleep(5)

            # Storage 업로드 + upload_logs 동기화를 await로 명시적 완료
            try:
                page.evaluate("""
                    async () => {
                        const files = (window.state && window.state.kitFiles) || [];
                        if (!files.length) return;
                        const uploads = files.map(f =>
                            window.uploadFileToStorage('kit/' + f.name, f.rawData)
                        );
                        await Promise.all(uploads);
                        const ts = Date.now();
                        const uname = (window.currentUser && window.currentUser.displayName) || '';
                        await window.syncUploadLog('kit', ts, uname, files.map(f => f.name));
                    }
                """)
                log("  ✅ Supabase 저장 완료")
            except Exception as e:
                log(f"  ⚠️  Supabase 저장 오류: {e} — 10초 추가 대기")
                time.sleep(10)

        # 완료 팝업
        page.evaluate("""
            const el = document.createElement('div');
            el.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.55);z-index:99999;display:flex;align-items:center;justify-content:center';
            el.innerHTML = '<div style="background:white;border-radius:16px;padding:36px 52px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.35)">' +
                '<div style="font-size:52px;margin-bottom:12px">✅</div>' +
                '<div style="font-size:22px;font-weight:700;color:#1e3a5f;margin-bottom:8px">실행완료 되었습니다.</div>' +
                '<div style="font-size:13px;color:#666;margin-bottom:20px">키팅 파일이 자재부족현황에 업로드되었습니다.</div>' +
                '<button onclick="this.closest(\'div[style]\').remove()" style="padding:10px 36px;background:#1e3a5f;color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">확인</button>' +
                '</div>';
            document.body.appendChild(el);
        """)

        log("✅ 전체 자동화 완료!")

        # CDP 연결은 브라우저를 닫지 않음 — 새 브라우저는 60초 후 종료
        if not is_cdp:
            try:
                page.wait_for_timeout(60000)
            except Exception:
                pass
            browser.close()


# ──────────────────────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────────────────────
def main():
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
            _, main_win = _get_smes_window()
            if main_win:
                main_win.close()
                log("  ✅ sMES 창 닫기 완료")
            else:
                log("  ⚠️  sMES 창을 찾을 수 없음")
        except Exception as e:
            log(f"  ⚠️  sMES 닫기 실패: {e}")

    except Exception as e:
        log(f"❌ 오류: {e}")
        import traceback; traceback.print_exc()
        input("오류 확인 후 Enter → ")


if __name__ == "__main__":
    main()
