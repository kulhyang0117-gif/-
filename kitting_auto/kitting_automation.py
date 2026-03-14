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
DOWNLOAD_DIR = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\kitting 자재")
HTML_FILE    = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\자재부족현황.html")

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

    # ── 생산관리 클릭 ───────────────────────────────────────────────────────
    log("  생산관리 클릭...")
    if not _try_click(main_win, ["생산관리"]):
        log("  ⚠️  수동으로 '생산관리'를 클릭해주세요.")
        input("  완료 후 Enter → ")
    time.sleep(STEP_DELAY)

    # ── 조립 자재 kitting 클릭 ─────────────────────────────────────────────
    log("  조립 자재 kitting 클릭...")
    if not _try_click(main_win, ["조립 자재 kitting", "kitting", "Kitting", "키팅", "자재 kitting"]):
        log("  ⚠️  수동으로 '조립 자재 kitting'을 클릭해주세요.")
        input("  완료 후 Enter → ")
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


def download_all_items(win):
    import pyautogui
    downloaded_files = []

    # 그리드 행 찾기 시도
    rows = []
    for ct in ["DataItem", "ListItem", "TreeItem", "Custom"]:
        try:
            candidates = win.children(control_type=ct)
            if len(candidates) > 0:
                rows = candidates
                log(f"  그리드 행 {len(rows)}개 발견 (type={ct})")
                break
        except Exception:
            pass

    if rows:
        for i, row in enumerate(rows):
            try:
                item_name = row.window_text().strip() or f"item_{i+1:03d}"
                log(f"  [{i+1}/{len(rows)}] {item_name}")
                row.click_input()
                time.sleep(STEP_DELAY)

                saved = _click_excel_download(win, i + 1, item_name)
                if saved:
                    downloaded_files.append(saved)
            except Exception as e:
                log(f"    ⚠️  [{i+1}] 실패: {e}")
    else:
        # 행을 못 찾은 경우: 키보드 Down 방식
        log("  그리드 행 감지 실패 → 키보드 Down 방식으로 전환")
        downloaded_files = _download_by_keyboard(win)

    log(f"  ✅ {len(downloaded_files)}개 파일 다운로드 완료")
    return downloaded_files


def _click_excel_download(win, idx, item_name):
    """Excel 다운로드 버튼 클릭 및 저장"""
    import pyautogui
    from pywinauto import Desktop

    # Excel 버튼 클릭
    if not _try_click(win, ["Excel 다운로드", "Excel", "엑셀", "엑셀 다운로드", "Export", "EXCEL"]):
        log(f"    ⚠️  Excel 버튼 미발견. 수동 저장 후 Enter → ")
        input()
        latest = _find_latest_download()
        if latest:
            dest = DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}.xlsx"
            import shutil; shutil.move(str(latest), str(dest))
            return str(dest)
        return None

    time.sleep(EXCEL_DELAY)

    # 저장 다이얼로그 처리
    save_path = str(DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}.xlsx")
    if _handle_save_dialog(save_path):
        log(f"    ✅ 저장: {Path(save_path).name}")
        return save_path

    # 다이얼로그 없이 자동 저장된 경우 — Downloads 폴더에서 이동
    latest = _find_latest_download()
    if latest:
        import shutil
        dest = DOWNLOAD_DIR / f"{_safe(item_name)}_{TODAY_KR}.xlsx"
        shutil.move(str(latest), str(dest))
        log(f"    ✅ 이동: {dest.name}")
        return str(dest)

    return None


def _download_by_keyboard(win):
    import pyautogui
    downloaded = []
    win.set_focus()
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'Home')
    time.sleep(0.3)

    for i in range(200):
        log(f"  [{i+1}] Excel 다운로드 시도...")
        saved = _click_excel_download(win, i + 1, f"kitting_{i+1:03d}")
        if saved:
            downloaded.append(saved)
        else:
            break
        pyautogui.press('down')
        time.sleep(0.5)

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
# Step 6 : 자재부족현황.html Playwright 업로드
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
        browser = p.chromium.launch(
            headless=False,
            args=["--start-maximized", "--disable-web-security"]
        )
        page = browser.new_context(no_viewport=True).new_page()
        page.goto(HTML_FILE.as_uri(), timeout=30000)
        time.sleep(2)

        # 로그인
        log("  로그인 중...")
        try:
            page.locator("#login-email").fill(WEB_EMAIL)
            page.locator("#login-pw").fill(WEB_PW)
            page.locator("#btn-login").click()
            time.sleep(2)
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

        # 파일 업로드
        valid = [f for f in downloaded_files if Path(f).exists()]
        log(f"  {len(valid)}개 파일 업로드 중...")
        if valid:
            page.locator("#file-kit").set_input_files(valid)
            time.sleep(max(3, len(valid) * 0.5))

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
        try:
            page.wait_for_timeout(60000)
        except Exception:
            pass


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

        # Step 3~5: 메뉴 이동 + 다운로드
        log("▶ Step 4~5: 메뉴 이동 및 다운로드...")
        downloaded = navigate_and_download(app)

        # Step 6: 업로드
        automate_upload(downloaded)

    except Exception as e:
        log(f"❌ 오류: {e}")
        import traceback; traceback.print_exc()
        input("오류 확인 후 Enter → ")


if __name__ == "__main__":
    main()
