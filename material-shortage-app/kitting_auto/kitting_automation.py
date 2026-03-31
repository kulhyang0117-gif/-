#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sMES 키팅 자재 자동 다운로드 & 자재부족현황 자동 업로드
========================================================
실행 전 준비:
  pip install pywinauto pyautogui pillow playwright psutil keyboard
  playwright install chromium
"""

import os, sys, time, ctypes, subprocess, argparse
from datetime import datetime
from pathlib import Path

# ──────────────────────────────────────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────────────────────────────────────
SMES_EXE     = Path(r"C:\Program Files (x86)\I2R\sMES\sMES.exe")
DOWNLOAD_DIR   = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\kitting 자재")
INVENTORY_DIR  = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\재고현황")
HTML_FILE      = Path(r"C:\Users\조립\Desktop\claude\Material Shortage Status vs. Production Plan\자재부족현황.html")
STOP_FLAG      = Path(__file__).parent / "stop_flag.txt"

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

def check_stop():
    """긴급정지 플래그 파일이 있으면 삭제 후 StopIteration 발생"""
    if STOP_FLAG.exists():
        STOP_FLAG.unlink(missing_ok=True)
        log("🛑 긴급정지 요청 감지 — 자동화를 중단합니다.")
        raise StopIteration("긴급정지")

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
# Step 2 : sMES 로그인  (PDF 3단계 우선순위 구현)
# ──────────────────────────────────────────────────────────────────────────────

def _type_password(pwd: str):
    """
    비밀번호 직접 키 입력.
    일반 문자는 typewrite, Shift 조합 특수문자는 hotkey로 각각 타이핑.
    """
    import pyautogui

    # Shift+키 조합이 필요한 특수문자 매핑
    SHIFT_MAP = {
        '!': '1', '@': '2', '#': '3', '$': '4', '%': '5',
        '^': '6', '&': '7', '*': '8', '(': '9', ')': '0',
        '_': '-', '+': '=', '{': '[', '}': ']', '|': '\\',
        ':': ';', '"': "'", '<': ',', '>': '.', '?': '/',
        '~': '`',
    }

    for ch in pwd:
        if ch in SHIFT_MAP:
            pyautogui.hotkey('shift', SHIFT_MAP[ch])
        else:
            pyautogui.typewrite(ch, interval=0.0)
        time.sleep(0.05)


def _check_caps_lock():
    """Caps Lock 켜져 있으면 자동 해제"""
    import pyautogui
    if ctypes.windll.user32.GetKeyState(0x14) & 1:
        log("  ⚠️  Caps Lock 켜져 있음 → 자동 해제")
        pyautogui.press("capslock")


def _verify_foreground(hwnd, fallback_hwnd=None):
    """
    클릭 전 MES 창이 포그라운드인지 검증.
    실패 시 자동 재활성화 1회 재시도 → 그래도 실패 시 False.
    """
    mes_hwnds = set()
    if hwnd:
        mes_hwnds.add(hwnd)
    if fallback_hwnd:
        mes_hwnds.add(fallback_hwnd)

    if not mes_hwnds:
        log("  MES 창 HWND 확인 불가 — 강제 진행")
        return True  # HWND 모를 땐 일단 허용

    fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
    if fg_hwnd not in mes_hwnds:
        log(f"  ⚠️  포그라운드 아님 (fg={fg_hwnd}) → 자동 재활성화 시도")
        _force_foreground(hwnd)
        time.sleep(0.4)
        fg_hwnd = ctypes.windll.user32.GetForegroundWindow()
        if fg_hwnd not in mes_hwnds:
            log(f"  ⚠️  재활성화 후에도 불일치 (fg={fg_hwnd}) — 그대로 진행")
            # 차단하지 않고 진행 (입력 시도)
    return True


def _force_foreground(hwnd):
    """UAC 환경에서도 창 강제 활성화 (AttachThreadInput 방식, ctypes 전용)"""
    try:
        fg_tid = ctypes.windll.user32.GetWindowThreadProcessId(
            ctypes.windll.user32.GetForegroundWindow(), None)
        my_tid = ctypes.windll.kernel32.GetCurrentThreadId()
        if fg_tid != my_tid:
            ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, True)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        if fg_tid != my_tid:
            ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, False)
    except Exception as e:
        log(f"  _force_foreground 실패: {e}")


def _check_and_dismiss_error_popup():
    """
    MES 에러 팝업 감지 및 자동 닫기.
    'PASSWORD를 입력하십시오' 등의 에러 다이얼로그가 뜨면 Enter로 닫고 메시지를 반환.
    에러 없으면 None 반환.
    """
    import pyautogui
    from pywinauto import Application

    ERROR_KEYWORDS = [
        "PASSWORD", "비밀번호", "오류", "실패", "틀렸", "error", "invalid",
        "wrong", "입력하십시오", "확인하십시오", "불일치"
    ]

    pid = _get_smes_pid()
    if not pid:
        return None

    try:
        app = Application(backend='win32').connect(process=pid, timeout=3)
        for w in app.windows():
            title = w.window_text().strip()
            # 작은 창 = 팝업 (메인 창 제외)
            rect = w.rectangle()
            is_small = rect.width() < 600 and rect.height() < 300

            found_msg = None
            # 창 제목에 에러 키워드
            if any(kw in title for kw in ERROR_KEYWORDS):
                found_msg = title
            # 작은 팝업 창의 내부 텍스트 확인
            if is_small and not found_msg:
                try:
                    for child in w.children():
                        txt = child.window_text().strip()
                        if txt and any(kw in txt for kw in ERROR_KEYWORDS):
                            found_msg = txt
                            break
                except Exception:
                    pass

            if found_msg:
                log(f"  ❌ 에러 팝업 감지: '{found_msg}' → 자동 닫기")
                try:
                    # 확인/OK 버튼 클릭
                    for btn in w.children(class_name_re=".*BUTTON.*"):
                        btn_txt = btn.window_text().strip().lower()
                        if btn_txt in ("확인", "ok", "닫기", "close", "예", "yes"):
                            btn.click_input()
                            time.sleep(0.3)
                            return found_msg
                    # 버튼 못 찾으면 Enter
                    pyautogui.press("enter")
                    time.sleep(0.3)
                except Exception:
                    pyautogui.press("enter")
                    time.sleep(0.3)
                return found_msg
    except Exception as e:
        log(f"  팝업 감지 중 오류: {e}")

    return None


def _find_login_dialog(timeout=20):
    """
    로그인 창 감지 — 4단계 탐지 (Desktop 전체 스캔 포함):
      1) Desktop 전체에서 제목 키워드 매칭 (가장 확실)
      2) Desktop 전체에서 Edit 2개 이상인 작은 창 (로그인 폼 특성)
      3) PID로 연결된 app.windows() 스캔
      4) PID 창 중 가장 큰 창 (메인 창에 패널 내장 구조)
    """
    from pywinauto import Application, Desktop

    # 창 제목 키워드
    TITLE_KEYWORDS = ["Shinsung", "IMMES", "sMES", "MES", "LOGIN", "로그인", "Login"]

    pid = _get_smes_pid()
    if not pid:
        raise RuntimeError("sMES 프로세스를 찾을 수 없습니다.")

    log(f"  sMES PID={pid} — 로그인 창 탐색 시작...")
    app = Application(backend='win32').connect(process=pid, timeout=10)

    deadline = time.time() + timeout
    while time.time() < deadline:

        # ── 1) Desktop 전체 창 스캔 (가장 확실한 방법) ──────────────────────
        try:
            desktop = Desktop(backend='win32')
            all_wins = desktop.windows()
            log(f"  Desktop 전체 창 {len(all_wins)}개 스캔 중...")
            for w in all_wins:
                try:
                    title = w.window_text().strip()
                    if not title:
                        continue
                    # 제목 키워드 매칭
                    if any(kw.lower() in title.lower() for kw in TITLE_KEYWORDS):
                        log(f"  ✅ 로그인 창 감지 (Desktop 제목): '{title}'")
                        # 이 창이 sMES 프로세스 소속인지 확인 (아니어도 허용)
                        try:
                            win_pid = ctypes.c_ulong()
                            ctypes.windll.user32.GetWindowThreadProcessId(
                                w.handle, ctypes.byref(win_pid))
                            if win_pid.value == pid:
                                log(f"    → sMES 프로세스 소속 확인")
                        except Exception:
                            pass
                        return app, w
                except Exception:
                    pass
        except Exception as e:
            log(f"  Desktop 스캔 오류: {e}")

        # ── 2) Desktop 전체에서 Edit 2개 이상인 창 ─────────────────────────
        try:
            desktop = Desktop(backend='win32')
            for w in desktop.windows():
                try:
                    # 창 크기 기준: 로그인 창은 보통 200~800px 범위
                    rect = w.rectangle()
                    if rect.width() < 100 or rect.height() < 100:
                        continue
                    edits = w.children(class_name_re="WindowsForms10.EDIT.*")
                    if 2 <= len(edits) <= 20:
                        title = w.window_text().strip()
                        log(f"  ✅ 로그인 창 감지 (Desktop Edit {len(edits)}개): '{title}'")
                        return app, w
                except Exception:
                    pass
        except Exception:
            pass

        # ── 3) PID 연결 창 스캔 ─────────────────────────────────────────────
        try:
            for w in app.windows():
                title = w.window_text().strip()
                if any(kw.lower() in title.lower() for kw in TITLE_KEYWORDS):
                    log(f"  ✅ 로그인 창 감지 (PID 창 제목): '{title}'")
                    return app, w
                try:
                    edits = w.children(class_name_re="WindowsForms10.EDIT.*")
                    if len(edits) >= 2:
                        log(f"  ✅ 로그인 창 감지 (PID Edit {len(edits)}개): '{title}'")
                        return app, w
                except Exception:
                    pass
        except Exception as e:
            log(f"  PID 스캔 오류: {e}")

        time.sleep(0.5)

    # ── 4) 최후 수단: sMES PID에서 가장 큰 창 ───────────────────────────────
    try:
        wins = app.windows()
        if wins:
            biggest = max(wins, key=lambda w: w.rectangle().width() * w.rectangle().height())
            log(f"  ⚠️  탐지 실패 → sMES 메인 창 사용: '{biggest.window_text()}'")
            return app, biggest
    except Exception:
        pass

    raise RuntimeError("로그인 폼을 찾을 수 없습니다.")


def _find_login_controls(dlg):
    """
    Login 버튼 + ID/PW Edit 컨트롤 반환.
    ① window_text 기반: SMES_ID 값 = ID필드, 'PASSWORD' = PW필드(플레이스홀더)
    ② 위치 기반 fallback: Login 버튼 바로 위 2개 컨트롤
    반환: (id_ctrl, pw_ctrl), login_btn
    """
    # Login 버튼
    login_btn = None
    for btn in dlg.children(class_name_re="WindowsForms10.BUTTON.*"):
        try:
            if btn.window_text().strip().lower() in ("login", "로그인"):
                login_btn = btn
                break
        except Exception:
            pass

    all_edits = list(dlg.children(class_name_re="WindowsForms10.EDIT.*"))
    log(f"  children() Edit 총 {len(all_edits)}개")

    id_ctrl = None
    pw_ctrl = None

    # ① 텍스트로 식별 (가장 신뢰도 높음)
    for e in all_edits:
        try:
            val = e.window_text().strip()
            if val == SMES_ID and id_ctrl is None:
                id_ctrl = e
                log(f"    → ID 필드 확인 (text='{val}')")
            elif val.upper() == 'PASSWORD' and pw_ctrl is None:
                pw_ctrl = e
                log(f"    → PW 필드 확인 (placeholder='PASSWORD')")
        except Exception:
            pass

    # ② 위치 기반 fallback: Login 버튼 바로 위 2개
    if (id_ctrl is None or pw_ctrl is None) and login_btn:
        btn_y = login_btn.rectangle().top
        above = sorted(
            [e for e in all_edits if e.rectangle().top < btn_y],
            key=lambda e: e.rectangle().top
        )
        if pw_ctrl is None and len(above) >= 1:
            pw_ctrl = above[-1]   # 버튼 바로 위 = PW
            log(f"    → PW 필드 위치 fallback (Y={pw_ctrl.rectangle().top})")
        if id_ctrl is None and len(above) >= 2:
            id_ctrl = above[-2]   # PW 위 = ID
            log(f"    → ID 필드 위치 fallback (Y={id_ctrl.rectangle().top})")

    return (id_ctrl, pw_ctrl), login_btn


def _login_by_autoid(dlg, hwnd):
    """1순위: children() 직접 스캔 — C1TextBox ID/PW + C1Button Login"""
    import pyautogui
    log("  [1순위] children() 직접 스캔 방식 시도...")

    try:
        (id_ctrl, pw_ctrl), login_btn = _find_login_controls(dlg)

        if id_ctrl is None or pw_ctrl is None:
            raise RuntimeError(f"컨트롤 식별 실패 (id={id_ctrl}, pw={pw_ctrl})")
        if not login_btn:
            raise RuntimeError("Login 버튼 없음")

        log(f"  ID ctrl: '{id_ctrl.window_text()[:20]}' / PW ctrl: 확인됨")

        # ① ID
        existing_id = id_ctrl.window_text().strip()
        if existing_id and existing_id.upper() != 'PASSWORD':
            log(f"  ① ID '{existing_id}' 이미 있음 → 스킵")
        else:
            _verify_foreground(hwnd)
            id_ctrl.click_input(); time.sleep(0.3)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("delete")
            time.sleep(0.1)
            pyautogui.typewrite(SMES_ID, interval=0.05)
            log(f"  ① ID 입력: {SMES_ID}")
            time.sleep(0.3)
            # 입력 확인
            entered = id_ctrl.window_text().strip()
            if not entered or entered.upper() == 'PASSWORD':
                raise RuntimeError(f"ID 입력 실패 (현재: '{entered}')")
            log(f"  ① ID 입력 확인: '{entered}'")

        _check_caps_lock()

        # ② PW
        _verify_foreground(hwnd)
        pw_ctrl.click_input(); time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("delete")
        time.sleep(0.1)
        _type_password(SMES_PW)
        log(f"  ② PW 입력 완료 ({len(SMES_PW)}자)")
        time.sleep(0.3)

        # ③ Login 버튼
        _verify_foreground(hwnd)
        login_btn.click_input()
        log("  ③ Login 버튼 클릭")
        return True

    except Exception as e:
        log(f"  [1순위] 실패: {e}")
        return False


def _login_by_coords(dlg, hwnd):
    """2순위: 창 기준 상대 좌표 방식 (비율 계산)"""
    import pyautogui, pyperclip
    log("  [2순위] 창 기준 상대 좌표 방식 시도...")

    try:
        import pygetwindow as gw

        # 창 객체 확보 (pygetwindow)
        title = dlg.window_text()
        wins = gw.getWindowsWithTitle(title) if title else []
        login_win = wins[0] if wins else None

        if not login_win:
            # pywinauto rectangle로 직접 계산
            rect = dlg.rectangle()
            left, top = rect.left, rect.top
            width = rect.width()
            height = rect.height()
        else:
            if login_win.isMinimized:
                login_win.restore()
            login_win.activate()
            time.sleep(0.5)
            left   = login_win.left
            top    = login_win.top
            width  = login_win.width
            height = login_win.height

        cx      = left + int(width * 0.509)
        id_pos  = (cx, top + int(height * 0.544))
        pwd_pos = (cx, top + int(height * 0.649))
        btn_pos = (left + int(width * 0.318), top + int(height * 0.765))

        log(f"  창 위치: left={left} top={top} w={width} h={height}")
        log(f"  ID 좌표: {id_pos} / PW 좌표: {pwd_pos} / 버튼 좌표: {btn_pos}")

        # ① ID 필드 — Edit 컨트롤 값 읽어 기존 값 있으면 스킵
        try:
            edits = sorted(
                dlg.descendants(class_name_re="WindowsForms10.EDIT.*"),
                key=lambda c: c.rectangle().top
            )
            existing_id = edits[0].window_text().strip() if edits else ""
        except Exception:
            existing_id = ""

        if existing_id:
            log(f"  ① ID 필드에 '{existing_id}' 있음 → 스킵, 패스워드로 이동")
            if not _verify_foreground(hwnd):
                return False
            pyautogui.click(*pwd_pos)
            time.sleep(0.3)
        else:
            if not _verify_foreground(hwnd):
                return False
            pyperclip.copy(SMES_ID)
            pyautogui.click(*id_pos)
            time.sleep(0.4)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.hotkey("ctrl", "v")
            log(f"  ① ID 입력: {SMES_ID}")

        # Caps Lock 확인
        _check_caps_lock()

        # ② PW 입력 — 클릭 전 포그라운드 재검증 (ID 스킵 시 이미 pwd_pos 클릭됨)
        if existing_id:
            pass  # 이미 pwd_pos로 이동됨
        else:
            if not _verify_foreground(hwnd):
                return False
            pyautogui.press("tab")
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a")
        pyautogui.press("delete")
        _type_password(SMES_PW)
        log("  ② PW 입력 완료")

        # ③ 로그인 버튼 — children()에서 먼저 탐색, 없으면 좌표
        _verify_foreground(hwnd)
        btn_clicked = False
        try:
            _, login_btn = _find_login_controls(dlg)
            if login_btn:
                login_btn.click_input()
                log("  ③ 로그인 버튼 클릭 (children)")
                btn_clicked = True
        except Exception:
            pass
        if not btn_clicked:
            pyautogui.click(*btn_pos)
            log(f"  ③ 로그인 버튼 클릭 (좌표 fallback) {btn_pos}")
        return True

    except Exception as e:
        log(f"  [2순위] 실패: {e}")
        return False


def _login_by_edit_scan(dlg, hwnd):
    """3순위: children() 직접 스캔 — 로그 덤프 확인 기반 (C1TextBox + C1Button)"""
    import pyautogui, pyperclip
    log("  [3순위] children() 직접 스캔 방식 시도...")

    try:
        # ── Edit 컨트롤 (C1TextBox) — 직접 자식만, Y좌표 정렬 ──────────────
        edits = []
        # 1차: children() 직접 자식 (로그 덤프 확인: 2개만 존재)
        try:
            edits = sorted(
                dlg.children(class_name_re="WindowsForms10.EDIT.*"),
                key=lambda c: c.rectangle().top
            )
            log(f"  children() Edit {len(edits)}개")
        except Exception:
            pass

        # 2차: 그래도 부족하면 descendants() fallback
        if len(edits) < 2:
            try:
                all_edits = sorted(
                    dlg.descendants(class_name_re="WindowsForms10.EDIT.*"),
                    key=lambda c: c.rectangle().top
                )
                # 로그인 패널 Edit만 필터: 빈 값이거나 ID/PW 후보인 것
                login_edits = [
                    e for e in all_edits
                    if e.window_text().strip() in ('', SMES_ID, SMES_PW)
                    or len(e.window_text().strip()) <= 20
                ][:4]  # 최대 4개
                edits = login_edits
                log(f"  descendants() Edit 후보 {len(edits)}개")
            except Exception:
                pass

        if len(edits) < 2:
            log(f"  Edit 컨트롤 2개 미만 → 실패")
            return False

        id_ctrl = edits[0]   # Y 작은 = 위쪽 = ID
        pw_ctrl = edits[1]   # Y 큰  = 아래쪽 = PW
        log(f"  ID: class='{id_ctrl.class_name()}' val='{id_ctrl.window_text()[:10]}' pos={id_ctrl.rectangle().top}")
        log(f"  PW: class='{pw_ctrl.class_name()}' val='****' pos={pw_ctrl.rectangle().top}")

        # ── ① ID 입력 ────────────────────────────────────────────────────
        existing_id = ""
        try:
            existing_id = id_ctrl.window_text().strip()
        except Exception:
            pass

        if existing_id:
            log(f"  ① ID '{existing_id}' 이미 있음 → 스킵")
        else:
            _verify_foreground(hwnd)
            try:
                id_ctrl.click_input()
                time.sleep(0.2)
                pyautogui.hotkey("ctrl", "a")
                pyperclip.copy(SMES_ID)
                pyautogui.hotkey("ctrl", "v")
                log(f"  ① ID 입력: {SMES_ID}")
            except Exception:
                try:
                    id_ctrl.set_edit_text(SMES_ID)
                    log(f"  ① ID set_edit_text: {SMES_ID}")
                except Exception as e2:
                    log(f"  ① ID 입력 실패: {e2}")
            time.sleep(0.3)

        _check_caps_lock()

        # ── ② PW 입력 ────────────────────────────────────────────────────
        _verify_foreground(hwnd)
        try:
            pw_ctrl.click_input()
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("delete")
            _type_password(SMES_PW)
            log("  ② PW 입력 완료")
        except Exception:
            try:
                pw_ctrl.set_edit_text(SMES_PW)
                log("  ② PW set_edit_text 완료")
            except Exception as e2:
                log(f"  ② PW 입력 실패: {e2}")
        time.sleep(0.3)

        # ── ③ 로그인 버튼 — title='Login' 직접 탐색 ─────────────────────
        _verify_foreground(hwnd)
        btn_found = False
        # children() 에서 BUTTON 중 'Login' 찾기 (덤프 확인됨)
        try:
            for btn in dlg.children(class_name_re="WindowsForms10.BUTTON.*"):
                try:
                    txt = btn.window_text().strip()
                    if txt.lower() in ("login", "로그인", "확인", "ok"):
                        btn.click_input()
                        log(f"  ③ 버튼 클릭: '{txt}'")
                        btn_found = True
                        break
                except Exception:
                    pass
        except Exception:
            pass

        if not btn_found:
            # descendants 에서도 탐색
            try:
                for btn_kw in ["Login", "로그인", "LOGIN"]:
                    b = dlg.child_window(title=btn_kw, class_name_re="WindowsForms10.BUTTON.*")
                    if b.exists(timeout=0.5):
                        b.click_input()
                        log(f"  ③ 버튼 클릭 (child_window): '{btn_kw}'")
                        btn_found = True
                        break
            except Exception:
                pass

        if not btn_found:
            pw_ctrl.type_keys("{ENTER}")
            log("  ③ Enter 키로 로그인")

        return True

    except Exception as e:
        log(f"  [3순위] 실패: {e}")
        return False


def login_smes(_unused_win=None):
    """
    MES 자동 로그인 — 3단계 우선순위
      1순위: pywinauto auto_id (txt_id / txt_pw / btn_login)
      2순위: 창 기준 상대 좌표 (비율 방식)
      3순위: Edit 컨트롤 자동 스캔 (auto_id·좌표 불필요)
      최후: 수동 로그인 요청
    """
    import pyautogui

    log("  로그인 폼 탐색 중...")
    app, dlg = _find_login_dialog(timeout=20)

    # 창 활성화
    try:
        dlg.set_focus()
    except Exception:
        pass
    try:
        hwnd = dlg.handle
        _force_foreground(hwnd)
    except Exception:
        hwnd = None
    time.sleep(0.5)

    # ── 1순위: auto_id (텍스트 기반 컨트롤 식별) ──────────────────────────
    login_attempted = False
    if _login_by_autoid(dlg, hwnd):
        login_attempted = True
    # ── 2순위: 창 기준 상대 좌표 ────────────────────────────────────────────
    elif _login_by_coords(dlg, hwnd):
        login_attempted = True
    # ── 3순위: Edit 컨트롤 자동 스캔 ────────────────────────────────────────
    elif _login_by_edit_scan(dlg, hwnd):
        login_attempted = True
    # ── 최후: 수동 로그인 ───────────────────────────────────────────────────
    else:
        log("  ⚠️  자동 로그인 실패 → 수동으로 로그인해주세요.")
        input("  로그인 완료 후 Enter → ")
        time.sleep(LOAD_DELAY)
        return

    if not login_attempted:
        return

    # ── 로그인 성공 여부 확인 ────────────────────────────────────────────────
    log("  로그인 결과 확인 중...")
    deadline = time.time() + 12
    while time.time() < deadline:
        time.sleep(0.5)

        # ① 에러 팝업 먼저 감지 (가장 중요)
        err_msg = _check_and_dismiss_error_popup()
        if err_msg:
            raise RuntimeError(f"MES 로그인 오류 팝업: {err_msg}")

        # ② Login 버튼이 사라졌는지 확인 (가장 신뢰도 높음)
        try:
            _, login_btn = _find_login_controls(dlg)
            if not login_btn or not login_btn.is_visible():
                log("  ✅ 로그인 성공 (Login 버튼 비가시)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            pass

        # ③ Edit 컨트롤(ID/PW)이 사라졌는지 확인
        try:
            (id_c, pw_c), _ = _find_login_controls(dlg)
            if id_c is None and pw_c is None:
                log("  ✅ 로그인 성공 (로그인 패널 사라짐)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            pass

    raise RuntimeError("로그인 확인 실패 — Login 버튼이 12초 내에 사라지지 않음")

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

    # ── 생산관리 클릭 (C1MainMenu 대응) ────────────────────────────────────
    log("  생산관리 클릭...")
    clicked = False

    # Method 1: descendants() 전체 순회 - C1MainMenu 아이템 탐색
    if not clicked:
        try:
            for ctrl in main_win.descendants():
                try:
                    if ctrl.window_text().strip() == "생산관리":
                        ctrl.click_input()
                        clicked = True
                        log("  ✅ descendants 탐색 성공")
                        break
                except Exception:
                    pass
        except Exception as e:
            log(f"  descendants 탐색 실패: {e}")

    # Method 2: child_window title_re 탐색
    if not clicked:
        clicked = _try_click(main_win, ["생산관리"])
        if clicked:
            log("  ✅ child_window 탐색 성공")

    # Method 3: 창 좌표 기준 메뉴바 클릭 (생산관리 = 3번째 항목)
    # 이미지 기준: System(~40) → 기본정보관리(~100) → 생산관리(~160~175)
    # 최대화 창은 rect.top/left 가 음수 → max()로 화면 내 좌표 보정
    if not clicked:
        try:
            rect = main_win.rectangle()
            # 최대화 창: rect.top이 -8~-12 → 타이틀바(~22) 아래 메뉴바는 화면 Y=42 근처
            menu_y = max(rect.top + 42, 42)
            visible_left = max(rect.left, 0)
            for x_offset in [360, 355, 370, 345, 375]:
                menu_x = visible_left + x_offset
                pyautogui.click(menu_x, menu_y)
                time.sleep(0.5)
                clicked = True
                log(f"  ✅ 좌표 클릭: ({menu_x}, {menu_y})")
                break
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

    # ── 생산일자 설정 (Tab x2 → 날짜입력 → Tab x5 → Enter) ───────────────
    log(f"  생산일자 설정: {TODAY}")
    import pyautogui
    # Tab 5회: 생산일자 필드로 포커스 이동
    for _ in range(5):
        pyautogui.press('tab'); time.sleep(0.15)
    # 날짜 전체 선택 후 당일 날짜 입력 (YYYY-MM-DD)
    pyautogui.hotkey('ctrl', 'a'); time.sleep(0.1)
    pyautogui.typewrite(TODAY, interval=0.05)
    log(f"  날짜 입력 완료: {TODAY}")
    time.sleep(0.2)
    # Tab 5회: 조회 버튼으로 포커스 이동
    for _ in range(5):
        pyautogui.press('tab'); time.sleep(0.15)
    # Enter: 조회 실행
    pyautogui.press('enter')
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

    # row_height 단위로 각 행의 평균 '파란색 점수' 계산
    row_scores = {}
    for y in range(h):
        row_idx = y // row_height
        try:
            pixel = img.getpixel((5, y))
            r, g, b = pixel[0], pixel[1], pixel[2]
        except Exception:
            continue
        blue_score = b - (r + g) / 2   # 파란색이 강할수록 높음
        row_scores.setdefault(row_idx, []).append(blue_score)

    if not row_scores:
        return None

    # 평균 blue_score 최대 행
    best_row, scores = max(row_scores.items(), key=lambda kv: sum(kv[1]) / len(kv[1]))
    avg = sum(scores) / len(scores)
    if avg < 20:          # 임계값 미달 → 감지 실패
        return None

    selected_y = grid_top + best_row * row_height + row_height // 2
    return selected_y


def download_all_items(win):
    # 좌표 기준 첫 번째 행 클릭 후 Down 방식으로 전체 다운로드
    log("  좌표 기준 첫 번째 행 클릭 후 다운로드 시작...")
    count = _download_by_keyboard(win)
    log(f"  ✅ {count}개 파일 다운로드 완료")
    # 실제 파일 목록은 DOWNLOAD_DIR 스캔으로 반환 (automate_upload에서 사용)
    return []


def _click_excel_download(win, idx, item_name):
    """Excel 다운로드 버튼 클릭 및 저장"""
    import pyautogui
    import pyperclip
    import win32gui

    # 클릭 전 창 목록 수집
    hwnds_before = set()
    win32gui.EnumWindows(lambda h, _: hwnds_before.add(h), None)

    # Excel 버튼 클릭
    pyautogui.click(365, 538)
    log(f"    Excel 좌표 클릭: (365, 538)")

    # 저장 다이얼로그 등장 여부 확인 (최대 3초)
    dialog_hwnd = None
    for _ in range(6):
        time.sleep(0.5)
        hwnds_after = set()
        win32gui.EnumWindows(lambda h, _: hwnds_after.add(h), None)
        for h in hwnds_after - hwnds_before:
            try:
                title = win32gui.GetWindowText(h)
                if any(k in title for k in ["저장", "Save", "다른 이름"]):
                    dialog_hwnd = h
                    break
            except Exception:
                pass
        if dialog_hwnd:
            break

    if dialog_hwnd is None:
        log(f"    ⚠️  [{idx}] 저장 창 없음 → 키팅 자재 없음, 다음 품목으로")
        return False   # 다운로드 없음

    # 저장 다이얼로그 포커스 → Alt+D → Ctrl+V → Alt+S
    try:
        win32gui.SetForegroundWindow(dialog_hwnd)
    except Exception:
        pass
    time.sleep(0.4)

    save_folder = str(DOWNLOAD_DIR)
    pyperclip.copy(save_folder)
    pyautogui.hotkey('alt', 'd'); time.sleep(0.4)
    pyautogui.hotkey('ctrl', 'v'); time.sleep(0.3)
    pyautogui.hotkey('alt', 's')
    log(f"    ✅ [{idx}] 저장 완료 → {save_folder}")

    time.sleep(LOAD_DELAY)

    # Excel 현재 창 닫기 (Ctrl+W)
    try:
        excel_hwnds = []
        def _collect_excel(hwnd, _):
            title = win32gui.GetWindowText(hwnd)
            cls   = win32gui.GetClassName(hwnd)
            if win32gui.IsWindowVisible(hwnd) and (
                'Excel' in title or cls.startswith('XLMAIN')
            ):
                excel_hwnds.append(hwnd)
        win32gui.EnumWindows(_collect_excel, None)
        for hwnd in excel_hwnds:
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'w')
            time.sleep(0.8)
            # 저장 여부 묻는 경우 '저장 안 함(N)' 처리
            pyautogui.press('n')
            time.sleep(0.3)
            log(f"    Excel 창 닫기 완료")
    except Exception as e:
        log(f"    ⚠️  Excel 닫기 실패: {e}")

    return True   # 다운로드 성공


def _get_total_rows(win):
    """MES 창에서 '조회 결과 : N Rows' 텍스트를 읽어 총 행 수 반환. 실패 시 None."""
    import re
    try:
        for ctrl in win.descendants():
            try:
                txt = ctrl.window_text().strip()
                if not txt:
                    continue
                m = re.search(r'(\d+)\s*Rows', txt, re.IGNORECASE)
                if m:
                    return int(m.group(1))
            except Exception:
                pass
    except Exception:
        pass
    return None


def _download_by_keyboard(win):
    import pyautogui

    ROW_X      = 1000   # 품목명 컬럼 X (절대 좌표)
    ROW_Y      = 172    # 첫 번째 행 중심 Y — 2560x1600 스크린 실측값
    ROW_HEIGHT = 33     # 행 높이 — 2560x1600 스크린 실측값
    GRID_BOT   = 750    # 그리드 하단 경계 Y (스크롤 여유 포함)

    def row_click_y(idx):
        y = ROW_Y + idx * ROW_HEIGHT
        return min(y, GRID_BOT - ROW_HEIGHT // 2)

    # ── 총 행 수 읽기 ──────────────────────────────────────────────────────
    total_rows = _get_total_rows(win)
    if total_rows:
        log(f"  조회 결과: {total_rows}개 행 감지 → 정확히 {total_rows}번 반복")
    else:
        log("  ⚠️  행 수 감지 실패 → 스크린샷 비교 방식으로 폴백")

    # 첫 번째 행 클릭
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    pyautogui.click(ROW_X, ROW_Y)
    log(f"  첫 번째 행 클릭: ({ROW_X}, {ROW_Y})")
    time.sleep(0.5)

    downloaded = 0
    row_index  = 0
    max_iter   = total_rows if total_rows else 500

    for i in range(max_iter):
        check_stop()
        log(f"  [{i+1}/{max_iter}] Excel 다운로드 시도...")
        success = _click_excel_download(win, i + 1, "")
        if success:
            downloaded += 1

        # 마지막 행이면 Down 없이 종료
        if i >= max_iter - 1:
            log(f"  ✅ 마지막 행 완료 — 전체 {downloaded}개 다운로드 완료")
            break

        # MES 창 포커스 복귀 후 현재 행 클릭 → ↓
        try:
            win.set_focus()
        except Exception as e:
            log(f"    ⚠️  set_focus 실패({e})")
        time.sleep(0.4)

        cy = row_click_y(row_index)
        log(f"    행[{row_index+1}] 클릭: ({ROW_X}, {cy}) → ↓")
        pyautogui.click(ROW_X, cy)
        time.sleep(0.2)
        pyautogui.press('down')
        row_index += 1
        time.sleep(0.5)

    # total_rows 감지 실패 시 폴백: 스크린샷 비교로 추가 탐지
    if not total_rows:
        log("  (스크린샷 비교 폴백 — 행 수 미감지)")
        region = (ROW_X - 200, 145, 500, 615)
        same_count = 0
        for i in range(max_iter, 500):
            check_stop()
            log(f"  [{i+1}] Excel 다운로드 시도...")
            success = _click_excel_download(win, i + 1, "")
            if success:
                downloaded += 1
            try:
                win.set_focus()
            except Exception:
                pass
            time.sleep(0.4)
            before = pyautogui.screenshot(region=region)
            cy = row_click_y(row_index)
            pyautogui.click(ROW_X, cy)
            time.sleep(0.2)
            pyautogui.press('down')
            row_index += 1
            time.sleep(0.5)
            after = pyautogui.screenshot(region=region)
            if before.tobytes() == after.tobytes():
                same_count += 1
                log(f"    화면 변화 없음 ({same_count}/3)")
                if same_count >= 3:
                    log(f"  ✅ 마지막 행 도달 — 전체 {downloaded}개 완료")
                    break
            else:
                same_count = 0

    return downloaded

    return downloaded




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
    import win32gui, win32clipboard

    # 다운로드 전 현재 창 목록 수집
    hwnds_before = set()
    win32gui.EnumWindows(lambda h, _: hwnds_before.add(h), None)

    log("  Tab×1 → Enter → Excel 다운로드...")
    pyautogui.press('tab')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(EXCEL_DELAY)

    # ── 새로 생긴 Save As 다이얼로그 감지 ────────────────────────────────────
    file_name = f"재고현황{datetime.now().strftime('%H%M%S')}"
    save_folder = str(INVENTORY_DIR)
    dialog_hwnd = None

    for _ in range(20):
        hwnds_after = set()
        win32gui.EnumWindows(lambda h, _: hwnds_after.add(h), None)
        for h in hwnds_after - hwnds_before:
            try:
                title = win32gui.GetWindowText(h)
                if any(k in title for k in ["저장", "Save", "다른 이름"]):
                    dialog_hwnd = h
                    break
            except Exception:
                pass
        if dialog_hwnd:
            break
        time.sleep(0.5)

    if dialog_hwnd:
        log(f"  Save As 다이얼로그 감지 → {file_name}.xlsx 저장...")
        try:
            win32gui.SetForegroundWindow(dialog_hwnd)
        except Exception:
            pass
        time.sleep(0.4)

        # Alt+D → 주소창 → Ctrl+V(폴더경로) → Enter
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(save_folder, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1.5)

        # 파일명 입력 (Ctrl+A → 클립보드로 붙여넣기)
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(file_name, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)

        # Alt+S → 저장
        pyautogui.hotkey('alt', 's')
        time.sleep(1.5)

        # 덮어쓰기 확인 팝업
        try:
            from pywinauto import Desktop as _Desktop
            conf = _Desktop(backend='uia').window(title_re=".*(덮어|overwrite|Confirm).*")
            if conf.exists(timeout=2):
                pyautogui.press('enter')
                time.sleep(0.5)
        except Exception:
            pass

        log(f"  ✅ 저장 완료: {file_name}.xlsx")
    else:
        log("  ⚠️  Save As 다이얼로그 미감지 — Downloads 폴더 폴백")
        latest = _find_latest_download()
        if latest:
            import shutil
            dest = INVENTORY_DIR / f"{file_name}.xlsx"
            shutil.move(str(latest), str(dest))
            log(f"  ✅ 이동 완료: {dest.name}")
        else:
            log("  ⚠️  파일 저장 실패")

    # ── Excel 창 포커스 후 Ctrl+W 닫기 ──────────────────────────────────────
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
    return str(INVENTORY_DIR / f"{file_name}.xlsx")


# ──────────────────────────────────────────────────────────────────────────────
# Step 6 : 자재부족현황 Playwright 업로드
# ──────────────────────────────────────────────────────────────────────────────
def automate_upload(downloaded_files):
    from playwright.sync_api import sync_playwright

    log("▶ Step 6: 자재부족현황 자동 업로드...")

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

        # 키팅 로컬 상태 초기화 — const state는 window.state로 접근 불가, clearKitFiles() 함수 사용
        log("  키팅 로컬 상태 초기화...")
        try:
            page.evaluate("() => { if (typeof clearKitFiles === 'function') clearKitFiles(); }")
        except Exception as e:
            log(f"  ⚠️  초기화 평가 오류: {e}")
        time.sleep(1)

        # ── 키팅 자재 업로드 (먼저) ──────────────────────────────────────────
        all_kit = sorted(DOWNLOAD_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime)
        all_kit += sorted(DOWNLOAD_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime)
        valid = [str(f) for f in all_kit if f.exists()]
        log(f"  kitting 자재 폴더 전체 {len(valid)}개 파일 업로드 중...")
        if valid:
            page.locator("#file-kit").set_input_files(valid)

            # kit-chip 렌더링 + state.kitFiles 채워질 때까지 폴링 (최대 30초)
            log("  파일 처리 완료 대기 중...")
            deadline = time.time() + 30
            loaded = 0
            while time.time() < deadline:
                time.sleep(1)
                try:
                    loaded = page.evaluate("() => { try { return state.kitFiles.length; } catch(e) { return 0; } }")
                    if loaded >= len(valid):
                        break
                except Exception:
                    pass
            log(f"  state.kitFiles 로드 완료: {loaded}개")

            # handleKitFiles 내부에서 이미 Supabase 업로드가 시작됨 → 완료까지 추가 대기
            time.sleep(5)

            # upload_logs 동기화 확인 (선택적)
            try:
                result = page.evaluate("""
                    async () => {
                        const files = (typeof state !== 'undefined' && state.kitFiles) ? state.kitFiles : [];
                        if (!files.length) return { ok: false, reason: 'state.kitFiles 비어있음' };
                        const ts = Date.now();
                        const uname = (typeof currentUser !== 'undefined' && currentUser) ? currentUser.displayName || '' : '';
                        await syncUploadLog('kit', ts, uname, files.map(f => f.name));
                        localStorage.setItem('ms_uptime_kit', String(ts));
                        if (uname) localStorage.setItem('ms_uploader_kit', uname);
                        return { ok: true, count: files.length, names: files.map(f => f.name) };
                    }
                """)
                if result and result.get('ok'):
                    log(f"  ✅ 키팅 업로드 완료 ({result.get('count')}개): {result.get('names')}")
                else:
                    log(f"  ⚠️  키팅 업로드 결과: {result}")
            except Exception as e:
                log(f"  ⚠️  키팅 업로드 로그 오류: {e}")

        # ── 재고현황 업로드 (키팅 완료 후) ──────────────────────────────────
        inv_files = sorted(INVENTORY_DIR.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True)
        inv_files += sorted(INVENTORY_DIR.glob("*.xls"), key=lambda f: f.stat().st_mtime, reverse=True)
        if inv_files:
            latest_inv = str(inv_files[0])
            log(f"  재고현황 업로드: {inv_files[0].name}")
            page.locator("#file-inv").set_input_files(latest_inv)
            time.sleep(4)
            log("  ✅ 재고현황 업로드 완료")
        else:
            log("  ⚠️  재고현황 폴더에 파일 없음")

        # 완료 팝업 — () => {...} 함수 형태로 감싸야 SyntaxError 방지
        try:
            page.evaluate("""
                () => {
                    const el = document.createElement('div');
                    el.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.55);z-index:99999;display:flex;align-items:center;justify-content:center';
                    el.innerHTML = '<div style="background:white;border-radius:16px;padding:36px 52px;text-align:center;box-shadow:0 20px 60px rgba(0,0,0,0.35)">' +
                        '<div style="font-size:52px;margin-bottom:12px">\\u2705</div>' +
                        '<div style="font-size:22px;font-weight:700;color:#1e3a5f;margin-bottom:8px">\\uc2e4\\ud589\\uc644\\ub8cc \\ub418\\uc5c8\\uc2b5\\ub2c8\\ub2e4.</div>' +
                        '<div style="font-size:13px;color:#666;margin-bottom:20px">\\ud0a4\\ud305 \\ud30c\\uc77c\\uc774 \\uc790\\uc7ac\\ubd80\\uc871\\ud604\\ud669\\uc5d0 \\uc5c5\\ub85c\\ub4dc\\ub418\\uc5c8\\uc2b5\\ub2c8\\ub2e4.</div>' +
                        '<button onclick="this.parentElement.parentElement.remove()" style="padding:10px 36px;background:#1e3a5f;color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">\\ud655\\uc778</button>' +
                        '</div>';
                    document.body.appendChild(el);
                }
            """)
        except Exception as e:
            log(f"  ⚠️  완료 팝업 오류: {e}")

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
    parser = argparse.ArgumentParser()
    parser.add_argument('--step', default='all',
                        help='실행 단계: login | navigate | all')
    # mes-kit:// 프로토콜에서 전달되는 URL 인자 무시
    parser.add_argument('url', nargs='?', default='')
    args, _ = parser.parse_known_args()

    log("=" * 60)
    log(f"sMES 키팅 자동화 시작  ({TODAY})  [step={args.step}]")
    log("=" * 60)

    if not is_admin():
        log("관리자 권한 필요. 재실행 중...")
        time.sleep(1)
        elevate()
        return

    try:
        # Step 0: 기존 키팅 파일 전체 삭제
        log("▶ Step 0: 기존 키팅 파일 삭제...")
        deleted = 0
        for f in DOWNLOAD_DIR.iterdir():
            if f.is_file():
                try:
                    f.unlink()
                    deleted += 1
                except Exception as e:
                    log(f"  ⚠️  삭제 실패: {f.name} → {e}")
        log(f"  ✅ {deleted}개 파일 삭제 완료")

        # Step 1: sMES 실행
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

        log("")
        log("✅ sMES 로그인 완료!")

        if args.step == 'login':
            log("=" * 60)
            log("  [Step 1 완료] MES 실행 및 로그인 성공.")
            log("  다음 단계(메뉴 이동 → 다운로드)는 추후 추가 예정입니다.")
            log("=" * 60)
            input("  확인 후 Enter → ")
            return

        # Step 3~5: 키팅 자재 메뉴 이동 + 다운로드 (step=navigate 또는 all)
        log("▶ Step 4~5: 키팅 자재 메뉴 이동 및 다운로드...")
        downloaded = navigate_and_download(app)

        if args.step == 'navigate':
            log("=" * 60)
            log(f"  [Step 2 완료] 다운로드 {len(downloaded)}개 파일.")
            log("=" * 60)
            input("  확인 후 Enter → ")
            return

        # Step 5b: 재고현황 다운로드
        navigate_and_download_inventory(app)

        # Step 6: 업로드 (step=all)
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

                # 닫기 확인 팝업 자동 Enter
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

    except StopIteration:
        log("🛑 자동화가 긴급정지로 중단되었습니다.")
        STOP_FLAG.unlink(missing_ok=True)
    except Exception as e:
        log(f"❌ 오류: {e}")
        import traceback; traceback.print_exc()
        input("오류 확인 후 Enter → ")


if __name__ == "__main__":
    main()
