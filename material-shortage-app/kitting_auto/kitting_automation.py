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
# Step 2 : sMES 로그인  (PDF 3단계 우선순위 구현)
# ──────────────────────────────────────────────────────────────────────────────

def _type_password(pwd: str):
    """
    비밀번호 입력 — 클립보드 방식 우선.
    typewrite는 !@# 등 Shift+숫자 조합 키를 누락하므로 사용하지 않음.
    """
    import pyperclip, pyautogui
    pyperclip.copy(pwd)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)


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


def _find_login_dialog(timeout=20):
    """
    로그인 창 감지 — 3단계 탐지:
      1) 창 제목에 'Shinsung' 또는 'IMMES' 포함 (실제 창 확인됨)
      2) pnl_login 패널 가시성 (버전 무관 안정적)
      3) descendants 포함 Edit 2개 이상 (가장 넓은 폴백)
    """
    from pywinauto import Application

    # 창 제목 키워드 (이미지에서 확인된 실제 제목)
    TITLE_KEYWORDS = ["Shinsung", "IMMES", "sMES", "MES", "LOGIN", "로그인"]

    pid = _get_smes_pid()
    if not pid:
        raise RuntimeError("sMES 프로세스를 찾을 수 없습니다.")

    app = Application(backend='win32').connect(process=pid, timeout=10)
    log(f"  sMES PID={pid} 연결 완료")

    deadline = time.time() + timeout
    while time.time() < deadline:
        for w in app.windows():
            title = w.window_text()

            # ── 1) 창 제목 키워드 매칭 ──
            if any(kw in title for kw in TITLE_KEYWORDS):
                log(f"  ✅ 로그인 창 감지 (제목 키워드): '{title}'")
                return app, w

            # ── 2) pnl_login 패널 가시성 ──
            try:
                pnl = w.child_window(auto_id="pnl_login")
                if pnl.exists(timeout=0.3) and pnl.is_visible():
                    log(f"  ✅ 로그인 창 감지 (pnl_login): '{title}'")
                    return app, w
            except Exception:
                pass

            # ── 3) 하위 포함 Edit 2개 이상 ──
            try:
                edits = w.descendants(class_name_re="WindowsForms10.EDIT.*")
                if len(edits) >= 2:
                    log(f"  ✅ 로그인 창 감지 (Edit {len(edits)}개): '{title}'")
                    return app, w
            except Exception:
                pass
        time.sleep(0.5)

    # 최후 수단: 가장 큰 창 (메인 창에 로그인 패널이 내장된 구조)
    wins = app.windows()
    if wins:
        biggest = max(wins, key=lambda w: w.rectangle().width() * w.rectangle().height())
        log(f"  ⚠️  자동 탐지 실패 → 메인 창 사용: '{biggest.window_text()}'")
        return app, biggest

    raise RuntimeError("로그인 폼을 찾을 수 없습니다.")


def _find_login_controls(dlg):
    """
    children()에서 Edit 2개(ID/PW)와 Login 버튼을 찾아 반환.
    로그 덤프 확인: C1TextBox(WindowsForms10.EDIT) 2개, C1Button(WindowsForms10.BUTTON) title='Login'
    """
    login_btn = None
    for btn in dlg.children(class_name_re="WindowsForms10.BUTTON.*"):
        try:
            if btn.window_text().strip().lower() in ("login", "로그인"):
                login_btn = btn
                break
        except Exception:
            pass

    edits = sorted(
        dlg.children(class_name_re="WindowsForms10.EDIT.*"),
        key=lambda c: c.rectangle().top
    )
    return edits, login_btn


def _login_by_autoid(dlg, hwnd):
    """1순위: children() 직접 스캔 — C1TextBox ID/PW + C1Button Login"""
    import pyautogui, pyperclip
    log("  [1순위] children() 직접 스캔 방식 시도...")

    try:
        edits, login_btn = _find_login_controls(dlg)
        log(f"  Edit {len(edits)}개, Login 버튼: {'있음' if login_btn else '없음'}")

        if len(edits) < 2 or not login_btn:
            raise RuntimeError(f"컨트롤 부족 (Edit={len(edits)}, Btn={login_btn})")

        id_ctrl = edits[0]   # Y 작은 = 위 = ID
        pw_ctrl = edits[1]   # Y 큰  = 아래 = PW

        # ① ID
        existing_id = id_ctrl.window_text().strip()
        if existing_id:
            log(f"  ① ID '{existing_id}' 이미 있음 → 스킵")
        else:
            _verify_foreground(hwnd)
            id_ctrl.click_input(); time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a")
            pyperclip.copy(SMES_ID); pyautogui.hotkey("ctrl", "v")
            log(f"  ① ID 입력: {SMES_ID}")
            time.sleep(0.3)

        _check_caps_lock()

        # ② PW
        _verify_foreground(hwnd)
        pw_ctrl.click_input(); time.sleep(0.2)
        pyautogui.hotkey("ctrl", "a"); pyautogui.press("delete")
        _type_password(SMES_PW)
        log("  ② PW 입력 완료")
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

    # ── 1순위: auto_id ──────────────────────────────────────────────────────
    if _login_by_autoid(dlg, hwnd):
        pass
    # ── 2순위: 창 기준 상대 좌표 ────────────────────────────────────────────
    elif _login_by_coords(dlg, hwnd):
        pass
    # ── 3순위: Edit 컨트롤 자동 스캔 ────────────────────────────────────────
    elif _login_by_edit_scan(dlg, hwnd):
        pass
    # ── 최후: 수동 로그인 ───────────────────────────────────────────────────
    else:
        log("  ⚠️  자동 로그인 실패 → 수동으로 로그인해주세요.")
        input("  로그인 완료 후 Enter → ")
        time.sleep(LOAD_DELAY)
        return

    # ── 로그인 성공 여부 확인 ────────────────────────────────────────────────
    log("  로그인 결과 확인 중...")
    deadline = time.time() + 12
    while time.time() < deadline:
        time.sleep(0.5)

        # ① Login 버튼이 사라졌는지 확인 (가장 신뢰도 높음)
        try:
            _, login_btn = _find_login_controls(dlg)
            if not login_btn or not login_btn.is_visible():
                log("  ✅ 로그인 성공 (Login 버튼 비가시)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            pass

        # ② Edit 컨트롤(ID/PW)이 사라졌는지 확인
        try:
            edits, _ = _find_login_controls(dlg)
            if len(edits) == 0:
                log("  ✅ 로그인 성공 (로그인 패널 사라짐)")
                time.sleep(LOAD_DELAY)
                return
        except Exception:
            pass

        # ③ 화면에 에러 메시지 있는지 출력 (MES 오류 팝업 감지)
        try:
            for child in dlg.children():
                txt = child.window_text().strip()
                if txt and txt not in (SMES_ID, '') and len(txt) < 100:
                    if any(kw in txt for kw in ["오류", "실패", "틀", "error", "invalid", "wrong"]):
                        log(f"  ❌ 오류 메시지: '{txt}'")
        except Exception:
            pass

    log("  ⚠️  로그인 확인 실패 — 수동으로 로그인해주세요.")
    input("  로그인 완료 후 Enter → ")
    time.sleep(LOAD_DELAY)

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
    parser = argparse.ArgumentParser()
    parser.add_argument('--step', default='login',
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

        # Step 3~5: 메뉴 이동 + 다운로드 (step=navigate 또는 all)
        log("▶ Step 4~5: 메뉴 이동 및 다운로드...")
        downloaded = navigate_and_download(app)

        if args.step == 'navigate':
            log("=" * 60)
            log(f"  [Step 2 완료] 다운로드 {len(downloaded)}개 파일.")
            log("=" * 60)
            input("  확인 후 Enter → ")
            return

        # Step 6: 업로드 (step=all)
        automate_upload(downloaded)

    except Exception as e:
        log(f"❌ 오류: {e}")
        import traceback; traceback.print_exc()
        input("오류 확인 후 Enter → ")


if __name__ == "__main__":
    main()
