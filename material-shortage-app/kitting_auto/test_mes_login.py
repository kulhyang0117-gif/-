#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sMES 로그인 창 진단 & 테스트 스크립트
======================================
사용법:
  1. sMES를 실행해서 로그인 화면이 뜬 상태로 두세요.
  2. python test_mes_login.py            → 창 구조 진단만 (로그인 안 함)
  3. python test_mes_login.py --login    → 진단 후 실제 로그인 시도
  4. python test_mes_login.py --fg-test  → 포그라운드 검증 동작 테스트
"""

import sys, time, ctypes
from datetime import datetime

def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)

def hr(title=""):
    print(f"\n{'─'*60}")
    if title:
        print(f"  {title}")
        print(f"{'─'*60}")

# ──────────────────────────────────────────────────────────────
# 설정 (kitting_automation.py 와 동일하게 맞춰주세요)
# ──────────────────────────────────────────────────────────────
SMES_ID = "SSAT045"
SMES_PW = "rlatndus1!"

# ──────────────────────────────────────────────────────────────
# 테스트 1: sMES 프로세스 & 창 감지
# ──────────────────────────────────────────────────────────────
def test_find_process():
    hr("TEST 1: sMES 프로세스 감지")
    import psutil
    found = []
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            name = (proc.info['name'] or '').lower()
            exe  = (proc.info['exe']  or '').lower()
            if 'smes' in name or 'smes' in exe:
                found.append(proc)
                log(f"  ✅ 발견: PID={proc.info['pid']}  name={proc.info['name']}  exe={proc.info['exe']}")
        except Exception:
            pass
    if not found:
        log("  ❌ sMES 프로세스 없음 — sMES를 먼저 실행해주세요.")
    return found[0].info['pid'] if found else None


# ──────────────────────────────────────────────────────────────
# 테스트 2: 창 목록 & 크기 출력
# ──────────────────────────────────────────────────────────────
def test_list_windows(pid):
    hr("TEST 2: sMES 창 목록")
    from pywinauto import Application
    try:
        app = Application(backend='win32').connect(process=pid, timeout=5)
        wins = app.windows()
        log(f"  창 {len(wins)}개 발견:")
        for w in wins:
            r = w.rectangle()
            log(f"    title='{w.window_text()}'  "
                f"handle={w.handle}  "
                f"w={r.width()} h={r.height()}  "
                f"pos=({r.left},{r.top})")
        return app, wins
    except Exception as e:
        log(f"  ❌ 창 목록 실패: {e}")
        return None, []


# ──────────────────────────────────────────────────────────────
# 테스트 3: pnl_login 패널 & auto_id 컨트롤 존재 여부
# ──────────────────────────────────────────────────────────────
def test_autoid(app, wins):
    hr("TEST 3: auto_id 컨트롤 존재 확인")
    TARGET_IDS = ["pnl_login", "txt_id", "txt_pw", "btn_login"]

    for w in wins:
        log(f"\n  창: '{w.window_text()}' (handle={w.handle})")
        for aid in TARGET_IDS:
            try:
                ctrl = w.child_window(auto_id=aid)
                exists = ctrl.exists(timeout=0.5)
                visible = ctrl.is_visible() if exists else False
                status = "✅ 존재+가시" if (exists and visible) else \
                         "⚠️  존재+숨김" if exists else "❌ 없음"
                log(f"    auto_id='{aid}': {status}")
            except Exception as e:
                log(f"    auto_id='{aid}': ❌ 오류 — {e}")


# ──────────────────────────────────────────────────────────────
# 테스트 4: 전체 컨트롤 덤프 (auto_id 모를 때 확인용)
# ──────────────────────────────────────────────────────────────
def test_dump_controls(app, wins):
    hr("TEST 4: 전체 컨트롤 덤프 (로그인 창 추정)")
    # Edit 2개 이상이거나 pnl_login 있는 창을 로그인 창으로 추정
    login_win = None
    for w in wins:
        try:
            pnl = w.child_window(auto_id="pnl_login")
            if pnl.exists(timeout=0.3):
                login_win = w
                log(f"  pnl_login 방식으로 로그인 창 선택: '{w.window_text()}'")
                break
        except Exception:
            pass
        try:
            edits = w.children(class_name_re="WindowsForms10.EDIT.*")
            if len(edits) >= 2:
                login_win = w
                log(f"  Edit 방식으로 로그인 창 선택: '{w.window_text()}' (Edit {len(edits)}개)")
                break
        except Exception:
            pass

    if not login_win and wins:
        login_win = min(wins, key=lambda w: w.rectangle().width() * w.rectangle().height())
        log(f"  ⚠️  자동 선택 실패 → 가장 작은 창: '{login_win.window_text()}'")

    if not login_win:
        log("  ❌ 로그인 창 후보 없음")
        return None

    log(f"\n  컨트롤 목록:")
    try:
        for ctrl in login_win.descendants():
            try:
                r = ctrl.rectangle()
                log(f"    [{ctrl.element_info.control_type:12s}] "
                    f"auto_id='{ctrl.element_info.automation_id}'  "
                    f"class='{ctrl.class_name()}'  "
                    f"title='{ctrl.window_text()[:30]}'  "
                    f"pos=({r.left},{r.top})")
            except Exception:
                pass
    except Exception as e:
        log(f"  descendants 실패: {e}")
        try:
            log("  children 방식으로 재시도:")
            for ctrl in login_win.children():
                try:
                    log(f"    class='{ctrl.class_name()}'  title='{ctrl.window_text()[:30]}'")
                except Exception:
                    pass
        except Exception:
            pass

    return login_win


# ──────────────────────────────────────────────────────────────
# 테스트 5: 포그라운드 검증 동작 테스트
# ──────────────────────────────────────────────────────────────
def test_foreground(pid):
    hr("TEST 5: 포그라운드 검증 (_verify_foreground) 동작 테스트")
    from pywinauto import Application

    app = Application(backend='win32').connect(process=pid, timeout=5)
    wins = app.windows()
    if not wins:
        log("  ❌ 창 없음")
        return

    w = wins[0]
    hwnd = w.handle
    mes_hwnds = {hwnd}

    # 케이스 A: MES 창이 포그라운드일 때
    try:
        w.set_focus()
        time.sleep(0.5)
        fg = ctypes.windll.user32.GetForegroundWindow()
        result_a = fg in mes_hwnds
        log(f"  케이스 A (MES 포그라운드): fg_hwnd={fg}, mes_hwnd={hwnd} → {'✅ PASS' if result_a else '❌ FAIL'}")
    except Exception as e:
        log(f"  케이스 A 실패: {e}")

    # 케이스 B: 다른 창(탐색기)이 포그라운드일 때
    log("  케이스 B: 3초 후 다른 창을 클릭해서 포그라운드를 바꿔보세요...")
    for i in range(3, 0, -1):
        print(f"    {i}초...", end="\r", flush=True)
        time.sleep(1)
    print()
    fg = ctypes.windll.user32.GetForegroundWindow()
    result_b = fg not in mes_hwnds
    log(f"  케이스 B (다른 창 포그라운드): fg_hwnd={fg}, mes_hwnd={hwnd} → {'✅ PASS (차단 정상)' if result_b else '⚠️  여전히 MES 포그라운드'}")


# ──────────────────────────────────────────────────────────────
# 테스트 6: 실제 로그인 시도 (--login 옵션 시만)
# ──────────────────────────────────────────────────────────────
def test_do_login(login_win, hwnd):
    hr("TEST 6: 실제 로그인 시도")
    import pyautogui, pyperclip

    SAFE = set(
        "abcdefghijklmnopqrstuvwxyz"
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        "0123456789 !@#$%^&*()-_=+[]{}|;':\",./<>?"
    )

    def type_pw(pwd):
        if all(c in SAFE for c in pwd):
            pyautogui.typewrite(pwd, interval=0.06)
        else:
            pyperclip.copy(pwd)
            pyautogui.hotkey("ctrl", "v")

    def verify_fg(hwnd):
        mes_hwnds = {hwnd} if hwnd else set()
        if not mes_hwnds:
            log("  HWND 확인 불가 — 중단")
            return False
        fg = ctypes.windll.user32.GetForegroundWindow()
        if fg not in mes_hwnds:
            log(f"  포그라운드 아님 (fg={fg}) — 중단")
            return False
        return True

    # Caps Lock 해제
    if ctypes.windll.user32.GetKeyState(0x14) & 1:
        log("  Caps Lock 해제")
        pyautogui.press("capslock")

    # ── 1순위: auto_id ──
    log("  [1순위] auto_id 방식...")
    try:
        id_f = login_win.child_window(auto_id="txt_id")
        pw_f = login_win.child_window(auto_id="txt_pw")
        if id_f.exists(timeout=2) and pw_f.exists(timeout=2):
            if not verify_fg(hwnd): return
            id_f.click_input(); time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a"); pyperclip.copy(SMES_ID); pyautogui.hotkey("ctrl", "v")
            log(f"  ① ID 입력: {SMES_ID}")
            time.sleep(0.3)
            if not verify_fg(hwnd): return
            pw_f.click_input(); time.sleep(0.2)
            pyautogui.hotkey("ctrl", "a"); pyautogui.press("delete")
            type_pw(SMES_PW)
            log("  ② PW 입력 완료")
            time.sleep(0.3)
            if not verify_fg(hwnd): return
            btn = login_win.child_window(auto_id="btn_login")
            if btn.exists(timeout=1):
                btn.click_input()
                log("  ③ 버튼 클릭 (auto_id) → 로그인 시도 완료")
            else:
                pw_f.type_keys("{ENTER}")
                log("  ③ Enter 로 로그인 시도 완료")
            log("  ✅ [1순위] 로그인 입력 완료 — 결과를 직접 확인해주세요")
            return
        raise RuntimeError("auto_id 없음")
    except Exception as e:
        log(f"  [1순위] 실패: {e}")

    # ── 2순위: 좌표 방식 ──
    log("  [2순위] 창 기준 좌표 방식...")
    try:
        r = login_win.rectangle()
        left, top, w, h = r.left, r.top, r.width(), r.height()
        cx      = left + int(w * 0.509)
        id_pos  = (cx, top + int(h * 0.544))
        pw_pos  = (cx, top + int(h * 0.649))
        btn_pos = (left + int(w * 0.318), top + int(h * 0.765))
        log(f"  계산된 좌표 — ID:{id_pos}  PW:{pw_pos}  BTN:{btn_pos}")
        if not verify_fg(hwnd): return
        pyperclip.copy(SMES_ID)
        pyautogui.click(*id_pos); time.sleep(0.4)
        pyautogui.hotkey("ctrl", "a"); pyautogui.hotkey("ctrl", "v")
        log(f"  ① ID 입력: {SMES_ID}")
        if not verify_fg(hwnd): return
        pyautogui.press("tab"); time.sleep(0.3)
        pyautogui.hotkey("ctrl", "a"); pyautogui.press("delete")
        type_pw(SMES_PW)
        log("  ② PW 입력 완료")
        if not verify_fg(hwnd): return
        pyautogui.click(*btn_pos)
        log("  ③ 버튼 클릭 (좌표) → 로그인 시도 완료")
        log("  ✅ [2순위] 로그인 입력 완료 — 결과를 직접 확인해주세요")
    except Exception as e:
        log(f"  [2순위] 실패: {e}")
        log("  ❌ 자동 로그인 불가 — 수동으로 로그인해주세요")


# ──────────────────────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────────────────────
def main():
    do_login  = "--login"   in sys.argv
    do_fg     = "--fg-test" in sys.argv

    print("=" * 60)
    print("  sMES 로그인 창 진단 스크립트")
    if do_login:  print("  모드: 실제 로그인 시도")
    elif do_fg:   print("  모드: 포그라운드 검증 테스트")
    else:         print("  모드: 창 구조 진단만 (로그인 안 함)")
    print("=" * 60)

    # TEST 1: 프로세스
    pid = test_find_process()
    if not pid:
        print("\n  sMES를 실행하고 다시 시도해주세요.")
        input("  Enter → 종료")
        return

    # TEST 2: 창 목록
    app, wins = test_list_windows(pid)
    if not wins:
        input("  Enter → 종료")
        return

    # TEST 3: auto_id 확인
    test_autoid(app, wins)

    # TEST 4: 전체 컨트롤 덤프
    login_win = test_dump_controls(app, wins)

    # TEST 5: 포그라운드 검증 (--fg-test)
    if do_fg:
        test_foreground(pid)

    # TEST 6: 실제 로그인 (--login)
    if do_login and login_win:
        hwnd = login_win.handle
        # 창 활성화
        try:
            login_win.set_focus()
            fg_tid = ctypes.windll.user32.GetWindowThreadProcessId(
                ctypes.windll.user32.GetForegroundWindow(), None)
            my_tid = ctypes.windll.kernel32.GetCurrentThreadId()
            if fg_tid != my_tid:
                ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, True)
            ctypes.windll.user32.SetForegroundWindow(hwnd)
            ctypes.windll.user32.BringWindowToTop(hwnd)
            if fg_tid != my_tid:
                ctypes.windll.user32.AttachThreadInput(fg_tid, my_tid, False)
            time.sleep(0.5)
        except Exception as e:
            log(f"  창 활성화 실패: {e}")

        log(f"\n  3초 후 로그인을 시도합니다. MES 창을 건드리지 마세요...")
        for i in range(3, 0, -1):
            print(f"    {i}초...", end="\r", flush=True)
            time.sleep(1)
        print()
        test_do_login(login_win, hwnd)

    hr()
    print("  진단 완료.")
    print()
    print("  ※ auto_id가 모두 '❌ 없음'이면:")
    print("    → TEST 4의 컨트롤 덤프에서 실제 auto_id를 확인하세요.")
    print("    → kitting_automation.py의 auto_id 값을 수정해야 합니다.")
    print()
    print("  ※ 포그라운드 검증이 FAIL이면:")
    print("    → _force_foreground() 가 제대로 작동하지 않는 것입니다.")
    print("    → UAC 권한으로 스크립트를 실행해보세요.")
    input("\n  Enter → 종료")


if __name__ == "__main__":
    main()
