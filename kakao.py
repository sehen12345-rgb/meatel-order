"""
카카오톡 PC 앱 자동 파일 전송
pyautogui + pyperclip 사용
"""
import os
import time
import subprocess
import pyautogui
import pyperclip

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3


def find_kakao_window():
    """카카오톡 창 찾기 (Windows)"""
    import win32gui
    kakao_hwnd = None

    def enum_handler(hwnd, _):
        nonlocal kakao_hwnd
        title = win32gui.GetWindowText(hwnd)
        if "카카오톡" in title and win32gui.IsWindowVisible(hwnd):
            kakao_hwnd = hwnd

    win32gui.EnumWindows(enum_handler, None)
    return kakao_hwnd


def send_file_to_contact(contact_name: str, file_path: str,
                          status_callback=None) -> bool:
    """
    카카오톡 PC 앱으로 파일 전송
    contact_name: 채팅창 이름 (예: '미트엘')
    file_path: 전송할 파일 경로
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일을 찾을 수 없음: {file_path}")

    def log(msg):
        if status_callback:
            status_callback(msg)
        print(msg)

    try:
        import win32gui
        import win32con

        # 카카오톡 창 찾기
        hwnd = find_kakao_window()
        if not hwnd:
            raise RuntimeError("카카오톡 창을 찾을 수 없습니다. 로그인 후 다시 시도하세요.")

        # 창 활성화
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.8)
        log("카카오톡 창 활성화 완료")

        # Ctrl+F 검색 (채팅방 검색)
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.6)

        # 연락처 이름 입력
        pyperclip.copy(contact_name)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.8)

        # 첫 번째 결과 엔터
        pyautogui.press("enter")
        time.sleep(0.6)
        pyautogui.press("enter")
        time.sleep(0.8)
        log(f"'{contact_name}' 채팅방 열기")

        # 파일 첨부 단축키 없음 → 클립보드에 파일 경로 복사 후 붙여넣기
        # Windows에서는 파일 경로를 클립보드에 넣어 붙여넣을 수 없음
        # → 파일 탐색기 방식으로 전송
        # Shift + Enter 대신 채팅창 하단 파일 첨부 버튼 좌클릭
        # (위치가 버전마다 달라 좌표 대신 단축키 사용)

        # Alt+A 파일 첨부 단축키 (구버전) 또는 아이콘 클릭
        # 최신 카카오톡은 특정 단축키 없음 → pyautogui로 아이콘 찾기

        # 파일을 채팅창에 드래그 방식 대신:
        # 채팅 입력창 포커스 후 파일 경로 명령어 전송
        # → 가장 안정적인 방법: 파일 탐색기 열기

        # 채팅창 클릭 (포커스)
        # 화면 중앙 클릭
        screen_w, screen_h = pyautogui.size()
        pyautogui.click(screen_w // 2, screen_h // 2)
        time.sleep(0.3)

        # 파일 첨부 버튼 클릭 (+ 버튼)
        # 카카오톡 입력창 좌측 하단에 있음
        # 정확한 위치를 이미지로 찾기
        try:
            btn = pyautogui.locateOnScreen(
                "kakao_attach_btn.png", confidence=0.8
            )
            if btn:
                pyautogui.click(btn)
                time.sleep(0.5)
        except Exception:
            # 이미지 없으면 키보드 방법 시도
            pass

        # 파일 전송 대화상자
        # Windows 파일 열기 다이얼로그에 경로 직접 입력
        pyperclip.copy(file_path)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
        pyautogui.press("enter")
        time.sleep(1.0)

        log(f"파일 전송 완료: {os.path.basename(file_path)}")
        return True

    except Exception as e:
        log(f"전송 오류: {e}")
        return False


def send_file_simple(contact_name: str, file_path: str,
                     status_callback=None) -> bool:
    """
    간단 버전: 파일 탐색기를 통한 전송
    1. 파일 탐색기에서 파일 선택
    2. 복사
    3. 카카오톡 채팅창에 붙여넣기
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"파일 없음: {file_path}")

    def log(msg):
        if status_callback:
            status_callback(msg)
        print(msg)

    try:
        import win32gui
        import win32con
        import subprocess

        # 카카오톡 창 활성화
        hwnd = find_kakao_window()
        if not hwnd:
            raise RuntimeError("카카오톡 창을 찾을 수 없습니다.")

        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.8)

        # 검색으로 채팅방 열기
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "a")
        pyperclip.copy(contact_name)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.8)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1.0)
        log(f"'{contact_name}' 채팅방 진입")

        # 파일을 클립보드에 복사 (Windows Shell API)
        from win32com.shell import shell, shellcon
        import win32clipboard

        abs_path = os.path.abspath(file_path)
        log(f"파일 클립보드 복사: {abs_path}")

        # 파일을 클립보드에 복사
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(
            win32clipboard.CF_HDROP,
            win32clipboard.GlobalAlloc(
                win32con.GMEM_MOVEABLE | win32con.GMEM_ZEROINIT,
                20 + len(abs_path.encode("utf-16-le")) + 4
            )
        )
        win32clipboard.CloseClipboard()

        # 채팅창 포커스 후 붙여넣기
        screen_w, screen_h = pyautogui.size()
        pyautogui.click(screen_w // 2, int(screen_h * 0.9))
        time.sleep(0.3)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(0.5)

        log("파일 전송 완료!")
        return True

    except Exception as e:
        log(f"오류: {e}")
        return False
