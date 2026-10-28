import os

# cryptography 는 OpenSSL 레거시 프로바이더를 못 찾으면 import 단계에서 바로 죽는다.
# (EXE 로 묶었을 때 실제 발생: "OpenSSL 3.0's legacy provider failed to load")
# 이 프로그램은 Fernet(AES-CBC + HMAC)만 쓰고 레거시 알고리즘은 쓰지 않으므로 꺼둔다.
# cryptography 를 import 하기 전에 설정해야 효과가 있다.
os.environ.setdefault("CRYPTOGRAPHY_OPENSSL_NO_LEGACY", "1")

import json
import tkinter as tk
from tkinter import messagebox, colorchooser, ttk, simpledialog, Menu
from functools import partial
import sys
import subprocess
from tkinter import filedialog
import time
from datetime import datetime
import re
import ast
import threading
import hashlib
import uuid
import webbrowser
from tkinterdnd2 import *
from cryptography.fernet import Fernet
import base64
from ttkthemes import ThemedTk
import requests
from icalendar import Calendar
import calendar
from tkcalendar import Calendar as TkCalendar  # tkcalendar의 Calendar에 별칭 부여
import shutil  # 추가된 부분
import platform
import random
from win32com.client import Dispatch
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font



# 표준 출력 인코딩을 UTF-8 로 고정
# 한글 윈도우의 기본 콘솔 인코딩(cp949)은 폴더 버튼 이름에 붙는 📁(U+1F4C1) 같은
# 이모지를 출력하지 못한다. 그대로 두면 디버그 print 에서 UnicodeEncodeError 가 나고,
# 그 예외가 MyButton.__init__ 밖으로 전파되어 폴더 버튼 생성이 통째로 실패한다.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass  # 콘솔 없는 실행(--windowed) 등에서는 무시

# JSON 파일 저장을 위한 디렉토리 경로 설정
exe_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
json_directory = f"C:/json/{exe_name}"
if not os.path.exists(json_directory):
    os.makedirs(json_directory)
COPY_FILE_PATH = f"{json_directory}/copied_button.json"
LAST_OPEN_FILE = f"{json_directory}/last_youtube_open.json"
print(f"[DEBUG] 실행 파일 이름: {exe_name}, JSON 디렉토리: {json_directory}")

# 암호화 키 로드 또는 생성 (비밀번호 암호화를 위한 키 관리)
key_file = os.path.join(json_directory, "encryption_key.key")
if os.path.exists(key_file):
    with open(key_file, "rb") as f:
        encryption_key = f.read()
else:
    encryption_key = Fernet.generate_key()
    with open(key_file, "wb") as f:
        f.write(encryption_key)
cipher = Fernet(encryption_key)

# 전역 변수
selected_buttons = []
is_dragging = False
is_button_long_pressed = False
press_time = 0
press_x = 0
press_y = 0
selected_labels = []  # 추가: 다중 선택된 라벨 저장
calendar_window = None  # 달력 창을 추적하기 위한 전역 변수
calendar_app = None    # ModernCalendarApp 인스턴스를 추적하기 위한 전역 변수
edit_window = None
show_cmd = False  # CMD 창 기본 숨김
selected_buttons = []
show_cmd = False  # 기본적으로 cmd 창을 표시하지 않음
KNOWN_MODULE_ALIASES = {}
debug_window = None  # 디버깅 창 추적
debug_text = None    # 디버깅 텍스트 위젯
copied_button_data = None  # 추가: 복사된 버튼 데이터 저장
drag_target_type = None  # 드래그 대상 유형: 'button', 'label', 또는 None
PROGRAM_VERSION = None
grid_on = False
copied_label_data = None  # 복사된 라벨 데이터를 저장

COPY_FILE_PATH = "C:/json/shared_copied_button.json"


def send_to_nateon(username, program_name, location, access_time, program_version):
    url = "https://teamroom.nate.com/api/webhook/be7e8ea5/C53VcbV4FeBhLDtCkw2EvH6t"
    
    # 메시지 구성
    message = (
        f"User: {username}, Program: {program_name}, "
        f"Location: {location}, Time: {access_time}, Version: {program_version}"
    )
    
    # JSON 페이로드
    payload = {"message": message}
    
    try:
        # POST 요청 전송
        response = requests.post(url, json=payload, timeout=5)
        response.raise_for_status()  # HTTP 오류 발생 시 예외 발생
    except requests.exceptions.RequestException:
        pass  # 오류 무시
        

# 암호화/복호화 함수
def encrypt_data(data):
    if not data:  # 빈 문자열이면 암호화하지 않음
        return ""
    try:
        encrypted = cipher.encrypt(data.encode())
        return base64.b64encode(encrypted).decode()  # JSON 저장을 위해 Base64 인코딩
    except Exception as e:
        print(f"암호화 중 오류: {e}")
        return data  # 오류 시 평문 반환 (보안 강화를 위해 수정 가능)

def decrypt_data(encrypted_data):
    if not encrypted_data:  # 빈 문자열이면 복호화하지 않음
        return ""
    try:
        decoded = base64.b64decode(encrypted_data)
        decrypted = cipher.decrypt(decoded).decode()
        return decrypted
    except Exception as e:
        print(f"복호화 중 오류: {e}")
        return encrypted_data  # 오류 시 암호화된 상태로 반환 (보안 강화를 위해 수정 가능)

def sanitize_filename(filename):
    # "📁"와 같은 특수 이모지 제거 후, 파일 이름에 사용할 수 없는 문자 제거
    clean_filename = filename.replace(" 📁", "").replace(" ↲", "")  # "📁" 명시적 제거
    return re.sub(r'[\/:*?"<>|]', '', clean_filename)  # 파일 시스템에서 금지된 문자 제거


# 윈도우가 바로가기에 붙이는 꼬리말. 버튼 이름에서는 떼어낸다.
_SHORTCUT_SUFFIX_RE = re.compile(r"\s*[-\u2013\u2014]?\s*(바로\s*가기|Shortcut)\s*$", re.IGNORECASE)


def clean_shortcut_name(name):
    """'크롬 - 바로가기' -> '크롬'. 다 지워지면 원래 이름을 그대로 둔다."""
    cleaned = _SHORTCUT_SUFFIX_RE.sub("", name).strip()
    return cleaned or name


def _new_label_uid():
    """라벨 JSON 파일명에 사용할 고유 ID."""
    return uuid.uuid4().hex[:12]


def label_uid(label):
    """라벨 위젯의 고유 ID 반환(없으면 새로 부여).

    라벨 JSON 파일명을 '텍스트'가 아니라 이 ID 로 만들기 때문에
    같은 텍스트의 라벨을 여러 개 만들어도 서로 덮어쓰지 않는다.
    """
    uid = getattr(label, "uid", None)
    if not uid:
        uid = _new_label_uid()
        try:
            label.uid = uid
        except Exception:
            pass
    return uid


def legacy_label_json_path(label_name):
    """예전(텍스트 기반) 라벨 JSON 경로."""
    sanitized = sanitize_filename(label_name).replace("\\", "_").replace(" ", "_").replace(chr(10), "")
    return os.path.join(json_directory, f"{sanitized}_label.json")


def label_json_paths(label):
    """해당 라벨의 JSON 경로 후보(uid 기반 + 레거시 텍스트 기반)."""
    paths = []
    uid = getattr(label, "uid", None)
    if uid:
        paths.append(os.path.join(json_directory, f"{uid}_label.json"))
    try:
        paths.append(legacy_label_json_path(label.cget("text")))
    except Exception:
        pass
    return list(dict.fromkeys(paths))


def delete_label_json(label):
    """라벨의 JSON 파일을 지운다. 하나라도 지워졌으면 True."""
    removed = False
    for path in label_json_paths(label):
        if os.path.exists(path):
            try:
                os.remove(path)
                removed = True
                print(f"[DEBUG] 라벨 JSON 삭제됨: {path}")
            except OSError as e:
                print(f"[ERROR] 라벨 JSON 삭제 실패: {path} - {e}")
    return removed

def update_json_file(name, json_directory, new_data):
    sanitized_name = sanitize_filename(name)
    file_path = os.path.join(json_directory, f"{sanitized_name}_button.json")
    default_data = {
        "name": name,
        "path": "",
        "x": 0,
        "y": 0,
        "color": "#FFFFFF",
        "password": "",
        "font_size": 10,
        "fg": "white",
        "font_weight": "normal",
        "width": 15,
        "height": 1,
        "hotkey": None,
        "is_folder": False  # 기본값 추가
    }
    existing_data = default_data.copy()
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                file_data = json.load(file)
                existing_data.update(file_data)
        except json.JSONDecodeError:
            print(f"[WARNING] 기존 파일 파싱 실패, 기본값 사용: {file_path}")
    
    if "password" in new_data and new_data["password"]:
        new_data["password"] = encrypt_data(new_data["password"])
    
    if "fg" not in new_data:
        new_data["fg"] = existing_data["fg"]
    
    # "is_folder" 상태를 new_data에 반영
    if "is_folder" in new_data:
        existing_data["is_folder"] = new_data["is_folder"]
    
    existing_data.update(new_data)
    try:
        with open(file_path, "w", encoding="utf-8") as file:
            json.dump(existing_data, file, indent=4)
        print(f"[DEBUG] JSON 파일 저장 성공: {file_path}")
    except Exception as e:
        print(f"[ERROR] JSON 파일 저장 중 오류 발생: {file_path} - {e}")
        
        
        



def delete_selected_items(event=None):
    global selected_buttons, selected_labels
    
    # 선택된 버튼 또는 라벨이 있는지 확인
    if not selected_buttons and not selected_labels:
        print("[DEBUG] 삭제할 버튼 또는 라벨이 선택되지 않음")
        return
    
    # 삭제 확인 메시지
    if not messagebox.askyesno("삭제 확인", "선택된 버튼 및 라벨을 삭제하시겠습니까?"):
        print("[DEBUG] 삭제가 사용자에 의해 취소됨")
        return
    
    # 선택된 버튼 삭제
    for button in selected_buttons[:]:  # 리스트 복사본 사용
        button_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")

        # button.name이 json_key(해시 포함)인 경우와 일반 이름인 경우 모두 처리
        json_key = getattr(button, "name", sanitize_filename(button_name))
        json_path = os.path.join(json_directory, f"{json_key}_button.json")

        # 후보 경로: 해시 기반 키, 정규 이름, 그리고 "레거시" 파일명.
        # sanitize_filename 은 " 📁" 를 지우므로, 레거시 파일을 지우려면
        # 원본 텍스트를 가공하지 않은 형태도 후보에 넣어야 한다.
        raw_text = button.cget("text")
        candidate_paths = [
            json_path,
            os.path.join(json_directory, f"{sanitize_filename(button_name)}_button.json"),
            os.path.join(json_directory, f"{sanitize_filename(raw_text)}_button.json"),
            os.path.join(json_directory, f"{raw_text}_button.json"),          # 예: "이름 📁_button.json"
            os.path.join(json_directory, f"{getattr(button, 'name', '')}_button.json"),
        ]
        for path in dict.fromkeys(candidate_paths):  # 중복 제거하면서 순서 유지
            if os.path.exists(path):
                try:
                    os.remove(path)
                    print(f"[DEBUG] JSON 파일 삭제됨: {path}")
                except OSError as e:
                    print(f"[ERROR] JSON 파일 삭제 실패: {path} - {e}")

        # UI 및 상태에서 버튼 제거
        button.destroy()
        if button in selected_buttons:
            selected_buttons.remove(button)
        print(f"[DEBUG] 버튼 삭제됨: {button_name}, path={button.path}")
    
    # 선택된 라벨 삭제
    for label in selected_labels[:]:  # 리스트 복사본 사용
        label_name = label.cget("text")

        # JSON 파일 삭제 (uid 기반 파일명 + 예전 텍스트 기반 파일명 모두 처리)
        if not delete_label_json(label):
            print(f"[WARNING] 삭제 대상 JSON 파일이 존재하지 않음: {label_name}")
        
        # UI 및 상태에서 라벨 제거
        label.destroy()
        if label in selected_labels:
            selected_labels.remove(label)
        print(f"[DEBUG] 라벨 삭제됨: {label_name}")
    
    # UI 업데이트
    root.update_idletasks()
    print("[DEBUG] 선택된 항목 삭제 완료")
    
    
    
    
    
    
    

def set_button_size(button, width_mm, height_mm):
    width_px = int(width_mm * 3.78)
    height_px = int(height_mm * 3.78)
    button.config(width=width_px, height=height_px)

def auto_change_slash(path):
    # NAS 경로 처리: IP 주소나 서버 이름으로 시작하거나 이미 \\로 시작하는 경우
    nas_pattern = re.compile(r'^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+\\|^[a-zA-Z0-9-]+\\|^\\\\[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+|^\\\\[a-zA-Z0-9-]+')
    if nas_pattern.search(path):
        normalized_path = path.replace("/", "\\")  # 슬래시를 백슬래시로 변환
        if not normalized_path.startswith("\\\\"):
            return "\\\\" + os.path.normpath(normalized_path)  # \\로 시작하도록 보장
        return normalized_path  # 이미 \\로 시작하면 유지
    
    # URL 처리 (기존 동작 유지)
    if path.startswith("http://") or path.startswith("https://"):
        return path.replace("\\", "/")
    
    # 일반 경로 처리
    # os.path.normpath 는 UNC 경로(\server\share)의 선행 \ 를 그대로 유지한다.
    # 여기서 앞 두 글자를 잘라내면 nas_pattern 에 걸리지 않은 UNC 경로
    # (예: \my_server\share 처럼 서버명에 _ 가 들어간 경우)가 깨지므로 잘라내지 않는다.
    return os.path.normpath(path)


# 동적으로 업데이트되는 모듈-패키지 매핑


def install_missing_modules(python_executable, modules, root=None):
    if root is None:
        root = tk.Tk()
        root.withdraw()  # GUI 창 숨기기

    all_installed = True
    for module in modules:
        try:
            # 모듈이 이미 설치되어 있는지 확인
            subprocess.check_call([python_executable, "-c", f"import {module}"],
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                 creationflags=subprocess.CREATE_NO_WINDOW)  # CMD 창 숨기기
            print(f"[DEBUG] 모듈 {module} 실제로 사용 가능")
        except (subprocess.CalledProcessError, ImportError) as e:
            print(f"[DEBUG] 모듈 {module} 설치 시도 중...")
            package_name = KNOWN_MODULE_ALIASES.get(module, module)
            install_success = False

            # 1차 시도: 현재 패키지 이름으로 설치
            try:
                result = subprocess.run([python_executable, "-m", "pip", "install", package_name, "--user"],
                                       capture_output=True, text=True, check=True,
                                       creationflags=subprocess.CREATE_NO_WINDOW)  # CMD 창 숨기기
                print(f"[DEBUG] 모듈 {module} 설치 성공 (패키지: {package_name}): {result.stdout}")
                install_success = True
            except subprocess.CalledProcessError as e:
                error_msg = f"모듈 {module} 설치 실패 (패키지: {package_name}):\nSTDERR: {e.stderr}"
                print(f"[DEBUG] {error_msg}")

            # 2차 시도: 잘 알려진 대체 이름으로 자동 설치
            if not install_success:
                fallback_packages = {
                    'PIL': 'Pillow',
                    'cv2': 'opencv-python'
                }
                if module in fallback_packages:
                    package_name = fallback_packages[module]
                    print(f"[DEBUG] 모듈 {module}을 {package_name}으로 자동 재시도 중...")
                    try:
                        result = subprocess.run([python_executable, "-m", "pip", "install", package_name, "--user"],
                                               capture_output=True, text=True, check=True,
                                               creationflags=subprocess.CREATE_NO_WINDOW)  # CMD 창 숨기기
                        print(f"[DEBUG] 모듈 {module} 설치 성공 (패키지: {package_name}): {result.stdout}")
                        KNOWN_MODULE_ALIASES[module] = package_name
                        install_success = True
                    except subprocess.CalledProcessError as e:
                        error_msg = f"모듈 {module} 설치 실패 (패키지: {package_name}):\nSTDERR: {e.stderr}"
                        print(f"[DEBUG] {error_msg}")

            # 3차 시도: 사용자 입력으로 대체 설치
            if not install_success:
                retry_package = simpledialog.askstring(
                    "설치 실패",
                    f"{error_msg}\n\n'{module}' 대신 설치할 올바른 패키지 이름을 입력하세요 (취소 시 건너뜀):",
                    initialvalue=KNOWN_MODULE_ALIASES.get(module, module),
                    parent=root
                )
                if retry_package and retry_package != package_name:
                    try:
                        result = subprocess.run([python_executable, "-m", "pip", "install", retry_package, "--user"],
                                               capture_output=True, text=True, check=True,
                                               creationflags=subprocess.CREATE_NO_WINDOW)  # CMD 창 숨기기
                        print(f"[DEBUG] 모듈 {module} 대체 설치 성공 (패키지: {retry_package}): {result.stdout}")
                        KNOWN_MODULE_ALIASES[module] = retry_package
                        install_success = True
                    except subprocess.CalledProcessError as retry_e:
                        print(f"[DEBUG] 대체 설치 실패: {retry_e.stderr}")
                        messagebox.showerror("오류", f"대체 패키지 {retry_package} 설치 실패:\nSTDERR: {retry_e.stderr}", parent=root)
                        all_installed = False
                else:
                    messagebox.showwarning("경고", f"{module} 설치가 건너뛰어졌습니다. 스크립트 실행에 문제가 발생할 수 있습니다.", parent=root)
                    all_installed = False

            if not install_success:
                all_installed = False

    if root.master is None:
        root.destroy()
    return all_installed



def get_imported_modules(file_path):
    imported_modules = set()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            tree = ast.parse(content)
            
            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    for name in node.names:
                        module_name = name.name.split('.')[0]
                        if module_name not in sys.builtin_module_names:
                            imported_modules.add(module_name)
                elif isinstance(node, ast.ImportFrom):
                    module_name = node.module.split('.')[0] if node.module else ''
                    if module_name and module_name not in sys.builtin_module_names:
                        imported_modules.add(module_name)
                        
        print(f"[DEBUG] {file_path}에서 추출된 모듈: {imported_modules}")
        return imported_modules
    except FileNotFoundError as e:
        print(f"[DEBUG] 파일을 찾을 수 없음: {e}")
        return set()
    except UnicodeDecodeError as e:
        print(f"[DEBUG] 파일 디코딩 오류: {e}")
        return set()
    except SyntaxError as e:
        print(f"[DEBUG] 구문 오류: {e}")
        return set()
    except Exception as e:
        print(f"[DEBUG] 기타 오류: {e}")
        return set()


def run_file(file_paths_with_passwords):
    if not isinstance(file_paths_with_passwords, list):
        file_paths_with_passwords = [file_paths_with_passwords]

    # Tcl/Tk는 스레드 안전하지 않으므로, 워커 스레드에서의 UI 작업은 모두
    # 메인 스레드(root.after)로 넘겨서 처리한다.
    def ui_call(func, *a, **kw):
        root.after(0, lambda: func(*a, **kw))

    def show_error(title, message):
        ui_call(messagebox.showerror, title, message)

    def install_python(callback):
        """Python을 자동으로 다운로드하고 설치하는 함수"""
        install_window = None
        status_label = None

        def create_window():
            nonlocal install_window, status_label
            install_window = tk.Toplevel(root)
            install_window.title("Python 설치 중")
            install_window.geometry("300x150")
            install_window.attributes('-topmost', True)
            install_window.configure(bg="white")
            status_label = tk.Label(install_window, text="Python을 다운로드 중입니다...", bg="white")
            status_label.pack(pady=20)

        def set_status(text):
            if status_label is not None and status_label.winfo_exists():
                status_label.config(text=text)

        def close_window(delay=0):
            if install_window is not None and install_window.winfo_exists():
                if delay:
                    install_window.after(delay, install_window.destroy)
                else:
                    install_window.destroy()

        try:
            ui_call(create_window)

            python_url = "https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"
            installer_path = os.path.join(os.getenv("TEMP"), "python_installer.exe")

            response = requests.get(python_url, stream=True)
            response.raise_for_status()
            with open(installer_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            ui_call(set_status, "다운로드 완료, 설치를 시작합니다...")

            install_command = [installer_path, "/quiet", "InstallAllUsers=1", "PrependPath=1"]
            subprocess.run(install_command, check=True)

            ui_call(set_status, "Python 설치 완료!")
            ui_call(close_window, 2000)

            python_executable = shutil.which("python") or shutil.which("python3")
            if python_executable:
                callback(python_executable)
            else:
                show_error("오류", "Python 설치 후에도 실행 파일을 찾을 수 없습니다.")
        except Exception as e:
            show_error("오류", f"Python 설치 중 오류 발생: {e}")
            ui_call(close_window)

    def _execute_file(file_path, python_executable=None):
        try:
            file_path = auto_change_slash(file_path)
            if file_path.endswith('.py'):
                if not python_executable:
                    python_executable = shutil.which("python") or shutil.which("python3")
                    if not python_executable and sys.executable.endswith('.exe') and not sys.executable.lower().endswith('python.exe'):
                        python_executable = None

                if not python_executable:
                    def ask_and_install():
                        if messagebox.askyesno("Python 필요", "Python이 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            def after_install(executable):
                                _execute_file(file_path, executable)
                            threading.Thread(target=install_python, args=(after_install,), daemon=True).start()
                    ui_call(ask_and_install)
                    return
                else:
                    # 예전에는 여기에 "--bypass-password" 를 붙였다.
                    # 받는 쪽 구현이 어디에도 없어서, 이 인자를 모르는 프로그램은
                    # (argparse 를 쓰는 경우 특히) 그냥 오류로 종료됐다.
                    # 창을 숨긴 채 실행하므로 사용자에게는 "아무 반응 없음" 으로 보였다.
                    cmd = [python_executable, file_path]
                    # CMD 창 항상 숨김
                    process = subprocess.Popen(
                        cmd,
                        shell=False,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    print(f"[DEBUG] .py 실행 (PID={process.pid}): {' '.join(cmd)}")
            elif file_path.endswith('.exe'):
                cmd = [file_path]
                process = subprocess.Popen(
                    cmd,
                    shell=False,
                    creationflags=subprocess.CREATE_NO_WINDOW
                )
                print(f"[DEBUG] .exe 실행 (PID={process.pid}): {' '.join(cmd)}")
            elif file_path.startswith("http://") or file_path.startswith("https://"):
                webbrowser.open(file_path)
            elif file_path.startswith("\\\\"):  # NAS 경로
                if os.path.exists(file_path):
                    subprocess.Popen(['explorer', file_path], shell=False)
                else:
                    show_error("오류", f"NAS 경로를 찾을 수 없습니다: {file_path}")
            else:
                base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
                full_file_path = os.path.join(base_path, file_path)
                if os.path.exists(full_file_path):
                    subprocess.Popen(['explorer', full_file_path], shell=False)
                else:
                    show_error("오류", f"파일을 찾을 수 없습니다: {full_file_path}")
        except Exception as e:
            show_error("오류", f"파일 실행 중 오류 발생: {e}")

    url_paths = []
    non_url_paths = []

    for path_with_pw in file_paths_with_passwords:
        if "|" in path_with_pw:
            path, password = path_with_pw.split("|", 1)
        else:
            path, password = path_with_pw, ""
        if path.startswith("http://") or path.startswith("https://"):
            url_paths.append((path, password))
        else:
            non_url_paths.append((path, password))

    if url_paths:
        def open_urls(choice):
            chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe"
            if not os.path.exists(chrome_path):
                chrome_path = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe"
            if not os.path.exists(chrome_path):
                messagebox.showerror("오류", "Chrome 브라우저를 찾을 수 없습니다.")
                return
            for url, pw in url_paths:
                if pw:
                    entered_pw = simpledialog.askstring("비밀번호 입력", f"{url}에 대한 비밀번호:", show="*")
                    if entered_pw != pw:
                        messagebox.showerror("오류", f"{url}에 대한 비밀번호가 일치하지 않습니다.")
                        continue
                cmd = [chrome_path, "--incognito", url] if choice == 'yes' else [chrome_path, url]
                subprocess.Popen(cmd, shell=False)
            window_url.destroy()

        window_url = tk.Toplevel(root)
        window_url.title("URL 실행")
        window_url.geometry("300x70")
        window_url.attributes('-topmost', True)
        tk.Label(window_url, text=f"선택된 {len(url_paths)}개의 URL을 시크릿 모드에서 열겠습니까?").pack()
        button_frame_url = tk.Frame(window_url)
        button_frame_url.pack()
        tk.Button(button_frame_url, text="예", command=lambda: open_urls('yes'), padx=15, pady=5).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame_url, text="아니오", command=lambda: open_urls('no'), padx=15, pady=5).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame_url, text="취소", command=window_url.destroy, padx=15, pady=5).pack(side=tk.LEFT, padx=10)

    # 비-URL 경로를 병렬로 실행
    threads = []
    for path, pw in non_url_paths:
        if pw:
            entered_pw = simpledialog.askstring("비밀번호 입력", "비밀번호:", show="*")
            if entered_pw != pw:
                messagebox.showerror("오류", f"{path}에 대한 비밀번호가 일치하지 않습니다.")
                continue
        thread = threading.Thread(target=_execute_file, args=(path,))
        threads.append(thread)
        thread.start()
        print(f"[DEBUG] 스레드 시작: {path}")
        
        
        
        
        
        
    
   



        

def adjust_font_size(button, text, max_width):
    font_name = button.cget("font").split(" ")[0]
    font_size = button.font_size
    font_weight = button.font_weight
    current_font = (font_name, font_size, font_weight)
    button.config(font=current_font)
    # 버튼 크기(width, height)는 변경하지 않고 글자 크기만 조정
    while button.winfo_reqwidth() > max_width and font_size > 8:
        font_size -= 1
        button.config(font=(font_name, font_size, font_weight))
    button.font_size = font_size




def create_password_entry_add(parent):
    password_label_add = tk.Label(parent, text="비밀번호:", bg="white")
    password_label_add.grid(row=3, column=0, sticky="w")
    password_entry_add = tk.Entry(parent, show="*")
    password_entry_add.grid(row=3, column=1)
    return password_entry_add

def add_entry():
    try:
        name = name_entry.get()
        path_with_password = path_entry.get()
    except NameError:
        return

    path_with_password = auto_change_slash(path_with_password)
    if "|" in path_with_password:
        path, password = path_with_password.split("|", 1)
    else:
        path = path_with_password
        password = ""

    # .py 또는 .exe 파일의 경우 비밀번호를 12345로 설정
    if path.endswith(".py") or path.endswith(".exe"):
        password = "12345"

    # 기존 버튼 이름 목록 가져오기
    existing_buttons = [button.cget("text").replace(" 📁", "").replace(" ↲", "") for button in content_frame.winfo_children() if isinstance(button, tk.Button)]
    
    # 이름 중복 처리
    original_name = name
    counter = 1
    while name in existing_buttons:
        name = f"{original_name}-{counter}"
        counter += 1

    color = "#4CAF50"
    button = MyButton(
        content_frame, 
        path, 
        password=password, 
        text=name, 
        command=partial(run_file, path + ("|12345" if path.endswith(".py") or path.endswith(".exe") else ("|" + password if password else ""))),
        width=15, 
        height=1, 
        bg=color, 
        fg="white"
    )
    button.place(x=50, y=50)
    button.path = path
    button.password = password
    adjust_font_size(button, name, max_width=150)
    # MyButton.__init__에서 이벤트 바인딩 자동 처리됨

    new_button_data = {
        "name": name,
        "path": path,
        "x": 50,
        "y": 50,
        "color": color,
        "password": password,
        "font_size": 10,
        "fg": "white",
        "font_weight": "normal"
    }
    update_json_file(name, json_directory, new_button_data)

    num_buttons = len([child for child in content_frame.winfo_children() if isinstance(child, tk.Button)])
    if 1 <= num_buttons <= 9:
        shortcut_key = f"<Control-{num_buttons}>"
    elif 10 <= num_buttons <= 35:
        shortcut_key = f"<Control-{chr(num_buttons + 87)}>"
    else:
        return

    try:
        root.bind(shortcut_key, partial(run_file, path + ("|12345" if path.endswith(".py") or path.endswith(".exe") else ("|" + password if password else ""))))
        print(f"단축키 {shortcut_key}가 {name} 버튼에 할당되었습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"단축키 할당 중 오류 발생: {e}")
        
        
        
        



def save_button_state(name, path, x, y, color, password, font_size, fg, font_weight, width, height, selected):
    clean_name = name.replace(" 📁", "").replace(" ↲", "")
    sanitized_name = sanitize_filename(clean_name)
    button_data = {
        "name": clean_name,
        "path": path,
        "x": x,
        "y": y,
        "color": color,
        "password": password,
        "font_size": font_size,
        "fg": fg,
        "font_weight": font_weight,
        "width": width,
        "height": height,
        "original_width": width,  # original_width 추가
        "original_height": height,  # original_height 추가
        "is_folder": os.path.isdir(path),
        "selected": selected
    }
    file_name = os.path.join(json_directory, f"{sanitized_name}_button.json")
    try:
        with open(file_name, "w", encoding="utf-8") as file:
            json.dump(button_data, file, indent=4, ensure_ascii=False)
        print(f"[DEBUG] 버튼 상태 저장됨: {file_name}, 위치: ({x}, {y})")
    except OSError as e:
        print(f"[ERROR] JSON 파일 저장 실패: {file_name} - {e}")
   
        
   
    
        
def normalize_button_files():
    """파일명이 정규형이 아닌 예전 JSON 을 정리한다.

    폴더 버튼 이름에는 " 📁" 가 붙는데, 저장할 때 이걸 지우지 않은 코드가 있어서
    같은 버튼이 "이름 📁_button.json" 과 "이름_button.json" 두 파일로 남았다.
    그러면 화면에 버튼이 두 개 생기고, 하나를 지워도 다른 파일이 남아 되살아난다.
    여기서 정규 파일명으로 합쳐 준다(내용이 다르면 최근 것을 남긴다).
    """
    try:
        names = [f for f in os.listdir(json_directory) if f.endswith("_button.json")]
    except OSError as e:
        print(f"[DEBUG] 버튼 파일 목록 읽기 실패: {e}")
        return

    suffix = "_button.json"
    for file_name in names:
        key = file_name[:-len(suffix)]
        canonical = sanitize_filename(key)
        if canonical == key or not canonical:
            continue
        src = os.path.join(json_directory, file_name)
        dst = os.path.join(json_directory, f"{canonical}{suffix}")
        try:
            if os.path.exists(dst):
                # 둘 다 있으면 최근에 저장된 쪽을 남긴다
                if os.path.getmtime(src) > os.path.getmtime(dst):
                    os.replace(src, dst)
                    print(f"[DEBUG] 중복 버튼 파일 병합(최신 유지): {file_name} -> {canonical}{suffix}")
                else:
                    os.remove(src)
                    print(f"[DEBUG] 중복 버튼 파일 제거: {file_name}")
            else:
                os.replace(src, dst)
                print(f"[DEBUG] 버튼 파일명 정규화: {file_name} -> {canonical}{suffix}")
        except OSError as e:
            print(f"[ERROR] 버튼 파일 정리 실패: {file_name} - {e}")


def load_button_states():
    normalize_button_files()   # 예전 중복 파일 정리 후 로드
    MyButton._created_buttons.clear()
    for btn in content_frame.winfo_children():
        if isinstance(btn, MyButton):
            btn.destroy()
    print("[DEBUG] _created_buttons 및 UI 초기화됨")

    button_files = [f for f in os.listdir(json_directory) if f.endswith("_button.json")]
    button_data_list = []

    for file_name in button_files:
        file_path = os.path.join(json_directory, file_name)
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                button_data = json.load(file)
            button_data["width"] = button_data.get("width", 15)
            button_data["height"] = button_data.get("height", 2)
            button_data["original_width"] = button_data.get("original_width", button_data["width"])
            button_data["original_height"] = button_data.get("original_height", button_data["height"])
            button_data["fg"] = button_data.get("fg", "white")
            button_data["font_weight"] = button_data.get("font_weight", "normal")
            button_data["x"] = button_data.get("x", 0)
            button_data["y"] = button_data.get("y", 0)
            button_data["font_size"] = button_data.get("font_size", 10)
            button_data["color"] = button_data.get("color", "#FFFFFF")
            button_data["path"] = button_data.get("path", "")
            button_data["name"] = button_data.get("name", file_name.replace("_button.json", ""))
            button_data["is_folder"] = button_data.get("is_folder", os.path.isdir(button_data["path"]))
            # JSON 파일명에서 json_key 복원 (삭제 시 정확한 파일 특정에 사용)
            button_data["_json_key"] = file_name.replace("_button.json", "")
            button_data_list.append(button_data)
            print(f"[DEBUG] 버튼 데이터 로드: {button_data['name']}, 위치: ({button_data['x']}, {button_data['y']})")
        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON 파싱 오류: {file_path} - {e}")

    for button_data in button_data_list:
        display_text = f"{button_data['name']} 📁" if button_data["is_folder"] else button_data["name"]
        button = MyButton(
            content_frame,
            path=button_data["path"],
            password=button_data.get("password", ""),
            text=display_text,
            command=partial(run_file, button_data["path"] + ("|" + button_data["password"] if button_data["password"] else "")),
            bg=button_data["color"],
            fg=button_data["fg"],
            font=("Arial", button_data["font_size"], button_data["font_weight"]),
            width=button_data["width"],
            height=button_data["height"],
            is_folder=button_data["is_folder"],
            is_copy=False
        )
        button.place(x=button_data["x"], y=button_data["y"])
        button.original_bg = button_data["color"]
        button.original_fg = button_data["fg"]
        button.original_width = button_data["original_width"]
        button.original_height = button_data["original_height"]
        button.hotkey = button_data.get("hotkey", None)
        # json_key를 button.name으로 복원 → 삭제 시 정확한 JSON 파일 제거 가능
        json_key = button_data["_json_key"]
        button.name = json_key
        # _created_buttons도 json_key 기준으로 재등록
        MyButton._created_buttons[json_key] = button
        adjust_font_size(button, display_text, max_width=150)
        print(f"[DEBUG] 버튼 생성 완료: {display_text}, json_key={json_key}, 위치: ({button_data['x']}, {button_data['y']})")

    print(f"[DEBUG] 총 로드된 버튼 수: {len(button_data_list)}")
    return button_data_list


def load_label_states():
    """라벨을 JSON 에서 다시 만든다.

    load_button_states() 와 마찬가지로 "기존 것을 모두 지우고 새로 만드는" 방식이라
    몇 번을 호출해도 결과가 같다(멱등). 호출한 쪽에서 반환값으로 다시 만들면
    라벨이 겹쳐서 생기므로, 라벨 생성은 반드시 이 함수만 담당한다.
    """
    for widget in content_frame.winfo_children():
        if isinstance(widget, (MyLabel, tk.Label)):
            try:
                widget.destroy()
            except Exception as e:
                print(f"[WARNING] 기존 라벨 제거 중 오류: {e}")

    label_data_list = load_labels()
    for label_data in label_data_list:
        try:
            # MyLabel.__init__ 에서 이벤트 바인딩이 자동으로 이루어진다
            label = MyLabel(
                content_frame,
                text=label_data["text"],
                font=(label_data["font_name"], label_data["font_size"], label_data["font_style"]),
                bg="white",
                fg=label_data.get("fg", label_data.get("color", "black")),
                uid=label_data.get("uid")
            )
            label.place(x=label_data["x"], y=label_data["y"])
            label.executed = False
        except Exception as e:
            print(f"[ERROR] 라벨 생성 중 오류 - 텍스트: {label_data.get('text')}, 오류: {e}")

    print(f"[DEBUG] 총 로드된 라벨 수: {len(label_data_list)}")
    return label_data_list




def save_label_state(label_name, x, y, color="black", font_name="Arial", font_size=12, font_style="normal", uid=None):
    """라벨 상태를 JSON 으로 저장한다.

    uid 가 주어지면 파일명을 uid 로 만든다. 이때문에 텍스트가 같은 라벨을
    여러 개 만들어도 서로의 파일을 덮어쓰지 않고, 이름을 바꿔도 같은 파일에 저장된다.
    uid 가 없으면 예전 방식(텍스트 기반 파일명)으로 동작한다.
    """
    label_data = {
        "text": label_name,
        "x": x,
        "y": y,
        "color": color,
        "fg": color,
        "font_name": font_name,
        "font_size": font_size,
        "font_style": font_style
    }
    if not os.path.exists(json_directory):
        os.makedirs(json_directory)
    if uid:
        label_data["uid"] = uid
        file_name = f"{json_directory}/{uid}_label.json"
    else:
        file_name = legacy_label_json_path(label_name)
    with open(file_name, "w", encoding="utf-8") as file:
        json.dump(label_data, file, indent=4)


class MyLabel(tk.Label):
    def __init__(self, master, *args, **kwargs):
        # 라벨 고유 ID(JSON 파일명으로 사용). 없으면 새로 부여한다.
        uid = kwargs.pop("uid", None)

        # font 인수를 kwargs에서 추출하고 제거
        font_tuple = kwargs.pop("font", ("Arial", 12, "normal"))
        if isinstance(font_tuple, tuple) and len(font_tuple) >= 2:
            font_name = font_tuple[0]
            font_size = font_tuple[1]
            font_style = font_tuple[2] if len(font_tuple) > 2 else "normal"
        else:
            font_name = "Arial"
            font_size = 12
            font_style = "normal"

        # 부모 클래스 호출 시 font를 명시적으로 전달
        super().__init__(master, *args, font=(font_name, font_size, font_style), **kwargs)
        
        self.uid = uid or _new_label_uid()  # 같은 이름의 라벨을 구분하는 키
        self.selected = False  # 선택 상태 초기화
        self.relative_x = 0    # 다중 드래그를 위한 상대적 위치
        self.relative_y = 0
        self.alarm_shown = False  # on_right_click_label에서 사용
        self.font_size = font_size  # 폰트 크기 저장
        self.font_style = font_style  # 폰트 스타일 저장
        
        # 이벤트 바인딩
        self.bind("<ButtonPress-1>", partial(on_label_press, self))
        self.bind("<B1-Motion>", partial(on_label_drag, self))
        self.bind("<ButtonRelease-1>", partial(on_label_release, self))
        self.bind("<Button-3>", partial(on_right_click_label, self))  # 우클릭 추가
        
        print(f"[DEBUG] MyLabel 초기화 완료 - 텍스트: {self.cget('text')}, font: {self.cget('font')}")
        
        

def load_labels():
    label_files = [f for f in os.listdir(json_directory) if f.endswith("_label.json")]
    label_data_list = []
    valid_styles = ["normal", "bold", "italic"]  # 허용된 스타일 목록
    
    print(f"[DEBUG] 라벨 파일 검색 완료 - 발견된 파일 수: {len(label_files)}")
    
    for file_name in label_files:
        file_path = os.path.join(json_directory, file_name)
        try:
            # 이관(파일 삭제)을 위해 읽기가 끝나면 바로 닫는다.
            with open(file_path, "r", encoding="utf-8") as file:
                label_data = json.load(file)

            # 기본값 설정
            label_data["fg"] = label_data.get("fg", label_data.get("color", "black"))
            label_data["font_name"] = label_data.get("font_name", "Arial")
            label_data["font_size"] = label_data.get("font_size", 12)
            label_data["font_style"] = label_data.get("font_style", "normal")
            label_data["x"] = label_data.get("x", 0)
            label_data["y"] = label_data.get("y", 0)

            # 필수 키 확인
            required_keys = ["text", "x", "y", "font_name", "font_size", "font_style"]
            if not all(key in label_data for key in required_keys):
                print(f"[WARNING] {file_name}에 필수 키 누락, 기본값 사용")

            # font_style 보정
            if label_data["font_style"] not in valid_styles:
                print(f"[WARNING] {file_name}의 font_style '{label_data['font_style']}'이(가) 유효하지 않음. 'normal'로 설정")
                label_data["font_style"] = "normal"

            # uid 부여 + 레거시(텍스트 기반 파일명) 파일 이관
            uid = label_data.get("uid")
            if not uid:
                uid = _new_label_uid()
                label_data["uid"] = uid
            expected_name = f"{uid}_label.json"
            if file_name != expected_name:
                try:
                    with open(os.path.join(json_directory, expected_name), "w", encoding="utf-8") as new_file:
                        json.dump(label_data, new_file, indent=4)
                    os.remove(file_path)
                    print(f"[DEBUG] 라벨 JSON 이관: {file_name} -> {expected_name}")
                except Exception as e:
                    print(f"[ERROR] 라벨 JSON 이관 실패: {file_name} - {e}")

            # 텍스트가 같은 라벨도 각각 유지한다(uid 로 구분)
            label_data_list.append(label_data)
            print(f"[DEBUG] 라벨 데이터 추가: {label_data['text']} (uid={uid})")

        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON 파싱 오류: {file_path} - {e}")
        except Exception as e:
            print(f"[ERROR] 라벨 파일 로드 중 오류: {file_path} - {e}")

    return label_data_list








from datetime import timedelta  # datetime 은 위에서 이미 import

class ModernCalendarApp:
    KOREA_ICAL_URL = "https://calendar.google.com/calendar/ical/ko.south_korea%23holiday%40group.v.calendar.google.com/public/basic.ics"
    CONFIG_FILE = "C:/json/calendar_config.json"

    def __init__(self, root, schedule=None):
        self.root = root
        self.root.title("Calender")
        self.root.geometry("900x700")
        
        # C:/json 디렉토리 생성
        json_dir = os.path.dirname(self.CONFIG_FILE)
        if not os.path.exists(json_dir):
            try:
                os.makedirs(json_dir)
                print(f"[DEBUG] C:/json 디렉토리 생성됨")
            except PermissionError as e:
                messagebox.showerror("오류", f"C:/json 디렉토리 생성 권한 오류: {e}")
                print(f"[ERROR] C:/json 디렉토리 생성 실패 - 권한 오류: {e}")
            except Exception as e:
                messagebox.showerror("오류", f"C:/json 디렉토리 생성 실패: {e}")
                print(f"[ERROR] C:/json 디렉토리 생성 실패: {e}")

        try:
            self.root.style = ttk.Style("flatly")
        except Exception:
            pass

        self.year = datetime.now().year
        self.month = datetime.now().month
        self.schedule = schedule if schedule is not None else {}
        self.holidays = self.fetch_holidays()
        self.custom_colors = self.load_custom_colors()

        self.style = ttk.Style()
        self.configure_styles()
        self.create_widgets()
        self.cells = {}
        self.draw_calendar()

        self.root.bind("<Configure>", self.on_resize)

    def fetch_holidays(self):
        try:
            response = requests.get(self.KOREA_ICAL_URL)
            response.raise_for_status()
            cal = Calendar.from_ical(response.text)
            holidays = {}
            non_holidays = ["-04-05", "-07-17", "-10-01", "-12-24", "-12-31"]
            for event in cal.walk('VEVENT'):
                dt = event.get('dtstart').dt
                if isinstance(dt, datetime):
                    dt = dt + timedelta(hours=9)
                    dt = dt.date()
                holiday_date = dt.strftime("%Y-%m-%d")
                if any(holiday_date.endswith(non_holiday) for non_holiday in non_holidays):
                    print(f"[DEBUG] 비공휴일 제외: {holiday_date} - {event.get('summary')}")
                    continue
                holidays[holiday_date] = event.get('summary')
                print(f"[DEBUG] 공휴일 추가: {holiday_date} - {event.get('summary')}")
            print(f"[DEBUG] 공휴일 로드 성공: {len(holidays)}개")
            return holidays
        except Exception as e:
            print(f"[DEBUG] 공휴일 로드 실패: {e}")
            return {}

    def configure_styles(self, cell_width=128, cell_height=100, schedule_count=1):
        max_font_size = max(6, min(14, int(cell_width / 15)))
        min_font_size = 4
        if schedule_count <= 1:
            base_font_size = max_font_size
        else:
            base_font_size = max(min_font_size, int(max_font_size - (schedule_count - 1) * 1.5))
        small_font_size = max(4, min(10, int(cell_width / 22)))

        self.style.configure("Header.TLabel", font=("Helvetica", 16, "bold"))
        self.style.configure("Day.TLabel", font=("Helvetica", 10, "bold"))
        self.style.configure("HolidayName.TLabel", font=("Helvetica", 9), foreground="red")
        self.style.configure("Schedule.TLabel", font=("Helvetica", base_font_size, "bold"),
                            foreground="#006400", wraplength=cell_width - 10)
        self.style.configure("SmallSchedule.TLabel", font=("Helvetica", small_font_size, "bold"),
                            foreground="#006400", wraplength=cell_width - 10)
        self.style.configure("Nav.TButton", font=("Helvetica", 12))
        self.style.configure("Cell.TFrame", background="#f0f0f0")
        self.style.configure("Hover.TFrame", background="#e0e0e0")

        self.base_font_size = base_font_size
        self.small_font_size = small_font_size

    def create_widgets(self):
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(expand=True, fill="both")

        self.header = ttk.Frame(self.main_frame, padding=5, style="Cell.TFrame")
        self.header.pack(fill="x", side="top")

        years = list(range(self.year - 10, self.year + 11))
        self.year_combo = ttk.Combobox(self.header, values=years, width=6, state="readonly")
        self.year_combo.set(self.year)
        self.year_combo.grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.year_combo.bind("<<ComboboxSelected>>", self.update_calendar)

        months = [f"{i}월" for i in range(1, 13)]
        self.month_combo = ttk.Combobox(self.header, values=months, width=6, state="readonly")
        self.month_combo.set(f"{self.month}월")
        self.month_combo.grid(row=0, column=1, padx=(5, 5), sticky="w")
        self.month_combo.bind("<<ComboboxSelected>>", self.update_calendar)

        self.prev_button = ttk.Button(self.header, text="◀", command=self.prev_month, style="Nav.TButton", width=5)
        self.next_button = ttk.Button(self.header, text="▶", command=self.next_month, style="Nav.TButton", width=5)
        self.month_label = ttk.Label(self.header, text="", style="Header.TLabel")

        self.header.columnconfigure(2, weight=1)
        self.header.columnconfigure(4, weight=1)
        self.prev_button.grid(row=0, column=2, padx=(5, 5), sticky="e")
        self.month_label.grid(row=0, column=3, sticky="ew")
        self.next_button.grid(row=0, column=4, padx=(5, 0), sticky="w")

        self.calendar_frame = ttk.Frame(self.main_frame, style="Cell.TFrame")
        self.calendar_frame.pack(expand=True, fill="both", padx=0, pady=0)
        self.calendar_frame.grid_propagate(False)

    def update_calendar(self, event=None):
        selected_year = int(self.year_combo.get())
        selected_month = int(self.month_combo.get().replace("월", ""))
        if selected_year != self.year or selected_month != self.month:
            self.year = selected_year
            self.month = selected_month
            self.draw_calendar()

    def draw_calendar(self):
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        self.cells.clear()

        self.month_label.config(text=f"{self.year}년 {self.month}월")
        print(f"[DEBUG] Drawing calendar for {self.year}-{self.month}, schedule: {self.schedule}")

        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        header_height = self.header.winfo_reqheight()
        calendar_height = root_height - header_height
        calendar_width = root_width
        self.calendar_frame.configure(width=calendar_width, height=calendar_height)

        cell_width = (root_width - 14) // 7
        if not self.schedule or not any(f"{self.year}-{self.month:02d}-{d:02d}" in self.schedule for d in range(1, 32)):
            schedule_count = 0
        else:
            schedule_count = max(len(self.schedule.get(f"{self.year}-{self.month:02d}-{d:02d}", [])) 
                                for d in range(1, 32) if f"{self.year}-{self.month:02d}-{d:02d}" in self.schedule)

        base_height = max(80, int(cell_width * 0.8))
        extra_height_per_item = max(12, int(self.base_font_size * 1.8))
        cell_height_default = base_height if schedule_count <= 5 else base_height + extra_height_per_item * (schedule_count - 5)
        self.configure_styles(cell_width, cell_height_default, schedule_count)

        days = ["일", "월", "화", "수", "목", "금", "토"]
        for col, day in enumerate(days):
            color = "red" if col in [0, 6] else "black"
            lbl = ttk.Label(self.calendar_frame, text=day, anchor="center", width=12,
                            style="Day.TLabel", foreground=color)
            lbl.grid(row=0, column=col, pady=(0, 5))

        cal = calendar.Calendar(firstweekday=6)
        row = 1
        for day, weekday in cal.itermonthdays2(self.year, self.month):
            if day == 0:
                continue

            date_str = f"{self.year}-{self.month:02d}-{day:02d}"
            adjusted_column = (weekday + 1) % 7
            print(f"[DEBUG] Date: {date_str}, Weekday: {weekday}, Adjusted Column: {adjusted_column}, Row: {row}")

            schedule_count = len(self.schedule.get(date_str, []))
            cell_height = base_height if schedule_count <= 5 else base_height + extra_height_per_item * (schedule_count - 5)

            cell_frame = ttk.Frame(self.calendar_frame, borderwidth=1, relief="flat", style="Cell.TFrame",
                                  height=cell_height, width=cell_width)
            cell_frame.grid(row=row, column=adjusted_column, padx=2, pady=2, sticky="nsew")
            cell_frame.grid_propagate(False)

            content_frame = ttk.Frame(cell_frame)
            content_frame.pack(fill="both", expand=True)

            is_holiday = date_str in self.holidays
            is_weekend = adjusted_column in [0, 6]
            day_color = self.custom_colors.get(date_str, "red" if (is_holiday or is_weekend) else "black")

            day_label = ttk.Label(content_frame, text=str(day), anchor="center", width=3,
                                  font=("Helvetica", 10, "bold"), foreground=day_color)
            day_label.place(relx=0.5, rely=0.1, anchor="center")
            if date_str in self.schedule and self.schedule[date_str]:
                self.create_tooltip(day_label, "\n".join(self.schedule[date_str]))
                print(f"[DEBUG] Tooltip added for {date_str} with schedule: {self.schedule[date_str]}")

            if is_holiday:
                holiday_name = self.holidays[date_str]
                holiday_label = ttk.Label(content_frame, text=holiday_name, style="HolidayName.TLabel",
                                         wraplength=cell_width-10, justify="center")
                holiday_label.place(relx=0.5, rely=0.3, anchor="center")

            if date_str in self.schedule and self.schedule[date_str]:
                base_rely = 0.4
                if schedule_count <= 5:
                    step = max(0.12, self.base_font_size / 80)
                    style = "Schedule.TLabel"
                else:
                    step = max(0.08, self.small_font_size / 60)
                    style = "SmallSchedule.TLabel"
                for idx, schedule_text in enumerate(self.schedule[date_str]):
                    font_size = self.base_font_size if style == "Schedule.TLabel" else self.small_font_size
                    schedule_label = ttk.Label(content_frame, text=schedule_text,
                                              font=("Helvetica", font_size, "bold"),
                                              foreground="#006400",
                                              wraplength=cell_width - 10,
                                              justify="left")
                    rely_pos = base_rely + step * idx
                    if rely_pos + step > 1.0:
                        rely_pos = min(rely_pos, 0.95 - step)
                    schedule_label.place(relx=0.1, rely=rely_pos, anchor="w")
                    schedule_label.bind("<Button-3>", lambda e, d=date_str, i=idx: self.show_schedule_context_menu(e, d, i))

            for widget in (cell_frame, day_label):
                widget.bind("<Button-3>", lambda e, d=date_str: self.show_color_menu(e, d))
                widget.bind("<Double-1>", lambda e, d=day: self.add_schedule(d, e))
                widget.bind("<Enter>", lambda e, f=cell_frame: f.configure(style="Hover.TFrame"), add="+")
                widget.bind("<Leave>", lambda e, f=cell_frame: f.configure(style="Cell.TFrame"), add="+")

            self.cells[f"{row}-{adjusted_column}"] = {'frame': cell_frame, 'day': day_label}

            if adjusted_column == 6:
                row += 1

        for i in range(7):
            self.calendar_frame.grid_columnconfigure(i, weight=1)
        for i in range(row + 1):
            self.calendar_frame.grid_rowconfigure(i, weight=1)

    def on_resize(self, event):
        if event.widget == self.root:
            self.draw_calendar()

    def show_color_menu(self, event, date):
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="빨간색으로 변경", command=lambda: self.set_custom_color(date, "red"))
        menu.add_command(label="검은색으로 변경", command=lambda: self.set_custom_color(date, "black"))
        menu.post(event.x_root, event.y_root)
        print(f"[DEBUG] Color menu shown for {date}")

    def show_schedule_context_menu(self, event, date, index):
        if date not in self.schedule or not self.schedule[date]:
            print(f"[DEBUG] No schedule for {date}, context menu skipped")
            return
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label=f"삭제: {self.schedule[date][index]}", command=lambda: self.delete_schedule(date, index))
        menu.post(event.x_root, event.y_root)
        print(f"[DEBUG] Schedule context menu shown for {date}, index {index}")

    def set_custom_color(self, date, color):
        self.custom_colors[date] = color
        self.save_custom_colors()
        self.draw_calendar()
        print(f"[DEBUG] Set custom color for {date}: {color}")

    def load_custom_colors(self):
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except PermissionError as e:
                messagebox.showerror("오류", f"파일 읽기 권한 오류: {self.CONFIG_FILE}")
                print(f"[ERROR] Permission error reading {self.CONFIG_FILE}: {e}")
            except json.JSONDecodeError as e:
                messagebox.showerror("오류", f"파일 형식이 잘못됨: {self.CONFIG_FILE}")
                print(f"[ERROR] JSON decode error in {self.CONFIG_FILE}: {e}")
            except Exception as e:
                messagebox.showerror("오류", f"파일 읽기 중 오류: {self.CONFIG_FILE}")
                print(f"[ERROR] Failed to load custom colors: {e}")
        return {}

    def save_custom_colors(self):
        try:
            json_dir = os.path.dirname(self.CONFIG_FILE)
            if not os.path.exists(json_dir):
                os.makedirs(json_dir)
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.custom_colors, f, indent=4)
            print(f"[DEBUG] Saved custom colors: {self.custom_colors}")
        except PermissionError as e:
            messagebox.showerror("오류", f"파일 저장 권한 오류: {self.CONFIG_FILE}")
            print(f"[ERROR] Permission error saving {self.CONFIG_FILE}: {e}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 저장 중 오류: {self.CONFIG_FILE}")
            print(f"[ERROR] Failed to save custom colors: {e}")

    def delete_schedule(self, date, index):
        if date in self.schedule and index < len(self.schedule[date]):
            del self.schedule[date][index]
            if not self.schedule[date]:
                del self.schedule[date]
            print(f"[DEBUG] Deleted schedule for {date}: {self.schedule.get(date, 'None')}")
            self.draw_calendar()

    def create_tooltip(self, widget, text):
        def show_tooltip(event):
            print(f"[DEBUG] Showing tooltip for {widget} with text: {text}, position: +{widget.winfo_rootx() + 20}+{widget.winfo_rooty() + 20}")
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.attributes('-topmost', True)
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + 20
            self.tooltip.wm_geometry(f"+{x}+{y}")
            label = ttk.Label(self.tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1, padding=5)
            label.pack()
            self.tooltip.update_idletasks()
        def hide_tooltip(event):
            print(f"[DEBUG] Hiding tooltip for {widget}")
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
                del self.tooltip
        widget.bind("<Enter>", show_tooltip, add="+")
        widget.bind("<Leave>", hide_tooltip, add="+")

    def prev_month(self):
        self.month -= 1
        if self.month == 0:
            self.month = 12
            self.year -= 1
        self.year_combo.set(self.year)
        self.month_combo.set(f"{self.month}월")
        self.draw_calendar()

    def next_month(self):
        self.month += 1
        if self.month == 13:
            self.month = 1
            self.year += 1
        self.year_combo.set(self.year)
        self.month_combo.set(f"{self.month}월")
        self.draw_calendar()

    def add_schedule(self, day, event):
        date = f"{self.year}-{self.month:02d}-{day:02d}"
        if date not in self.schedule:
            self.schedule[date] = []

        add_window = tk.Toplevel(self.root)
        add_window.title("일정 추가")
        add_window.geometry("300x150")
        add_window.attributes('-topmost', True)
        add_window.configure(bg="white")

        x = event.x_root
        y = event.y_root
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        if x + 300 > screen_width:
            x = screen_width - 300
        if y + 150 > screen_height:
            y = screen_height - 150
        add_window.geometry(f"+{x}+{y}")

        tk.Label(add_window, text=f"{date}의 일정", bg="white").pack(pady=10)
        schedule_entry = tk.Entry(add_window, width=30)
        schedule_entry.pack(pady=5)
        schedule_entry.focus()

        def save_schedule():
            schedule = schedule_entry.get().strip()
            if schedule:
                self.schedule[date].append(schedule)
                print(f"[DEBUG] Added schedule for {date}: {self.schedule[date]}")
                self.draw_calendar()
                self.fade_in_schedule(date, schedule)
            add_window.destroy()

        tk.Button(add_window, text="추가", command=save_schedule).pack(pady=10)
        schedule_entry.bind("<Return>", lambda e: save_schedule())

    def fade_in_schedule(self, date, schedule_text):
        cal = calendar.Calendar(firstweekday=6)
        for day, weekday in cal.itermonthdays2(self.year, self.month):
            if day == 0:
                continue
            date_str = f"{self.year}-{self.month:02d}-{day:02d}"
            if date_str == date:
                adjusted_column = (weekday + 1) % 7
                row = (weekday + sum(1 for d, w in cal.itermonthdays2(self.year, self.month) 
                                    if d != 0 and d < day and w < weekday)) // 7 + 1
                cell_key = f"{row}-{adjusted_column}"
                if cell_key in self.cells:
                    cell_frame = self.cells[cell_key]['frame']
                    for child in cell_frame.winfo_children():
                        content_frame = child
                        for label in content_frame.winfo_children():
                            if (isinstance(label, ttk.Label) and 
                                label.cget("text") == schedule_text and 
                                label.cget("foreground") == "#006400"):
                                alpha = 0.0
                                size_factor = 0.5
                                original_fg = "#006400"
                                original_size = self.base_font_size
                                label.configure(foreground=f"{original_fg}00", 
                                               font=("Helvetica", int(original_size * size_factor), "bold"))

                                def fade():
                                    nonlocal alpha, size_factor
                                    alpha += 0.1
                                    size_factor += 0.05
                                    if alpha <= 1.0:
                                        hex_alpha = f"{int(alpha * 255):02x}"
                                        label.configure(foreground=f"{original_fg}{hex_alpha}",
                                                       font=("Helvetica", int(original_size * size_factor), "bold"))
                                        print(f"[DEBUG] Fade step: alpha={alpha}, size={int(original_size * size_factor)}")
                                        self.root.after(100, fade)
                                    else:
                                        label.configure(foreground=original_fg, 
                                                       font=("Helvetica", original_size, "bold"))
                                        print(f"[DEBUG] Fade complete for {schedule_text}")

                                fade()
                                break
                    break
                
                


# 전역 변수로 edit_window 선언

def edit_selected_buttons(event=None):
    global grid_on, calendar_window, calendar_app, edit_window, copied_button_data

    def toggle_grid():
        global grid_on
        if 'grid_on' not in globals():
            grid_on = False
        grid_on = not grid_on
        # 캔버스가 제거되었으므로 draw_grid/clear_grid는 더 이상 사용되지 않음

    def load_calendar_schedule():
        schedule_file = os.path.join(json_directory, "calendar_schedule.json")
        if os.path.exists(schedule_file):
            with open(schedule_file, "r") as f:
                schedule = json.load(f)
                if not isinstance(schedule, dict):
                    print("[DEBUG] Invalid schedule format, resetting to empty dict")
                    return {}
                print(f"[DEBUG] Loaded schedule from JSON: {schedule}")
                return schedule
        print("[DEBUG] No schedule file found, returning empty dict")
        return {}

    def save_calendar_schedule(schedule):
        schedule_file = os.path.join(json_directory, "calendar_schedule.json")
        os.makedirs(json_directory, exist_ok=True)
        with open(schedule_file, "w") as f:
            json.dump(schedule, f, indent=4)
        print(f"[DEBUG] Saved schedule to JSON: {schedule}")

    def toggle_calendar(event=None):
        global calendar_window, calendar_app
        print(f"[DEBUG] toggle_calendar called: calendar_window exists = {calendar_window is not None}")

        if calendar_window is None or not calendar_window.winfo_exists():
            root.update_idletasks()
            root_width = root.winfo_width()
            root_height = root.winfo_height()
            root_x = root.winfo_x()
            root_y = root.winfo_y()
            root_geometry = f"{root_width}x{root_height}+{root_x}+{root_y}"
            
            calendar_window = ThemedTk(theme="arc")
            calendar_window.title("모던 달력")
            calendar_window.geometry(root_geometry)
            calendar_window.attributes('-topmost', True)
            calendar_window.update_idletasks()

            print(f"[DEBUG] Initial calendar geometry set to: {root_geometry}")

            style = ttk.Style()
            style.configure("Cell.TFrame", background="#ffffff")
            style.configure("Hover.TFrame", background="#f0f0f0")
            style.configure("Header.TLabel", font=("Helvetica", 16, "bold"))
            style.configure("Day.TLabel", font=("Helvetica", 10, "bold"))
            style.configure("Schedule.TLabel", font=("Helvetica", 9))
            style.configure("Nav.TButton", font=("Helvetica", 12))
            style.configure("Holiday.TLabel", font=("Helvetica", 10, "bold"), foreground="red")
            style.configure("HolidayName.TLabel", font=("Helvetica", 9), foreground="red")

            loaded_schedule = load_calendar_schedule()
            calendar_app = ModernCalendarApp(calendar_window, schedule=loaded_schedule)
            print(f"[DEBUG] Initial calendar_app.schedule: {calendar_app.schedule}")

            calendar_window.update()
            actual_geometry = f"{calendar_window.winfo_width()}x{calendar_window.winfo_height()}+{calendar_window.winfo_x()}+{calendar_window.winfo_y()}"
            print(f"[DEBUG] Final calendar geometry: {actual_geometry}")
            if actual_geometry != root_geometry:
                print(f"[DEBUG] Geometry mismatch detected, forcing reset")
                calendar_window.geometry(root_geometry)

            def on_calendar_close():
                global calendar_window, calendar_app
                try:
                    if calendar_app:
                        save_calendar_schedule(calendar_app.schedule)
                except Exception as e:
                    print(f"[DEBUG] Error saving schedule on close: {e}")
                calendar_window.destroy()
                calendar_window = None
                calendar_app = None
                print("[DEBUG] Calendar closed via window CLOSE")

            calendar_window.protocol("WM_DELETE_WINDOW", on_calendar_close)
            calendar_window.bind("<Control-c>", toggle_calendar)
            calendar_window.bind("<Control-C>", toggle_calendar)

            print(f"[DEBUG] Calendar opened with size: {root_width}x{root_height}, position: +{root_x}+{root_y}")
        else:
            try:
                if calendar_app:
                    save_calendar_schedule(calendar_app.schedule)
                if calendar_window:
                    calendar_window.destroy()
                    print("[DEBUG] Calendar window destroyed")
            except Exception as e:
                print(f"[DEBUG] Error during calendar close: {e}")
            finally:
                calendar_window = None
                calendar_app = None
                print("[DEBUG] Calendar closed via Ctrl+C")

    def copy_selected_buttons():
        global copied_button_data
        copied_button_data = []
        for btn in selected_buttons:
            btn_data = {
                "name": btn.cget("text"),
                "path": btn.path,
                "x": btn.winfo_x(),
                "y": btn.winfo_y(),
                "color": btn.original_bg,
                "password": btn.password,
                "font_size": btn.font_size,
                "fg": btn.cget("fg"),
                "font_weight": btn.font_weight,
                "width": btn.cget("width"),
                "height": btn.cget("height"),
                "hotkey": btn.hotkey
            }
            copied_button_data.append(btn_data)
        os.makedirs(json_directory, exist_ok=True)
        with open(COPY_FILE_PATH, "w") as f:
            json.dump(copied_button_data, f)
        print(f"[DEBUG] 다중 선택 버튼 복사됨: {len(copied_button_data)}개 버튼, 저장됨: {COPY_FILE_PATH}")

    # 이벤트 처리
    if event is not None:
        print(f"Event received: keysym={event.keysym}, state={event.state}")
        if event.keysym in ('g', 'G') and event.state & 0x0004:
            toggle_grid()
            print("Ctrl+G pressed: Toggling grid")
            return "break"
        if event.keysym in ('c', 'C') and event.state & 0x0004:
            toggle_calendar(event)
            print("Ctrl+C pressed: Toggling calendar")
            return "break"

    # 선택된 버튼 또는 라벨이 없는 경우
    if not selected_buttons and not selected_labels:
        messagebox.showinfo("알림", "편집할 버튼 또는 라벨을 선택하세요.")
        return

    # 편집 창 생성
    if edit_window is not None and edit_window.winfo_exists():
        edit_window.destroy()
    edit_window = tk.Toplevel(root)
    edit_window.title("선택된 항목 편집")
    edit_window.geometry("430x450")
    edit_window.configure(bg="white")
    edit_window.attributes('-topmost', True)
    edit_window.resizable(False, False)  # 창 크기 조정 불가

    # 휠 이벤트 차단
    def disable_scroll(event):
        return "break"  # 이벤트 전파 중단

    edit_window.bind("<MouseWheel>", disable_scroll)  # 마우스 휠 비활성화

    # 선택된 버튼 정보 표시
    if selected_buttons:
        tk.Label(edit_window, text=f"선택된 버튼 수: {len(selected_buttons)}", font=("Arial", 12, "bold"), bg="white").pack(pady=5)

        # 색상 선택
        color_frame = tk.Frame(edit_window, bg="white")
        color_frame.pack(pady=5)
        tk.Label(color_frame, text="배경색상:", bg="white").pack(side=tk.LEFT, padx=5)
        bg_color_button = tk.Button(color_frame, text="색상 선택", command=lambda: choose_bg_color(bg_color_var))
        bg_color_button.pack(side=tk.LEFT, padx=5)
        bg_color_var = tk.StringVar(value=selected_buttons[0].cget("bg"))
        tk.Label(color_frame, text="글자색상:", bg="white").pack(side=tk.LEFT, padx=5)
        fg_color_button = tk.Button(color_frame, text="색상 선택", command=lambda: choose_fg_color(fg_color_var))
        fg_color_button.pack(side=tk.LEFT, padx=5)
        fg_color_var = tk.StringVar(value=selected_buttons[0].cget("fg"))

        # 폰트 설정
        font_frame = tk.Frame(edit_window, bg="white")
        font_frame.pack(pady=5)
        tk.Label(font_frame, text="글자 크기:", bg="white").pack(side=tk.LEFT, padx=5)
        font_size_var = tk.IntVar(value=selected_buttons[0].font_size)
        font_size_spinbox = tk.Spinbox(font_frame, from_=8, to=20, textvariable=font_size_var, width=5)
        font_size_spinbox.pack(side=tk.LEFT, padx=5)
        tk.Label(font_frame, text="글자 두께:", bg="white").pack(side=tk.LEFT, padx=5)
        font_weight_var = tk.StringVar(value=selected_buttons[0].font_weight)
        font_weight_combo = ttk.Combobox(font_frame, textvariable=font_weight_var, values=["normal", "bold"], width=10)
        font_weight_combo.pack(side=tk.LEFT, padx=5)

        # 크기 설정
        size_frame = tk.Frame(edit_window, bg="white")
        size_frame.pack(pady=5)
        tk.Label(size_frame, text="버튼 너비:", bg="white").pack(side=tk.LEFT, padx=5)
        width_var = tk.IntVar(value=selected_buttons[0].cget("width"))
        width_spinbox = tk.Spinbox(size_frame, from_=5, to=30, textvariable=width_var, width=5)
        width_spinbox.pack(side=tk.LEFT, padx=5)
        tk.Label(size_frame, text="버튼 높이:", bg="white").pack(side=tk.LEFT, padx=5)
        height_var = tk.IntVar(value=selected_buttons[0].cget("height"))
        height_spinbox = tk.Spinbox(size_frame, from_=1, to=10, textvariable=height_var, width=5)
        height_spinbox.pack(side=tk.LEFT, padx=5)

    # 선택된 라벨 정보 표시
    if selected_labels:
        tk.Label(edit_window, text=f"선택된 라벨 수: {len(selected_labels)}", font=("Arial", 12, "bold"), bg="white").pack(pady=5)
        for lbl in selected_labels:
            lbl_frame = tk.Frame(edit_window, bg="white")
            lbl_frame.pack(fill="x", pady=2)
            tk.Label(lbl_frame, text=lbl.cget("text"), bg="white").pack(side="left", padx=5)

    # 간격 조절 옵션 추가
    spacing_frame = tk.Frame(edit_window, bg="white")
    spacing_frame.pack(pady=10, fill="x")
    tk.Label(spacing_frame, text="간격 조절", font=("Arial", 10, "bold"), bg="white").pack()

    # 가로 간격 적용 함수
    def apply_x_spacing():
        # 버튼 가로 간격 조절
        if selected_buttons and len(selected_buttons) >= 3:
            buttons_sorted_x = sorted(selected_buttons, key=lambda b: b.winfo_x())
            first_btn_x, last_btn_x = buttons_sorted_x[0], buttons_sorted_x[-1]

            first_x, first_y = first_btn_x.winfo_x(), first_btn_x.winfo_y()
            last_x, last_y = last_btn_x.winfo_x(), last_btn_x.winfo_y()

            num_gaps = len(selected_buttons) - 1
            total_x = last_x - first_x  # 가로 전체 거리

            if total_x <= 0:
                print("[DEBUG] 첫 번째와 마지막 버튼의 x 위치가 동일하거나 역순입니다.")
                return

            middle_buttons = buttons_sorted_x[1:-1]
            if middle_buttons:
                x_step = total_x / num_gaps  # 균등 간격 계산
                for i, btn in enumerate(middle_buttons, 1):
                    new_x = first_x + i * x_step  # 첫 번째 버튼에서 시작하여 균등 배치
                    new_y = btn.winfo_y()  # y 위치는 유지
                    btn.place(x=new_x, y=new_y)
                    temp_positions[btn] = (new_x, new_y)
                    print(f"[DEBUG] Button {i+1}: new_x = {new_x}, new_y = {new_y}, x_step = {x_step}")

            first_btn_x.place(x=first_x, y=first_y)
            last_btn_x.place(x=last_x, y=last_y)
            temp_positions[first_btn_x] = (first_x, first_y)
            temp_positions[last_btn_x] = (last_x, last_y)

        # 라벨 가로 간격 조절
        if selected_labels and len(selected_labels) >= 3:
            labels_sorted_x = sorted(selected_labels, key=lambda l: l.winfo_x())
            first_lbl_x, last_lbl_x = labels_sorted_x[0], labels_sorted_x[-1]

            first_x, first_y = first_lbl_x.winfo_x(), first_lbl_x.winfo_y()
            last_x, last_y = last_lbl_x.winfo_x(), last_lbl_x.winfo_y()

            num_gaps = len(selected_labels) - 1
            total_x = last_x - first_x

            if total_x <= 0:
                print("[DEBUG] 첫 번째와 마지막 라벨의 x 위치가 동일하거나 역순입니다.")
                return

            middle_labels = labels_sorted_x[1:-1]
            if middle_labels:
                x_step = total_x / num_gaps
                for i, lbl in enumerate(middle_labels, 1):
                    new_x = first_x + i * x_step
                    new_y = lbl.winfo_y()
                    lbl.place(x=new_x, y=new_y)
                    temp_positions[lbl] = (new_x, new_y)

            first_lbl_x.place(x=first_x, y=first_y)
            last_lbl_x.place(x=last_x, y=last_y)
            temp_positions[first_lbl_x] = (first_x, first_y)
            temp_positions[last_lbl_x] = (last_x, last_y)

    # 세로 간격 적용 함수
    def apply_y_spacing():
        # 버튼 세로 간격 조절
        if selected_buttons and len(selected_buttons) >= 3:
            buttons_sorted_y = sorted(selected_buttons, key=lambda b: b.winfo_y())
            first_btn_y, last_btn_y = buttons_sorted_y[0], buttons_sorted_y[-1]

            first_x, first_y = first_btn_y.winfo_x(), first_btn_y.winfo_y()
            last_x, last_y = last_btn_y.winfo_x(), last_btn_y.winfo_y()

            num_gaps = len(selected_buttons) - 1
            total_y = last_y - first_y  # 세로 전체 거리

            if total_y <= 0:
                print("[DEBUG] 첫 번째와 마지막 버튼의 y 위치가 동일하거나 역순입니다.")
                return

            middle_buttons = buttons_sorted_y[1:-1]
            if middle_buttons:
                y_step = total_y / num_gaps  # 균등 간격 계산
                for i, btn in enumerate(middle_buttons, 1):
                    new_y = first_y + i * y_step  # 첫 번째 버튼에서 시작하여 균등 배치
                    new_x = btn.winfo_x()  # x 위치는 유지
                    btn.place(x=new_x, y=new_y)
                    temp_positions[btn] = (new_x, new_y)
                    print(f"[DEBUG] Button {i+1}: new_x = {new_x}, new_y = {new_y}, y_step = {y_step}")

            first_btn_y.place(x=first_x, y=first_y)
            last_btn_y.place(x=last_x, y=last_y)
            temp_positions[first_btn_y] = (first_x, first_y)
            temp_positions[last_btn_y] = (last_x, last_y)

        # 라벨 세로 간격 조절
        if selected_labels and len(selected_labels) >= 3:
            labels_sorted_y = sorted(selected_labels, key=lambda l: l.winfo_y())
            first_lbl_y, last_lbl_y = labels_sorted_y[0], labels_sorted_y[-1]

            first_x, first_y = first_lbl_y.winfo_x(), first_lbl_y.winfo_y()
            last_x, last_y = last_lbl_y.winfo_x(), last_lbl_y.winfo_y()

            num_gaps = len(selected_labels) - 1
            total_y = last_y - first_y

            if total_y <= 0:
                print("[DEBUG] 첫 번째와 마지막 라벨의 y 위치가 동일하거나 역순입니다.")
                return

            middle_labels = labels_sorted_y[1:-1]
            if middle_labels:
                y_step = total_y / num_gaps
                for i, lbl in enumerate(middle_labels, 1):
                    new_y = first_y + i * y_step
                    new_x = lbl.winfo_x()
                    lbl.place(x=new_x, y=new_y)
                    temp_positions[lbl] = (new_x, new_y)

            first_lbl_y.place(x=first_x, y=first_y)
            last_lbl_y.place(x=last_x, y=last_y)
            temp_positions[first_lbl_y] = (first_x, first_y)
            temp_positions[last_lbl_y] = (last_x, last_y)

    # 정렬 기능
    temp_positions = {}  # 간격 조정 및 정렬용 임시 위치 저장
    preview_buttons = {}
    original_positions = {}  # 취소 시 복원을 위한 원래 위치 저장

    def clear_preview():
        nonlocal preview_buttons
        for btn in preview_buttons.values():
            btn.destroy()
        preview_buttons.clear()

    def draw_preview():
        clear_preview()
        for btn, (x, y) in temp_positions.items():
            preview_btn = tk.Button(
                root,
                text=btn.cget("text"),
                bg=btn.cget("bg"),
                fg=btn.cget("fg"),
                font=btn.cget("font"),
                width=btn.cget("width"),
                height=btn.cget("height"),
                relief="raised",
                state="disabled"
            )
            preview_btn.place(x=x, y=y)
            root.attributes('-alpha', 0.5)
            root.after(100, lambda: root.attributes('-alpha', 1.0))
            preview_buttons[btn] = preview_btn

    def align_top():
        if len(selected_buttons) < 2:
            return
        min_y = min(btn.winfo_y() for btn in selected_buttons)
        for btn in selected_buttons:
            temp_positions[btn] = (btn.winfo_x(), min_y)
        draw_preview()

    def align_bottom():
        if len(selected_buttons) < 2:
            return
        max_y = max(btn.winfo_y() + btn.winfo_height() for btn in selected_buttons)
        for btn in selected_buttons:
            new_y = max_y - btn.winfo_height()
            temp_positions[btn] = (btn.winfo_x(), new_y)
        draw_preview()

    def align_left():
        if len(selected_buttons) < 2:
            return
        min_x = min(btn.winfo_x() for btn in selected_buttons)
        for btn in selected_buttons:
            temp_positions[btn] = (min_x, btn.winfo_y())
        draw_preview()

    def align_right():
        if len(selected_buttons) < 2:
            return
        max_x = max(btn.winfo_x() + btn.winfo_reqwidth() for btn in selected_buttons)
        for btn in selected_buttons:
            new_x = max_x - btn.winfo_reqwidth()
            temp_positions[btn] = (new_x, btn.winfo_y())
        draw_preview()

    # 원래 위치 저장 (초기화 시점)
    if selected_buttons:
        for btn in selected_buttons:
            original_positions[btn] = (btn.winfo_x(), btn.winfo_y())
    if selected_labels:
        for lbl in selected_labels:
            original_positions[lbl] = (lbl.winfo_x(), lbl.winfo_y())

    if selected_buttons:
        tk.Label(edit_window, text="정렬:", bg="white").pack(pady=5)
        align_frame = tk.Frame(edit_window, bg="white")
        align_frame.pack(pady=5)
        tk.Button(align_frame, text="상", command=align_top).pack(side=tk.LEFT, padx=5)
        tk.Button(align_frame, text="하", command=align_bottom).pack(side=tk.LEFT, padx=5)
        tk.Button(align_frame, text="좌", command=align_left).pack(side=tk.LEFT, padx=5)
        tk.Button(align_frame, text="우", command=align_right).pack(side=tk.LEFT, padx=5)

    def choose_bg_color(var):
        color = colorchooser.askcolor()[1]
        if color:
            var.set(color)

    def choose_fg_color(var):
        color = colorchooser.askcolor()[1]
        if color:
            var.set(color)

    def apply_changes():
        global edit_window
        try:
            if selected_buttons:
                original_bgs = {button: button.original_bg for button in selected_buttons}
                initial_bg = selected_buttons[0].cget("bg")
                new_bg = bg_color_var.get()
                new_fg = fg_color_var.get()
                bg_changed = (new_bg != initial_bg)
                new_font_size = font_size_var.get()
                new_font_weight = font_weight_var.get()

                for button in selected_buttons[:]:
                    if new_font_size != button.font_size or new_font_weight != button.font_weight:
                        new_font = ("Arial", new_font_size, new_font_weight)
                        button.config(font=new_font)
                        button.font_size = new_font_size
                        button.font_weight = new_font_weight
                    else:
                        new_font = ("Arial", button.font_size, button.font_weight)
                        button.config(font=new_font)

                    new_width = width_var.get()
                    new_height = height_var.get()
                    if new_width != button.cget("width") or new_height != button.cget("height"):
                        button.config(width=new_width, height=new_height)

                    if bg_changed:
                        button.config(bg=new_bg)
                        button.original_bg = new_bg
                    else:
                        button.config(bg=original_bgs[button])
                        button.original_bg = original_bgs[button]

                    button.config(fg=new_fg)

                    # 간격 조정 반영 (temp_positions에 저장된 위치 적용)
                    if button in temp_positions:
                        new_x, new_y = temp_positions[button]
                        button.place(x=new_x, y=new_y)

                    if button.selected:
                        button.selected = False
                        button.config(relief="raised", bg=button.original_bg)
                        if button in selected_buttons:
                            selected_buttons.remove(button)

                    button.update_idletasks()

                    update_json_file(button.cget("text"), json_directory, {
                        "name": button.cget("text"),
                        "path": button.path,
                        "x": button.winfo_x(),
                        "y": button.winfo_y(),
                        "color": button.original_bg,
                        "password": button.password,
                        "font_size": button.font_size,
                        "fg": new_fg,
                        "font_weight": button.font_weight,
                        "width": new_width,
                        "height": new_height,
                        "selected": False,
                        "hotkey": button.hotkey
                    })

            if selected_labels:
                for lbl in selected_labels:
                    if lbl in temp_positions:
                        new_x, new_y = temp_positions[lbl]
                        lbl.place(x=new_x, y=new_y)
                        save_label_state(
                            lbl.cget("text"), new_x, new_y, lbl.cget("fg"),
                            "Arial", int(lbl.cget("font").split()[1]),
                            "bold" if "bold" in lbl.cget("font") else "normal",
                            uid=label_uid(lbl)
                        )

            selected_buttons.clear()
            selected_labels.clear()
            clear_preview()
            edit_window.destroy()
            edit_window = None
        except Exception as e:
            print(f"[DEBUG] 변경 적용 중 오류: {e}")

    def on_cancel():
        global edit_window
        # 취소 시 원래 위치로 복원
        if selected_buttons:
            for btn in selected_buttons:
                if btn in original_positions:
                    orig_x, orig_y = original_positions[btn]
                    btn.place(x=orig_x, y=orig_y)
        if selected_labels:
            for lbl in selected_labels:
                if lbl in original_positions:
                    orig_x, orig_y = original_positions[lbl]
                    lbl.place(x=orig_x, y=orig_y)
        clear_preview()
        edit_window.destroy()
        edit_window = None

    def execute_selected_buttons():
        paths_with_passwords = [
            f"{btn.path}|{btn.password}" if btn.password else btn.path
            for btn in selected_buttons
        ]
        run_file(paths_with_passwords)

    # 버튼 프레임 (가로/세로 간격 조절 버튼 분리)
    button_frame = tk.Frame(edit_window, bg="white")
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="적용", command=apply_changes, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="취소", command=on_cancel, bg="#bdc3c7", fg="black").pack(side=tk.LEFT, padx=5)
    if selected_buttons:
        tk.Button(button_frame, text="동시실행", command=execute_selected_buttons, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="실행버튼 복사", command=copy_selected_buttons, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="가로간격", command=apply_x_spacing, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="세로간격", command=apply_y_spacing, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)

    def on_edit_window_close():
        global edit_window
        on_cancel()  # 창 닫기 시 취소 동작 수행
        edit_window.destroy()
        edit_window = None

    edit_window.protocol("WM_DELETE_WINDOW", on_edit_window_close)
    
    
    
    
    
    
    
      
    
    


    
    
    
    
def toggle_debug_window(event=None):
    global debug_window, debug_text, sys
    
    if debug_window is None or not debug_window.winfo_exists():
        # 디버깅 창 생성
        debug_window = tk.Toplevel(root)
        debug_window.title("Spyder 디버깅 출력")
        debug_window.geometry("600x400")
        debug_window.configure(bg="white")
        debug_window.attributes('-topmost', True)
        
        # 스크롤 가능한 텍스트 위젯 추가
        scrollbar = tk.Scrollbar(debug_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        debug_text = tk.Text(debug_window, wrap="word", yscrollcommand=scrollbar.set, bg="#f0f0f0", fg="black")
        debug_text.pack(fill="both", expand=True, padx=5, pady=5)
        scrollbar.config(command=debug_text.yview)
        
        # 표준 출력 리디렉션 설정
        sys.stdout = StdoutRedirector(debug_text)
        
        # 초기 메시지 삽입
        debug_text.insert(tk.END, "[DEBUG] 디버깅 창이 열렸습니다.\n")
        #debug_text.insert(tk.END, f"[DEBUG] 현재 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        # 창 닫기 시 처리
        def on_debug_close():
            global debug_window, debug_text
            if debug_window and debug_window.winfo_exists():
                sys.stdout = sys.__stdout__  # 원래 stdout 복원
                debug_window.destroy()
                debug_window = None
                debug_text = None
                print("[DEBUG] 디버깅 창 닫힘 - stdout 복원")
        
        debug_window.protocol("WM_DELETE_WINDOW", on_debug_close)
        print("[DEBUG] 디버깅 창 표시됨 - Spyder 출력 캡처 시작")
    else:
        # 디버깅 창 닫기
        sys.stdout = sys.__stdout__  # 원래 stdout 복원
        if debug_window and debug_window.winfo_exists():
            debug_window.destroy()
        debug_window = None
        debug_text = None
        print("[DEBUG] 디버깅 창 닫힘 - stdout 복원")
    
    return "break"  # 이벤트 전파 중단
    







    
    
# 우클릭 시 서식 복사 기능 (복사된 버튼 실행 방지 강화)
def on_right_click_button(button, event):
    global copied_format, copied_button_data

    def copy_format():
        global copied_format
        # 선택 하이라이트 색(#ADD8E6 등)이 아니라 버튼 본래의 서식을 복사해야 한다.
        copied_format = {
            "bg": getattr(button, "original_bg", button.cget("bg")),
            "fg": getattr(button, "original_fg", button.cget("fg")),
            "font_weight": button.font_weight,
            "width": getattr(button, "original_width", button.cget("width")),
            "height": getattr(button, "original_height", button.cget("height"))
        }
        button.config(bg="#FFD700", command=None, relief="raised")
        button.is_copied_format = True
        button.custom_command = None
        print(f"[DEBUG] 서식 복사됨: {button.cget('text')} -> {copied_format}")

    def copy_button():
        global copied_button_data
        button_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")
        copied_button_data = {
            "name": button_name,
            "path": button.path,
            "x": button.winfo_x(),
            "y": button.winfo_y(),
            "color": button.original_bg,
            "password": button.password,
            "font_size": button.font_size,
            "fg": button.cget("fg"),
            "font_weight": button.font_weight,
            "width": button.cget("width"),
            "height": button.cget("height"),
            "hotkey": button.hotkey,
            "is_folder": getattr(button, "is_folder", False)
        }
        os.makedirs(json_directory, exist_ok=True)
        with open(COPY_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(copied_button_data, f)

    def change_path():
        def save_new_path():
            new_path = path_entry.get().strip()
            new_path = auto_change_slash(new_path)
            new_password = password_entry.get().strip()
            if button.password and old_password_entry.get() != button.password:
                messagebox.showerror("오류", "기존 비밀번호가 일치하지 않습니다.")
                return
            if new_path:
                button.path = new_path  # 기존 경로를 새 경로로 대체
                button.refresh_shortcut_mark()  # 바로가기 여부가 바뀔 수 있다
                button.password = new_password if new_password else ""
                button.custom_command = partial(run_file, new_path + ("|" + button.password if button.password else ""))
                save_button_state(
                    button.cget("text"), new_path, button.winfo_x(), button.winfo_y(),
                    button.cget("bg"), button.password, button.font_size, button.cget("fg"),
                    button.font_weight, button.cget("width"), button.cget("height"), button.selected
                )
                change_path_window.destroy()

        change_path_window = tk.Toplevel(root)
        change_path_window.title("경로 변경")
        change_path_window.geometry("300x350")
        change_path_window.attributes('-topmost', True)
        tk.Label(change_path_window, text="새 경로 입력:").pack(pady=5)
        path_entry = tk.Entry(change_path_window)
        path_entry.pack(pady=5)
        path_entry.insert(0, button.path)
        
        def choose_file():
            file_path = filedialog.askopenfilename(filetypes=[("All Files", "*.*")])
            if file_path:
                path_entry.delete(0, tk.END)  # 기존 경로 삭제
                path_entry.insert(0, file_path)  # 새 경로 삽입
                
        def choose_folder():
            folder_path = filedialog.askdirectory()
            if folder_path:
                path_entry.delete(0, tk.END)  # 기존 경로 삭제
                path_entry.insert(0, folder_path)  # 새 경로 삽입

        button_frame = tk.Frame(change_path_window)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text="파일 선택", command=choose_file).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="폴더 선택", command=choose_folder).pack(side=tk.LEFT, padx=5)
        
        if button.password:
            tk.Label(change_path_window, text="기존 비밀번호 입력:").pack(pady=5)
            old_password_entry = tk.Entry(change_path_window, show="*")
            old_password_entry.pack(pady=5)
        tk.Label(change_path_window, text="새 비밀번호 입력:").pack(pady=5)
        password_entry = tk.Entry(change_path_window, show="*")
        password_entry.pack(pady=5)
        tk.Button(change_path_window, text="저장", command=save_new_path).pack(pady=10)

    def change_color():
        color = colorchooser.askcolor()[1]
        if color:
            button.config(bg=color)
            button.original_bg = color
            save_button_state(
                button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                color, button.password, button.font_size, button.cget("fg"),
                button.font_weight, button.cget("width"), button.cget("height"), button.selected
            )

    def change_text_color():
        text_color = colorchooser.askcolor()[1]
        if text_color:
            button.config(fg=text_color)
            button.original_fg = text_color
            save_button_state(
                button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                button.cget("bg"), button.password, button.font_size, text_color,
                button.font_weight, button.cget("width"), button.cget("height"), button.selected
            )

    def toggle_font_weight():
        current_font = button.cget("font")
        font_name, font_size, *font_weight = current_font.split()
        font_weight = "normal" if "bold" in font_weight else "bold"
        button.config(font=(font_name, int(font_size), font_weight))
        button.font_weight = font_weight
        save_button_state(
            button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
            button.cget("bg"), button.password, button.font_size, button.cget("fg"),
            font_weight, button.cget("width"), button.cget("height"), button.selected
        )

    def change_font_size():
        def save_new_font_size():
            try:
                new_size = int(size_entry.get())
                if 8 <= new_size <= 72:
                    font_name = button.cget("font").split(" ")[0]
                    font_weight = button.cget("font").split(" ")[2] if len(button.cget("font").split(" ")) > 2 else "normal"
                    button.config(font=(font_name, new_size, font_weight))
                    button.font_size = new_size
                    save_button_state(
                        button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                        button.cget("bg"), button.password, new_size, button.cget("fg"),
                        font_weight, button.cget("width"), button.cget("height"), button.selected
                    )
                    change_size_window.destroy()
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자를 입력하세요.")
        change_size_window = tk.Toplevel(root)
        change_size_window.title("글자 크기 변경")
        change_size_window.geometry("250x150")
        tk.Label(change_size_window, text="글자 크기 입력 (8-72):").pack(pady=10)
        size_entry = tk.Entry(change_size_window)
        size_entry.pack(pady=5)
        size_entry.insert(0, button.cget("font").split(" ")[1])
        tk.Button(change_size_window, text="저장", command=save_new_font_size).pack(pady=5)

    def rename_button():
        def save_new_name():
            new_name = name_entry.get().strip()
            if new_name:
                old_button_name = button.cget("text")
                try:
                    os.remove(f"{json_directory}/{sanitize_filename(old_button_name)}_button.json")
                except FileNotFoundError:
                    pass
                button.config(text=new_name)
                save_button_state(
                    new_name, button.path, button.winfo_x(), button.winfo_y(),
                    button.cget("bg"), button.password, button.font_size, button.cget("fg"),
                    button.font_weight, button.cget("width"), button.cget("height"), button.selected
                )
                rename_window.destroy()
        rename_window = tk.Toplevel(root)
        rename_window.title("이름 변경")
        rename_window.geometry("300x150")
        rename_window.attributes('-topmost', True)
        tk.Label(rename_window, text="새 이름 입력:").pack(pady=10)
        name_entry = tk.Entry(rename_window)
        name_entry.pack(pady=5)
        name_entry.insert(0, button.cget("text"))
        tk.Button(rename_window, text="저장", command=save_new_name).pack(pady=5)

    def delete_button():
        if messagebox.askyesno("삭제 확인", f"'{button.cget('text')}' 버튼을 정말 삭제하시겠습니까?"):
            button_name = button.cget("text")
            if button in selected_buttons:
                selected_buttons.remove(button)
            button.destroy()
            try:
                os.remove(f"{json_directory}/{sanitize_filename(button_name)}_button.json")
            except FileNotFoundError:
                pass

    def change_password():
        def save_new_password():
            if button.password and old_password_entry.get() != button.password:
                messagebox.showerror("오류", "기존 비밀번호가 일치하지 않습니다.")
                return
            new_password = new_password_entry.get().strip()
            button.password = new_password
            button.custom_command = partial(run_file, button.path + ("|" + new_password if new_password else ""))
            save_button_state(
                button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                button.cget("bg"), button.password, button.font_size, button.cget("fg"),
                button.font_weight, button.cget("width"), button.cget("height"), button.selected
            )
            change_password_window.destroy()
        change_password_window = tk.Toplevel(root)
        change_password_window.title("비밀번호 변경")
        change_password_window.geometry("300x300")
        change_password_window.attributes('-topmost', True)
        tk.Label(change_password_window, text="기존 비밀번호:").pack(pady=5)
        old_password_entry = tk.Entry(change_password_window, show="*")
        old_password_entry.pack(pady=5)
        tk.Label(change_password_window, text="새 비밀번호:").pack(pady=5)
        new_password_entry = tk.Entry(change_password_window, show="*")
        new_password_entry.pack(pady=5)
        tk.Button(change_password_window, text="저장", command=save_new_password).pack(pady=5)

    def set_hotkey():
        def save_hotkey():
            modifier = modifier_var.get()
            key = key_entry.get().lower().strip()
            if not key or not ((key.isdigit() and 0 <= int(key) <= 9) or (key.isalpha() and len(key) == 1)):
                messagebox.showerror("오류", "숫자(0-9) 또는 알파벳(a-z)을 입력하세요.")
                return
            hotkey_str = f"<{modifier}-Key-{key}>"
            if hotkey_str in ["<Control-Key-g>", "<Control-Key-G>", "<Control-Key-c>", "<Control-Key-C>"]:
                messagebox.showerror("오류", f"{hotkey_str}는 시스템 예약 키입니다.")
                return
            for btn in content_frame.winfo_children():
                if isinstance(btn, tk.Button) and btn != button and hasattr(btn, "hotkey") and btn.hotkey == hotkey_str:
                    messagebox.showwarning("중복 핫키", f"'{hotkey_str}'는 이미 '{btn.cget('text')}'에 설정됨.")
                    return
            if hasattr(button, "hotkey") and button.hotkey and button.hotkey != "None":
                root.unbind(button.hotkey)
            root.bind(hotkey_str, lambda e: button.custom_command())
            button.hotkey = hotkey_str
            save_button_state(
                button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                button.cget("bg"), button.password, button.font_size, button.cget("fg"),
                button.font_weight, button.cget("width"), button.cget("height"), button.selected
            )
            hotkey_window.destroy()

        hotkey_window = tk.Toplevel(root)
        hotkey_window.title("핫키 설정")
        hotkey_window.geometry("300x250")
        hotkey_window.attributes('-topmost', True)
        tk.Label(hotkey_window, text="수정 키 선택:").pack(pady=5)
        modifier_var = tk.StringVar(value="Control")
        ttk.Combobox(hotkey_window, textvariable=modifier_var, values=["Control", "Alt"]).pack(pady=5)
        tk.Label(hotkey_window, text="키 입력 (0-9 또는 a-z):").pack(pady=5)
        key_entry = tk.Entry(hotkey_window)
        key_entry.pack(pady=5)
        tk.Button(hotkey_window, text="저장", command=save_hotkey).pack(pady=10)

    def change_button_size():
        def save_new_size():
            try:
                new_width, new_height = int(width_entry.get()), int(height_entry.get())
                if 5 <= new_width <= 30 and 1 <= new_height <= 10:
                    button.config(width=new_width, height=new_height)
                    save_button_state(
                        button.cget("text"), button.path, button.winfo_x(), button.winfo_y(),
                        button.cget("bg"), button.password, button.font_size, button.cget("fg"),
                        button.font_weight, new_width, new_height, button.selected
                    )
                    size_window.destroy()
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자를 입력하세요.")
        size_window = tk.Toplevel(root)
        size_window.title("버튼 크기 변경")
        size_window.geometry("300x200")
        size_window.attributes('-topmost', True)
        tk.Label(size_window, text="버튼 너비 (5-30):").pack(pady=10)
        width_entry = tk.Entry(size_window)
        width_entry.pack(pady=5)
        width_entry.insert(0, str(button.cget("width")))
        tk.Label(size_window, text="버튼 높이 (1-10):").pack(pady=10)
        height_entry = tk.Entry(size_window)
        height_entry.pack(pady=5)
        height_entry.insert(0, str(button.cget("height")))
        tk.Button(size_window, text="저장", command=save_new_size).pack(pady=10)

    def on_paste_button(event=None):
        global copied_button_data
        if not copied_button_data:
            return

        mouse_x = event.x_root - content_frame.winfo_rootx() if event else button.winfo_x()
        mouse_y = event.y_root - content_frame.winfo_rooty() if event else button.winfo_y() + button.winfo_height()

        window_width, window_height = root.winfo_width(), root.winfo_height()

        temp_button = tk.Button(
            root, text="Temp", font=("Arial", copied_button_data.get("font_size", 10), copied_button_data.get("font_weight", "normal"))
        )
        temp_button.update_idletasks()
        button_width, button_height = temp_button.winfo_reqwidth(), temp_button.winfo_reqheight()
        temp_button.destroy()

        new_x = max(5, min(mouse_x, window_width - button_width - 5))
        new_y = max(5, mouse_y)

        existing_buttons = [b for b in content_frame.winfo_children() if isinstance(b, tk.Button)]
        existing_names = [b.cget("text").replace(" 📁", "").replace(" ↲", "") for b in existing_buttons]

        name = copied_button_data.get("name", "새 버튼")
        path = copied_button_data.get("path", "")
        color = copied_button_data.get("color", "#4CAF50")
        fg = copied_button_data.get("fg", "white")
        font_size = copied_button_data.get("font_size", 10)
        font_weight = copied_button_data.get("font_weight", "normal")
        width = copied_button_data.get("width", 15)
        height = copied_button_data.get("height", 1)
        password = copied_button_data.get("password", "")
        is_folder = copied_button_data.get("is_folder", False)

        base_name = name.replace(" 📁", "").replace(" ↲", "")
        new_name = base_name
        counter = 1
        while new_name in existing_names:
            new_name = f"{base_name}_{counter}"
            counter += 1
        display_text = f"{new_name} 📁" if is_folder else new_name

        for btn in existing_buttons:
            bx, by = btn.winfo_x(), btn.winfo_y()
            bw, bh = btn.winfo_width(), btn.winfo_height()
            if abs(bx - new_x) < bw and new_y < by + bh and new_y + button_height > by:
                new_y = by + bh + 5
        new_y = min(new_y, window_height - button_height - 5)

        new_button = MyButton(
            content_frame, text=display_text, path=path, password=password,
            command=partial(run_file, path + ("|" + password if password else "")),
            bg=color, fg=fg, font=("Arial", font_size, font_weight),
            width=width, height=height, is_folder=is_folder, is_copy=False
        )
        new_button.place(x=new_x, y=new_y)
        new_button.update_idletasks()

        new_button.original_bg = color
        new_button.original_fg = fg
        new_button.original_width = width
        new_button.original_height = height
        new_button.font_size = font_size
        new_button.font_weight = font_weight
        # MyButton.__init__에서 이벤트 바인딩 자동 처리됨
        # on_right_click_button은 __init__에서 self에 대해 바인딩되나
        # 복사 버튼은 별도 참조가 필요하므로 우클릭만 재바인딩
        new_button.bind("<Button-3>", lambda e, b=new_button: on_right_click_button(b, e))

        button_data = {
            "name": new_name, "path": path, "x": new_button.winfo_x(), "y": new_button.winfo_y(),
            "color": color, "password": password, "font_size": font_size, "fg": fg,
            "font_weight": font_weight, "width": width, "height": height, "is_folder": is_folder,
            "selected": False
        }
        update_json_file(new_name, json_directory, button_data)

    if button.password:
        password_window = tk.Toplevel(root)
        password_window.title("비밀번호 입력")
        password_window.geometry("300x200")
        password_window.resizable(False, False)
        tk.Label(password_window, text="비밀번호:").pack(pady=15)
        password_entry = tk.Entry(password_window, show="*")
        password_entry.pack(pady=10)

        def check_password():
            if password_entry.get() == button.password:
                password_window.destroy()
                if len(selected_buttons) > 1 and button in selected_buttons:
                    edit_selected_buttons()
                else:
                    menu = tk.Menu(root, tearoff=0)
                    menu.add_command(label="서식 복사", command=copy_format)
                    menu.add_command(label="복사", command=copy_button)
                    menu.add_command(label="붙여넣기", command=lambda: on_paste_button(event))
                    menu.add_separator()
                    menu.add_command(label="이름 변경", command=rename_button)
                    menu.add_command(label="링크 변경", command=change_path)
                    menu.add_command(label="색상 변경", command=change_color)
                    menu.add_command(label="글자색 변경", command=change_text_color)
                    menu.add_command(label="글자 두께 변경", command=toggle_font_weight)
                    menu.add_command(label="글자 크기 변경", command=change_font_size)
                    menu.add_command(label="버튼 크기 변경", command=change_button_size)
                    menu.add_command(label="비밀번호 변경", command=change_password)
                    menu.add_command(label="핫키 설정", command=set_hotkey)
                    menu.add_command(label="시작 버튼 제거", command=delete_button)
                    menu.post(event.x_root, event.y_root)
            else:
                messagebox.showerror("오류", "비밀번호가 일치하지 않습니다.")
        tk.Button(password_window, text="확인", command=check_password).pack(pady=10)
    else:
        if len(selected_buttons) > 1 and button in selected_buttons:
            edit_selected_buttons()
        else:
            menu = tk.Menu(root, tearoff=0)
            menu.add_command(label="서식 복사", command=copy_format)
            menu.add_command(label="복사", command=copy_button)
            menu.add_command(label="붙여넣기", command=lambda: on_paste_button(event))
            menu.add_separator()
            menu.add_command(label="이름 변경", command=rename_button)
            menu.add_command(label="링크 변경", command=change_path)
            menu.add_command(label="색상 변경", command=change_color)
            menu.add_command(label="글자색 변경", command=change_text_color)
            menu.add_command(label="글자 두께 변경", command=toggle_font_weight)
            menu.add_command(label="글자 크기 변경", command=change_font_size)
            menu.add_command(label="버튼 크기 변경", command=change_button_size)
            menu.add_command(label="비밀번호 변경", command=change_password)
            menu.add_command(label="핫키 설정", command=set_hotkey)
            menu.add_command(label="시작 버튼 제거", command=delete_button)
            menu.post(event.x_root, event.y_root)
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
# 빈 공간에서 우클릭 시 호출되는 함수 (기존 코드 끝에 추가하거나 별도 함수로 붙여넣기)


def on_right_click_root(event):
    global copied_button_data, copied_label_data
    print(f"[DEBUG] on_right_click_root 호출됨: x={event.x}, y={event.y}")

    def paste_button():
        global copied_button_data
        print(f"[DEBUG] paste_button 호출됨")
        if not os.path.exists(COPY_FILE_PATH):
            print(f"[DEBUG] paste_button: 복사된 데이터 파일 없음 ({COPY_FILE_PATH})")
            return
        
        try:
            with open(COPY_FILE_PATH, "r", encoding="utf-8") as f:
                copied_data = json.load(f)
            print(f"[DEBUG] 복사된 데이터 로드됨: {copied_data}")
        except Exception as e:
            print(f"[DEBUG] paste_button: 파일 읽기 오류 - {e}")
            return
        
        if not copied_data:
            print("[DEBUG] paste_button: 복사된 데이터 없음")
            return

        if isinstance(copied_data, dict) and "name" not in copied_data:
            print("[DEBUG] paste_button: 버튼 데이터가 아님 (name 키 없음), 붙여넣기 생략")
            return

        paste_x = event.x_root - content_frame.winfo_rootx()
        paste_y = event.y_root - content_frame.winfo_rooty()
        existing_buttons = [btn.cget("text") for btn in content_frame.winfo_children() if isinstance(btn, tk.Button)]
        print(f"[DEBUG] 기존 버튼 이름 목록: {existing_buttons}")

        new_buttons = []
        if isinstance(copied_data, list):
            if copied_data:
                base_x = copied_data[0]["x"]
                base_y = copied_data[0]["y"]
            else:
                base_x, base_y = 0, 0

            for btn_data in copied_data:
                base_name = btn_data["name"]
                new_name = base_name
                counter = 1
                while new_name in existing_buttons or f"{new_name} 📁" in existing_buttons or os.path.exists(os.path.join(json_directory, f"{new_name}_button.json")):
                    new_name = f"{base_name}_{counter}"
                    counter += 1
                display_text = f"{new_name} 📁" if btn_data.get("is_folder", False) else new_name
                existing_buttons.append(display_text)

                original_path = btn_data["path"]
                relative_x = btn_data["x"] - base_x
                relative_y = btn_data["y"] - base_y
                new_x = paste_x + relative_x
                new_y = paste_y + relative_y
                print(f"[DEBUG] 붙여넣기 시도: 이름={new_name}, 경로={original_path}, 위치=({new_x}, {new_y})")

                new_button = MyButton(
                    content_frame,
                    path=original_path,
                    text=display_text,
                    command=partial(run_file, original_path + ("|" + btn_data["password"] if btn_data["password"] else "")),
                    bg=btn_data["color"],
                    fg=btn_data["fg"],
                    font=("Arial", btn_data["font_size"], btn_data["font_weight"]),
                    width=btn_data["width"],
                    height=btn_data["height"],
                    is_copy=True
                )
                new_button.place(x=new_x, y=new_y)
                new_button.password = btn_data["password"]
                new_button.original_bg = btn_data["color"]
                new_button.font_size = btn_data["font_size"]
                new_button.font_weight = btn_data["font_weight"]
                new_button.hotkey = btn_data.get("hotkey", None)
                new_button.is_folder = btn_data.get("is_folder", False)

                adjust_font_size(new_button, display_text, max_width=150)

                # JSON 파일에 저장 (고유 파일명 보장)
                json_file_path = os.path.join(json_directory, f"{new_name}_button.json")
                with open(json_file_path, "w", encoding="utf-8") as f:
                    json.dump({
                        "name": new_name,
                        "path": original_path,
                        "x": new_button.winfo_x(),
                        "y": new_button.winfo_y(),
                        "color": new_button.original_bg,
                        "password": new_button.password,
                        "font_size": new_button.font_size,
                        "fg": new_button.cget("fg"),
                        "font_weight": new_button.font_weight,
                        "width": new_button.cget("width"),
                        "height": new_button.cget("height"),
                        "hotkey": new_button.hotkey,
                        "is_folder": new_button.is_folder
                    }, f, ensure_ascii=False, indent=4)
                new_buttons.append(new_name)
                print(f"[DEBUG] 버튼 생성 및 JSON 저장 완료: {new_name} at ({new_button.winfo_x()}, {new_button.winfo_y()})")

            print(f"[DEBUG] 다수 버튼 붙여넣기 완료: {new_buttons}")
        else:
            base_name = copied_data["name"]
            new_name = base_name
            counter = 1
            while new_name in existing_buttons or f"{new_name} 📁" in existing_buttons or os.path.exists(os.path.join(json_directory, f"{new_name}_button.json")):
                new_name = f"{base_name}_{counter}"
                counter += 1
            display_text = f"{new_name} 📁" if copied_data.get("is_folder", False) else new_name

            original_path = copied_data["path"]
            new_x = paste_x
            new_y = paste_y
            print(f"[DEBUG] 단일 붙여넣기 시도: 이름={new_name}, 경로={original_path}, 위치=({new_x}, {new_y})")

            new_button = MyButton(
                content_frame,
                path=original_path,
                text=display_text,
                command=partial(run_file, original_path + ("|" + copied_data["password"] if copied_data["password"] else "")),
                bg=copied_data["color"],
                fg=copied_data["fg"],
                font=("Arial", copied_data["font_size"], copied_data["font_weight"]),
                width=copied_data["width"],
                height=copied_data["height"],
                is_copy=True
            )
            new_button.place(x=new_x, y=new_y)
            new_button.password = copied_data["password"]
            new_button.original_bg = copied_data["color"]
            new_button.font_size = copied_data["font_size"]
            new_button.font_weight = copied_data["font_weight"]
            new_button.hotkey = copied_data.get("hotkey", None)
            new_button.is_folder = copied_data.get("is_folder", False)

            adjust_font_size(new_button, display_text, max_width=150)

            # JSON 파일에 저장 (고유 파일명 보장)
            json_file_path = os.path.join(json_directory, f"{new_name}_button.json")
            with open(json_file_path, "w", encoding="utf-8") as f:
                json.dump({
                    "name": new_name,
                    "path": original_path,
                    "x": new_button.winfo_x(),
                    "y": new_button.winfo_y(),
                    "color": new_button.original_bg,
                    "password": new_button.password,
                    "font_size": new_button.font_size,
                    "fg": new_button.cget("fg"),
                    "font_weight": new_button.font_weight,
                    "width": new_button.cget("width"),
                    "height": new_button.cget("height"),
                    "hotkey": new_button.hotkey,
                    "is_folder": new_button.is_folder
                }, f, ensure_ascii=False, indent=4)
            new_buttons.append(new_name)
            print(f"[DEBUG] 단일 버튼 붙여넣기 및 JSON 저장 완료: 이름={new_name}, path={original_path}, 위치=({new_x}, {new_y})")

        try:
            os.remove(COPY_FILE_PATH)
            print(f"[DEBUG] 복사 데이터 삭제됨: {COPY_FILE_PATH}")
        except OSError as e:
            print(f"[DEBUG] 복사 데이터 삭제 실패: {e}")
            try:
                with open(COPY_FILE_PATH, "w", encoding="utf-8") as f:
                    json.dump({}, f)
                print(f"[DEBUG] 복사 데이터 초기화됨: {COPY_FILE_PATH}")
            except Exception as e2:
                print(f"[DEBUG] 복사 데이터 초기화 실패: {e2}")
        
        copied_button_data = None
        print(f"[DEBUG] copied_button_data 초기화됨")

    def paste_label():
        global copied_label_data
        if not copied_label_data:
            print("[DEBUG] paste_label: 복사된 라벨 데이터 없음")
            return
        
        # 라벨 이름은 중복되어도 되므로(uid 로 구분) 복사한 이름을 그대로 쓴다.
        new_name = copied_label_data["text"]

        paste_x = event.x_root - content_frame.winfo_rootx()
        paste_y = event.y_root - content_frame.winfo_rooty()
        new_label = MyLabel(
            content_frame,
            text=new_name,
            fg=copied_label_data["color"],
            font=(copied_label_data["font_name"], copied_label_data["font_size"], copied_label_data["font_style"]),
            bg="white"
        )
        new_label.place(x=paste_x, y=paste_y)
        new_label.bind("<ButtonPress-1>", lambda e, l=new_label: on_label_press(l, e))
        new_label.bind("<B1-Motion>", lambda e, l=new_label: on_label_drag(l, e))
        new_label.bind("<ButtonRelease-1>", lambda e, l=new_label: on_label_release(l, e))
        new_label.bind("<Button-3>", lambda e, l=new_label: on_right_click_label(l, e))
        save_label_state(
            new_label.cget("text"), paste_x, paste_y, new_label.cget("fg"),
            new_label.cget("font").split()[0], int(new_label.cget("font").split()[1]),
            new_label.cget("font").split()[2] if len(new_label.cget("font").split()) > 2 else "normal",
            uid=label_uid(new_label)
        )
        print(f"[DEBUG] 라벨 붙여넣기 완료: {new_name} at ({paste_x}, {paste_y})")
        copied_label_data = None

    def download_file_from_button():
        global selected_buttons
        
        # 저장할 기본 폴더 선택
        save_dir = filedialog.askdirectory(title="저장할 폴더 선택")
        if not save_dir:
            print("[DEBUG] 다운로드 취소: 저장 폴더 미선택")
            return
        
        # 선택된 버튼이 없는 경우 파일만 다운로드
        if not selected_buttons:
            print("[DEBUG] 선택된 버튼 없음: 파일만 다운로드 시작")
            downloaded_items = []
            failed_items = []
            for button in MyButton._created_buttons.values():
                source_path = button.path
                item_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")
                
                # 파일인지 폴더인지 확인
                is_folder = getattr(button, "is_folder", os.path.isdir(source_path))
                
                if is_folder:
                    print(f"[DEBUG] 다운로드 건너뜀: 폴더는 선택 시에만 다운로드 - {source_path}")
                    continue
                
                if not os.path.exists(source_path):
                    print(f"[DEBUG] 다운로드 건너뜀: 파일 없음 - {source_path}")
                    failed_items.append(item_name)
                    continue
                
                # 저장 경로 생성 (파일은 상위 폴더 이름 사용)
                parent_dir = os.path.basename(os.path.dirname(source_path)) or "UnknownFolder"
                final_save_dir = os.path.join(save_dir, parent_dir)
                try:
                    os.makedirs(final_save_dir, exist_ok=True)
                    print(f"[DEBUG] 저장 폴더 생성/확인: {final_save_dir}")
                except Exception as e:
                    print(f"[ERROR] 저장 폴더 생성 실패: {e}")
                    failed_items.append(item_name)
                    continue
                
                # 파일 복사
                destination_path = os.path.join(final_save_dir, os.path.basename(source_path))
                try:
                    shutil.copy2(source_path, destination_path)  # 파일 복사
                    downloaded_items.append(destination_path)
                    print(f"[DEBUG] 파일 복사 성공: {source_path} -> {destination_path}")
                except Exception as e:
                    print(f"[ERROR] 파일 복사 실패: {e}")
                    failed_items.append(item_name)
            
            # 결과 메시지 표시 및 메인 폴더 열기
            if downloaded_items:
                messagebox.showinfo("성공", f"{len(downloaded_items)}개 파일이 성공적으로 저장되었습니다:\n" + "\n".join(downloaded_items[:5]) + (f"\n...외 {len(downloaded_items)-5}개" if len(downloaded_items) > 5 else ""))
                try:
                    os.startfile(save_dir)  # 메인 폴더 열기
                    print(f"[DEBUG] 메인 폴더 열기 성공: {save_dir}")
                except Exception as e:
                    print(f"[ERROR] 메인 폴더 열기 실패: {e}")
            if failed_items:
                messagebox.showwarning("경고", f"{len(failed_items)}개 파일 다운로드 실패:\n" + "\n".join(failed_items[:5]) + (f"\n...외 {len(failed_items)-5}개" if len(failed_items) > 5 else ""))
            if not downloaded_items and not failed_items:
                messagebox.showinfo("알림", "다운로드할 파일이 없습니다.")
            return
        
        # 선택된 버튼이 있는 경우: 단일 또는 다중 파일/폴더 다운로드
        downloaded_items = []
        failed_items = []
        
        for button in selected_buttons:
            source_path = button.path
            item_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")
            
            # 파일인지 폴더인지 확인
            is_folder = getattr(button, "is_folder", os.path.isdir(source_path))
            
            if not os.path.exists(source_path):
                print(f"[DEBUG] 다운로드 건너뜀: 경로 없음 - {source_path}")
                failed_items.append(item_name)
                continue
            
            # 저장 경로 생성 (폴더는 현재 폴더 이름, 파일은 상위 폴더 이름 사용)
            if is_folder:
                final_save_dir = os.path.join(save_dir, os.path.basename(source_path))
            else:
                parent_dir = os.path.basename(os.path.dirname(source_path)) or "UnknownFolder"
                final_save_dir = os.path.join(save_dir, parent_dir)
            
            try:
                if not is_folder:  # 파일일 경우에만 상위 폴더 생성
                    os.makedirs(final_save_dir, exist_ok=True)
                print(f"[DEBUG] 저장 폴더 생성/확인: {final_save_dir}")
            except Exception as e:
                print(f"[ERROR] 저장 폴더 생성 실패: {e}")
                failed_items.append(item_name)
                continue
            
            # 파일 또는 폴더 복사
            destination_path = final_save_dir if is_folder else os.path.join(final_save_dir, os.path.basename(source_path))
            try:
                if is_folder:
                    if os.path.exists(destination_path):
                        shutil.rmtree(destination_path)  # 기존 폴더 삭제
                    shutil.copytree(source_path, destination_path)  # 폴더 복사
                    print(f"[DEBUG] 폴더 복사 성공: {source_path} -> {destination_path}")
                else:
                    shutil.copy2(source_path, destination_path)  # 파일 복사
                    print(f"[DEBUG] 파일 복사 성공: {source_path} -> {destination_path}")
                downloaded_items.append(destination_path)
            except Exception as e:
                print(f"[ERROR] 복사 실패: {e}")
                failed_items.append(item_name)
        
        # 결과 메시지 표시 및 메인 폴더 열기
        if downloaded_items:
            messagebox.showinfo("성공", f"{len(downloaded_items)}개 항목이 성공적으로 저장되었습니다:\n" + "\n".join(downloaded_items[:5]) + (f"\n...외 {len(downloaded_items)-5}개" if len(downloaded_items) > 5 else ""))
            try:
                os.startfile(save_dir)  # 메인 폴더 열기
                print(f"[DEBUG] 메인 폴더 열기 성공: {save_dir}")
            except Exception as e:
                print(f"[ERROR] 메인 폴더 열기 실패: {e}")
        if failed_items:
            messagebox.showwarning("경고", f"{len(failed_items)}개 항목 다운로드 실패:\n" + "\n".join(failed_items[:5]) + (f"\n...외 {len(failed_items)-5}개" if len(failed_items) > 5 else ""))
        if not downloaded_items and not failed_items:
            messagebox.showinfo("알림", "다운로드할 항목이 없습니다.")

    def create_new_label():  # 새 라벨 추가 함수 정의
        paste_x = event.x_root - content_frame.winfo_rootx()
        paste_y = event.y_root - content_frame.winfo_rooty()
        new_label = MyLabel(
            content_frame,
            text="새 라벨",
            fg="black",
            font=("Arial", 10, "normal"),
            bg="white"
        )
        new_label.place(x=paste_x, y=paste_y)
        new_label.bind("<ButtonPress-1>", lambda e, l=new_label: on_label_press(l, e))
        new_label.bind("<B1-Motion>", lambda e, l=new_label: on_label_drag(l, e))
        new_label.bind("<ButtonRelease-1>", lambda e, l=new_label: on_label_release(l, e))
        new_label.bind("<Button-3>", lambda e, l=new_label: on_right_click_label(l, e))
        save_label_state(
            new_label.cget("text"), paste_x, paste_y, new_label.cget("fg"),
            "Arial", 10, "normal", uid=label_uid(new_label)
        )
        print(f"[DEBUG] 새 라벨 생성 완료: 위치=({paste_x}, {paste_y})")

    widget = root.winfo_containing(event.x_root, event.y_root)
    if isinstance(widget, MyButton):
        print("[DEBUG] 버튼 위에서 우클릭 - on_right_click_button 호출")
        return

    menu = tk.Menu(root, tearoff=0)
    show_menu = False

    #menu.add_command(label="새 버튼 추가", command=add_entry)
    #menu.add_command(label="새 라벨 추가", command=create_new_label)
    menu.add_command(label="파일 다운로드", command=download_file_from_button)
    show_menu = True

    if os.path.exists(COPY_FILE_PATH):
        try:
            with open(COPY_FILE_PATH, "r", encoding="utf-8") as f:
                copied_data_check = json.load(f)
            if copied_data_check and isinstance(copied_data_check, (dict, list)) and ("name" in copied_data_check or (isinstance(copied_data_check, list) and len(copied_data_check) > 0 and "name" in copied_data_check[0])):
                menu.add_command(label="버튼 붙여넣기", command=paste_button)
                print(f"[DEBUG] 메뉴 표시 준비: COPY_FILE_PATH 존재 여부=True, 데이터 유효")
                show_menu = True
            else:
                print(f"[DEBUG] COPY_FILE_PATH 데이터가 버튼 형식이 아님: {copied_data_check}")
        except Exception as e:
            print(f"[DEBUG] COPY_FILE_PATH 읽기 오류: {e}, 버튼 붙여넣기 메뉴 생략")

    if copied_label_data:
        menu.add_command(label="라벨 붙여넣기", command=paste_label)
        print(f"[DEBUG] 메뉴 표시 준비: copied_label_data 존재 여부=True")
        show_menu = True

    try:
        menu.add_command(label="엑셀로 내보내기", command=export_to_excel)
        show_menu = True
    except NameError as e:
        print(f"[ERROR] 메뉴 추가 실패: {e}. export_to_excel 함수가 정의되지 않음")

    if grid_on:
        menu.add_command(label="그리드 끄기", command=toggle_grid)
    else:
        menu.add_command(label="그리드 켜기", command=toggle_grid)
    show_menu = True

    if show_menu:
        menu.post(event.x_root, event.y_root)
    else:
        print(f"[DEBUG] 복사된 데이터 없음: 메뉴 표시 생략")
        
        
        
        
        
        
        
        
        
        

def add_label():
    new_label = MyLabel(content_frame, text="New Label", fg="black", font=("Arial", 12), bg="white")
    new_label.place(x=100, y=100)
    print(f"[DEBUG] 새 라벨 추가됨: New Label")        

# 엑셀 내보내기 함수 정의 (요청에 따라 추가)






# 색상 이름-to-HEX 매핑
COLOR_MAP = {
    "white": "FFFFFF",
    "black": "000000",
    "red": "FF0000",
    "green": "00FF00",
    "blue": "0000FF",
}

def normalize_color(color):
    """색상 값을 HEX 코드로 정규화"""
    if not color or not isinstance(color, str):
        return "000000"  # 기본값: 검정
    color = color.strip().lower()
    if color.startswith("#"):
        return color.lstrip("#")
    return COLOR_MAP.get(color, "000000")  # 색상 이름이면 매핑, 아니면 기본값

def export_to_excel(event=None):
    # 버튼 데이터 수집: 선택된 버튼이 있으면 그것만, 없으면 전체 버튼
    if selected_buttons:  # 선택된 버튼이 있는 경우
        buttons = selected_buttons
        print("[DEBUG] 선택된 버튼만 엑셀로 내보냄")
    else:  # 선택된 버튼이 없는 경우
        buttons = [w for w in content_frame.winfo_children() if isinstance(w, MyButton)]
        print("[DEBUG] 전체 버튼을 엑셀로 내보냄")
    
    if not buttons:
        messagebox.showinfo("알림", "내보낼 버튼이 없습니다.")
        return
    
    # JSON 파일 경로 설정 - 동적으로 exe_name 결정
    exe_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    json_directory = f"C:/json/{exe_name}"
    
    # JSON 디렉토리가 존재하지 않으면 생성
    if not os.path.exists(json_directory):
        os.makedirs(json_directory)
        print(f"[DEBUG] JSON 디렉토리 생성: {json_directory}")
    
    # 버튼별 JSON 파일 읽기
    button_dict = {}
    for button in buttons:
        button_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")  # 폴더 아이콘 제거
        json_file = os.path.join(json_directory, f"{sanitize_filename(button_name)}_button.json")
        try:
            with open(json_file, "r", encoding="utf-8") as f:
                button_data = json.load(f)
                button_dict[button_name] = button_data
                print(f"[DEBUG] JSON 파일 로드 성공: {json_file}")
                print(f"[DEBUG] JSON 데이터: {json.dumps(button_data, indent=2, ensure_ascii=False)}")
        except FileNotFoundError:
            print(f"[DEBUG] JSON 파일 없음: {json_file}")
            # JSON 파일이 없어도 기본 데이터로 처리
            button_dict[button_name] = {
                "name": button_name,
                "path": button.path,
                "color": button.original_bg,
                "fg": button.original_fg,
                "font_weight": button.font_weight,
                "is_folder": button.is_folder
            }
        except Exception as e:
            print(f"[DEBUG] JSON 파일 로드 실패: {json_file}, 오류: {e}")
    
    # 버튼 데이터를 배경색 기준으로 정렬
    button_list = []
    for button in buttons:
        button_name = button.cget("text").replace(" 📁", "").replace(" ↲", "")
        json_info = button_dict.get(button_name, {})
        bg_color = normalize_color(json_info.get("color", button.original_bg))
        link = auto_change_slash(json_info.get("path", button.path)) if json_info.get("path", button.path) else "No Link"
        fg_color = normalize_color(json_info.get("fg", button.original_fg))
        font_weight = json_info.get("font_weight", button.font_weight)
        is_folder = json_info.get("is_folder", button.is_folder)
        
        # 폴더 링크 오류 방지
        if is_folder and not os.path.isdir(link):
            link = f"Invalid Folder Path: {link}"
        
        button_list.append((bg_color, button_name, link, fg_color, font_weight))
    
    # 배경색을 기준으로 정렬 (HEX 코드로 문자열 비교)
    button_list.sort(key=lambda x: x[0])
    
    # 엑셀 워크북 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "Button Details"
    
    # 눈금선 비활성화
    ws.sheet_view.showGridLines = False
    
    # 헤더 설정 - A열: 빈 열, B열: Button Name, C열: Link
    headers = ["", "Button Name", "Link"]
    ws.append(headers)
    
    # 헤더 스타일 적용 (A열 제외)
    header_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col, cell in enumerate(ws[1], start=1):
        if col != 1:  # A열은 스타일 적용 안 함
            cell.fill = header_fill
            cell.font = header_font
    
    # 정렬된 버튼 데이터 추가
    for idx, (bg_color, button_name, link, fg_color, font_weight) in enumerate(button_list, start=2):
        # 디버깅 로그
        print(f"[DEBUG] 버튼 이름: {button_name}")
        print(f"[DEBUG] 배경색: #{bg_color}, 링크: {link}, 글자색: #{fg_color}, 글자두께: {font_weight}")
        
        # A열: 빈 상태로 유지 (데이터나 색상 추가 안 함)
        
        # B열: 버튼 이름 (HYPERLINK 수식 추가)
        name_cell = ws[f"B{idx}"]
        if link and not link.startswith("Invalid Folder Path"):  # 유효한 링크만 수식 적용
            name_cell.value = f'=HYPERLINK("{link}", "{button_name}")'  # C열 링크 참조
        else:
            name_cell.value = button_name  # 링크가 없거나 유효하지 않으면 일반 텍스트
        
        # B열 셀에 배경색 적용
        cell_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        name_cell.fill = cell_fill
        
        # B열 셀에 글자색 및 글자두께 적용
        is_bold = font_weight.lower() == "bold"
        name_cell.font = Font(color=fg_color, bold=is_bold)
        
        # C열: 링크
        ws[f"C{idx}"].value = link
    
    # 열 너비 고정 설정
    ws.column_dimensions["A"].width = 3   # A열 너비 3
    ws.column_dimensions["B"].width = 25  # B열 너비 25
    ws.column_dimensions["C"].width = 30  # C열 너비 30
    
    # 파일 저장 및 열기
    default_filename = f"Button_Details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        initialfile=default_filename,
        title="엑셀 파일 저장"
    )
    
    if file_path:
        try:
            wb.save(file_path)
            messagebox.showinfo("성공", f"엑셀 파일이 저장되었습니다: {file_path}")
            print(f"[DEBUG] 엑셀 파일 저장 성공: {file_path}")
            
            # 운영 체제에 따라 엑셀 파일 자동 열기
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            elif system == "Linux":
                subprocess.run(["xdg-open", file_path])
            else:
                print(f"[DEBUG] 지원되지 않는 운영 체제: {system}")
            
        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 저장 또는 열기 중 오류 발생: {e}")
            print(f"[DEBUG] 엑셀 파일 처리 실패: {e}")
    else:
        print("[DEBUG] 엑셀 파일 저장 취소됨")
        
        
        
        
      
            
def copy_button(button):
    global copied_button_data
    copied_button_data = {
        "name": button.cget("text").replace(" 📁", "").replace(" ↲", ""),
        "path": button.path,
        "color": button.original_bg,
        "fg": button.original_fg,
        "font_size": button.font_size,
        "font_weight": button.font_weight,
        "width": button.original_width,
        "height": button.original_height,
        "password": button.password,
        "is_folder": button.is_folder
    }
    print(f"[DEBUG] 버튼 복사됨: {copied_button_data['name']}")

# 버튼 붙여넣기 함수
def on_paste_button(event=None):
    global copied_button_data
    if not copied_button_data:
        print("[DEBUG] 복사된 버튼 데이터가 없습니다.")
        return

    mouse_x = root.winfo_pointerx() - content_frame.winfo_rootx()
    mouse_y = root.winfo_pointery() - content_frame.winfo_rooty()
    print(f"[DEBUG] 마우스 클릭 위치: ({mouse_x}, {mouse_y})")

    window_width = root.winfo_width()
    window_height = root.winfo_height()
    print(f"[DEBUG] GUI 창 크기: ({window_width}, {window_height})")

    new_x = mouse_x
    new_y = mouse_y

    button_width = copied_button_data.get("width", 15) * 10
    button_height = copied_button_data.get("height", 1) * 20

    if new_x + button_width > window_width:
        new_x = window_width - button_width - 5
        print(f"[DEBUG] 오른쪽 경계 조정: new_x={new_x}")
    if new_x < 0:
        new_x = 5
        print(f"[DEBUG] 왼쪽 경계 조정: new_x={new_x}")
    if new_y + button_height > window_height:
        new_y = window_height - button_height - 5
        print(f"[DEBUG] 아래쪽 경계 조정: new_y={new_y}")
    if new_y < 0:
        new_y = 5
        print(f"[DEBUG] 위쪽 경계 조정: new_y={new_y}")

    existing_buttons = [b for b in content_frame.winfo_children() if isinstance(b, MyButton)]
    max_y = new_y
    for button in existing_buttons:
        bx = button.winfo_x()
        by = button.winfo_y()
        bw = button.winfo_width()
        bh = button.winfo_height()
        if abs(bx - new_x) < bw:
            if by + bh > max_y:
                max_y = by + bh + 5

    new_y = max_y
    print(f"[DEBUG] 최종 버튼 위치: ({new_x}, {new_y})")

    name = copied_button_data.get("name", "새 버튼")
    path = copied_button_data.get("path", "")
    color = copied_button_data.get("color", "#4CAF50")
    fg = copied_button_data.get("fg", "white")
    font_size = copied_button_data.get("font_size", 10)
    font_weight = copied_button_data.get("font_weight", "normal")
    width = copied_button_data.get("width", 15)
    height = copied_button_data.get("height", 1)
    password = copied_button_data.get("password", "")
    is_folder = copied_button_data.get("is_folder", False)

    existing_names = [b.cget("text").replace(" 📁", "").replace(" ↲", "") for b in existing_buttons]
    base_name = name.replace(" 📁", "").replace(" ↲", "")
    counter = 1
    new_name = base_name
    while new_name in existing_names:
        new_name = f"{base_name}_{counter}"
        counter += 1
    display_text = f"{new_name} 📁" if is_folder else new_name

    button = MyButton(
        content_frame,
        text=display_text,
        path=path,
        password=password,
        command=partial(run_file, path + ("|" + password if password else "")),
        bg=color,
        fg=fg,
        font=("Arial", font_size, font_weight),
        width=width,
        height=height,
        is_folder=is_folder,
        is_copy=False
    )
    button.place(x=new_x, y=new_y)
    button.original_bg = color
    button.original_fg = fg
    button.original_width = width
    button.original_height = height
    button.font_size = font_size
    button.font_weight = font_weight

    adjust_font_size(button, display_text, max_width=150)
    # MyButton.__init__에서 이벤트 바인딩 자동 처리됨

    button_data = {
        "name": new_name,
        "path": path,
        "x": new_x,
        "y": new_y,
        "color": color,
        "password": password,
        "font_size": font_size,
        "fg": fg,
        "font_weight": font_weight,
        "width": width,
        "height": height,
        "is_folder": is_folder,
        "selected": False
    }
    update_json_file(new_name, json_directory, button_data)
    print(f"[DEBUG] 새 버튼 생성 및 저장 완료: {new_name}, 위치: ({new_x}, {new_y})")
    
    
    
            
            
            
            
                       
            

def toggle_label_bold(label):
    current_font = label.cget("font")
    font_name, font_size, *font_style = current_font.split()
    if "bold" in font_style:
        label.config(font=(font_name, font_size))
        font_style = "normal"
    else:
        label.config(font=(font_name, font_size, "bold"))
        font_style = "bold"
    save_label_state(label.cget("text"), label.winfo_x(), label.winfo_y(), label.cget("fg"), font_name, font_size, font_style, uid=label_uid(label))
    
    

def edit_selected_labels(event=None):
    if not selected_labels:
        messagebox.showinfo("알림", "선택된 라벨이 없습니다.")
        return

    edit_window = tk.Toplevel(root)
    edit_window.title("선택된 라벨 편집")
    edit_window_width = 430
    edit_window_height = 450  # 간격 입력 UI 포함
    edit_window.configure(bg="white")
    edit_window.withdraw()

    # 첫 번째 선택된 라벨 위치 계산
    first_label = selected_labels[0]
    root.update_idletasks()
    label_name = first_label.cget("text")
    json_file = next((p for p in label_json_paths(first_label) if os.path.exists(p)), None)
    if json_file:
        with open(json_file, "r", encoding='utf-8') as f:
            label_data = json.load(f)
            json_x = label_data.get("x", first_label.winfo_x())
            json_y = label_data.get("y", first_label.winfo_y())
    else:
        json_x = first_label.winfo_x()
        json_y = first_label.winfo_y()

    ctrl_pressed = event is not None and (event.state & 0x0004) != 0
    if ctrl_pressed:
        pos_x = root.winfo_x() + json_x
        pos_y = root.winfo_y() + json_y
    else:
        pos_x = first_label.winfo_rootx()
        pos_y = first_label.winfo_rooty()

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    pos_x = max(0, min(pos_x, screen_width - edit_window_width))
    pos_y = max(0, min(pos_y, screen_height - edit_window_height))

    edit_window.geometry(f"{edit_window_width}x{edit_window_height}+{pos_x}+{pos_y}")
    edit_window.attributes('-topmost', True)
    edit_window.update_idletasks()
    edit_window.deiconify()

    tk.Label(edit_window, text=f"선택된 라벨 수: {len(selected_labels)}", bg="white").pack(pady=5)

    # 글자색상
    tk.Label(edit_window, text="글자색상:", bg="white").pack(pady=5)
    fg_color_button = tk.Button(edit_window, text="색상 선택", command=lambda: choose_fg_color(fg_color_var))
    fg_color_button.pack(pady=5)
    fg_color_var = tk.StringVar(value=selected_labels[0].cget("fg"))

    # 글자 크기
    tk.Label(edit_window, text="글자 크기:", bg="white").pack(pady=5)
    font_size_var = tk.IntVar(value=int(selected_labels[0].cget("font").split()[1]))
    font_size_spinbox = tk.Spinbox(edit_window, from_=8, to=72, textvariable=font_size_var)
    font_size_spinbox.pack(pady=5)

    # 글자 두께
    tk.Label(edit_window, text="글자 두께:", bg="white").pack(pady=5)
    font_weight_var = tk.StringVar(value="bold" if "bold" in selected_labels[0].cget("font") else "normal")
    font_weight_combo = ttk.Combobox(edit_window, textvariable=font_weight_var, values=["normal", "bold"])
    font_weight_combo.pack(pady=5)

    temp_positions = {}  # 간격 조정 및 정렬용 임시 위치 저장
    preview_labels = {}
    original_positions = {}  # 취소 시 복원을 위한 원래 위치 저장

    # 정렬 함수
    def align_top():
        if len(selected_labels) < 2:
            return
        min_y = min(lbl.winfo_y() for lbl in selected_labels)
        for lbl in selected_labels:
            temp_positions[lbl] = (lbl.winfo_x(), min_y)
        draw_preview()

    def align_bottom():
        if len(selected_labels) < 2:
            return
        max_y = max(lbl.winfo_y() + lbl.winfo_height() for lbl in selected_labels)
        for lbl in selected_labels:
            new_y = max_y - lbl.winfo_height()
            temp_positions[lbl] = (lbl.winfo_x(), new_y)
        draw_preview()

    def align_left():
        if len(selected_labels) < 2:
            return
        min_x = min(lbl.winfo_x() for lbl in selected_labels)
        for lbl in selected_labels:
            temp_positions[lbl] = (min_x, lbl.winfo_y())
        draw_preview()

    def align_right():
        if len(selected_labels) < 2:
            return
        max_x = max(lbl.winfo_x() + lbl.winfo_width() for lbl in selected_labels)
        for lbl in selected_labels:
            new_x = max_x - lbl.winfo_width()
            temp_positions[lbl] = (new_x, lbl.winfo_y())
        draw_preview()

    # 가로 간격 적용 함수
    def apply_x_spacing():
        if selected_labels and len(selected_labels) >= 3:
            labels_sorted_x = sorted(selected_labels, key=lambda l: l.winfo_x())
            first_lbl_x, last_lbl_x = labels_sorted_x[0], labels_sorted_x[-1]

            first_x, first_y = first_lbl_x.winfo_x(), first_lbl_x.winfo_y()
            last_x, last_y = last_lbl_x.winfo_x(), last_lbl_x.winfo_y()

            num_gaps = len(selected_labels) - 1
            total_x = last_x - first_x  # 가로 전체 거리

            if total_x <= 0:
                print("[DEBUG] 첫 번째와 마지막 라벨의 x 위치가 동일하거나 역순입니다.")
                return

            middle_labels = labels_sorted_x[1:-1]
            if middle_labels:
                x_step = total_x / num_gaps  # 균등 간격 계산
                for i, lbl in enumerate(middle_labels, 1):
                    new_x = first_x + i * x_step  # 첫 번째 라벨에서 시작하여 균등 배치
                    new_y = lbl.winfo_y()  # y 위치는 유지
                    lbl.place(x=new_x, y=new_y)
                    temp_positions[lbl] = (new_x, new_y)
                    print(f"[DEBUG] Label {i+1}: new_x = {new_x}, new_y = {new_y}, x_step = {x_step}")

            first_lbl_x.place(x=first_x, y=first_y)
            last_lbl_x.place(x=last_x, y=last_y)
            temp_positions[first_lbl_x] = (first_x, first_y)
            temp_positions[last_lbl_x] = (last_x, last_y)

    # 세로 간격 적용 함수
    def apply_y_spacing():
        if selected_labels and len(selected_labels) >= 3:
            labels_sorted_y = sorted(selected_labels, key=lambda l: l.winfo_y())
            first_lbl_y, last_lbl_y = labels_sorted_y[0], labels_sorted_y[-1]

            first_x, first_y = first_lbl_y.winfo_x(), first_lbl_y.winfo_y()
            last_x, last_y = last_lbl_y.winfo_x(), last_lbl_y.winfo_y()

            num_gaps = len(selected_labels) - 1
            total_y = last_y - first_y  # 세로 전체 거리

            if total_y <= 0:
                print("[DEBUG] 첫 번째와 마지막 라벨의 y 위치가 동일하거나 역순입니다.")
                return

            middle_labels = labels_sorted_y[1:-1]
            if middle_labels:
                y_step = total_y / num_gaps  # 균등 간격 계산
                for i, lbl in enumerate(middle_labels, 1):
                    new_y = first_y + i * y_step  # 첫 번째 라벨에서 시작하여 균등 배치
                    new_x = lbl.winfo_x()  # x 위치는 유지
                    lbl.place(x=new_x, y=new_y)
                    temp_positions[lbl] = (new_x, new_y)
                    print(f"[DEBUG] Label {i+1}: new_x = {new_x}, new_y = {new_y}, y_step = {y_step}")

            first_lbl_y.place(x=first_x, y=first_y)
            last_lbl_y.place(x=last_x, y=last_y)
            temp_positions[first_lbl_y] = (first_x, first_y)
            temp_positions[last_lbl_y] = (last_x, last_y)

    # 원래 위치 저장
    for lbl in selected_labels:
        original_positions[lbl] = (lbl.winfo_x(), lbl.winfo_y())

    def draw_preview():
        nonlocal preview_labels
        for lbl in preview_labels.values():
            lbl.destroy()
        preview_labels.clear()
        for lbl, (x, y) in temp_positions.items():
            preview_lbl = tk.Label(
                root,
                text=lbl.cget("text"),
                fg=lbl.cget("fg"),
                font=lbl.cget("font"),
                bg="white",
                relief="raised"
            )
            preview_lbl.place(x=x, y=y)
            root.attributes('-alpha', 0.5)
            root.after(100, lambda: root.attributes('-alpha', 1.0))
            preview_labels[lbl] = preview_lbl

    def clear_preview():
        nonlocal preview_labels
        for lbl in preview_labels.values():
            lbl.destroy()
        preview_labels.clear()

    # 정렬 UI
    tk.Label(edit_window, text="정렬:", bg="white").pack(pady=5)
    align_frame = tk.Frame(edit_window, bg="white")
    align_frame.pack(pady=5)
    tk.Button(align_frame, text="상", command=align_top).pack(side=tk.LEFT, padx=5)
    tk.Button(align_frame, text="하", command=align_bottom).pack(side=tk.LEFT, padx=5)
    tk.Button(align_frame, text="좌", command=align_left).pack(side=tk.LEFT, padx=5)
    tk.Button(align_frame, text="우", command=align_right).pack(side=tk.LEFT, padx=5)

    # 간격 UI
    tk.Label(edit_window, text="간격:", bg="white").pack(pady=5)
    spacing_frame = tk.Frame(edit_window, bg="white")
    spacing_frame.pack(pady=5)

    # 색상 선택
    def choose_fg_color(var):
        color = colorchooser.askcolor()[1]
        if color:
            var.set(color)

    # 변경 적용
    def apply_changes():
        try:
            new_fg = fg_color_var.get()
            new_font_size = font_size_var.get()
            new_font_weight = font_weight_var.get()

            for label in selected_labels[:]:
                label.config(fg=new_fg)
                new_font = ("Arial", new_font_size, new_font_weight)
                label.config(font=new_font)
                # 간격 조정 반영
                if label in temp_positions:
                    new_x, new_y = temp_positions[label]
                    label.place(x=new_x, y=new_y)
                label.selected = False
                label.config(relief="flat", bg="white")
                save_label_state(
                    label.cget("text"),
                    label.winfo_x(),
                    label.winfo_y(),
                    new_fg,
                    "Arial",
                    new_font_size,
                    new_font_weight,
                    uid=label_uid(label)
                )
                selected_labels.remove(label)
            clear_preview()
            edit_window.destroy()
        except Exception as e:
            messagebox.showerror("오류", f"변경 적용 중 오류 발생: {e}")

    def on_cancel():
        # 취소 시 원래 위치로 복원
        for lbl in selected_labels:
            if lbl in original_positions:
                orig_x, orig_y = original_positions[lbl]
                lbl.place(x=orig_x, y=orig_y)
        clear_preview()
        edit_window.destroy()

    # 버튼 프레임 (가로/세로 간격 조절 버튼 분리)
    button_frame = tk.Frame(edit_window, bg="white")
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="적용", command=apply_changes, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="취소", command=on_cancel, bg="#bdc3c7", fg="black").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="가로 간격", command=apply_x_spacing, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="세로 간격", command=apply_y_spacing, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)

    def on_edit_window_close():
        on_cancel()  # 창 닫기 시 취소 동작 수행
        edit_window.destroy()

    edit_window.protocol("WM_DELETE_WINDOW", on_edit_window_close)
    
    
    
    
    

    
def on_right_click_label(label, event):
    global copied_label_data  # 복사된 라벨 데이터를 저장하기 위한 전역 변수

    def change_label_name():
        # "라벨 이름 변경" GUI가 이미 열려 있는지 확인
        if hasattr(label, 'rename_window') and label.rename_window.winfo_exists():
            print(f"[DEBUG] 이미 '라벨 이름 변경' GUI가 열려 있음: {label.cget('text')}")
            return

        def save_new_label_name():
            new_label_name = label_entry.get()
            if new_label_name and new_label_name != label.cget("text"):
                old_label_name = label.cget("text")
                # 파일명이 uid 기반이므로 이름을 바꿔도 같은 파일에 그대로 저장된다.
                # (텍스트 기준으로 지우면 같은 이름의 다른 라벨 파일까지 지울 수 있다)
                label.config(text=new_label_name)
                label.alarm_shown = False
                save_label_state(
                    new_label_name, 
                    label.winfo_x(), 
                    label.winfo_y(), 
                    label.cget("fg"), 
                    label.cget("font").split(" ")[0], 
                    int(label.cget("font").split(" ")[1]), 
                    label.cget("font").split(" ")[2] if len(label.cget("font").split(" ")) > 2 else "normal",
                    uid=label_uid(label)
                )
                print(f"[DEBUG] 라벨 이름 변경: {old_label_name} -> {new_label_name}")
            rename_window.destroy()

        # "라벨 이름 변경" GUI 생성
        rename_window = tk.Toplevel(root)
        label.rename_window = rename_window  # 라벨 객체에 창 참조 저장
        rename_window.title("라벨 이름 변경")
        rename_window.geometry("300x150")
        rename_window.attributes('-topmost', True)
        rename_window.configure(bg="white")  # 흰색 배경 적용

        # 클릭된 위치에 GUI 표시
        x = event.x_root
        y = event.y_root
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        if x + 300 > screen_width:
            x = screen_width - 300
        if y + 150 > screen_height:
            y = screen_height - 150
        rename_window.geometry(f"+{x}+{y}")
        print(f"[DEBUG] '라벨 이름 변경' GUI 생성됨: 위치 ({x}, {y})")

        tk.Label(rename_window, text="새 이름 입력:", bg="white").pack(pady=10)
        label_entry = tk.Entry(rename_window, font=("Arial", 12))
        label_entry.pack(pady=5)
        label_entry.insert(0, label.cget("text"))
        label_entry.focus_set()
        tk.Button(rename_window, text="저장", command=save_new_label_name).pack(pady=5)

        # 창이 닫힐 때 참조 제거
        def on_rename_window_close():
            if hasattr(label, 'rename_window'):
                del label.rename_window
            rename_window.destroy()

        rename_window.protocol("WM_DELETE_WINDOW", on_rename_window_close)
        label_entry.bind("<Return>", lambda e: save_new_label_name())

    def change_label_color():
        color = colorchooser.askcolor()[1]
        if color:
            label.config(fg=color)
            save_label_state(
                label.cget("text"), 
                label.winfo_x(), 
                label.winfo_y(), 
                color, 
                label.cget("font").split(" ")[0], 
                int(label.cget("font").split(" ")[1]), 
                label.cget("font").split(" ")[2] if len(label.cget("font").split(" ")) > 2 else "normal",
                uid=label_uid(label)
            )

    def delete_label():
        label_name = label.cget("text")
        # 파일명이 uid 기반이므로 같은 이름의 다른 라벨은 영향받지 않는다.
        if not delete_label_json(label):
            print(f"[WARNING] 삭제 대상 라벨 JSON 이 없음: {label_name}")
        label.destroy()

    def toggle_bold():
        toggle_label_bold(label)

    def change_label_font_size():
        def save_new_font_size():
            new_size = size_entry.get()
            try:
                new_size = int(new_size)
                if 8 <= new_size <= 72:
                    font_name = label.cget("font").split(" ")[0]
                    font_style = "bold" if "bold" in label.cget("font") else "normal"
                    label.config(font=(font_name, new_size, font_style))
                    save_label_state(
                        label.cget("text"), 
                        label.winfo_x(), 
                        label.winfo_y(), 
                        label.cget("fg"), 
                        font_name, 
                        new_size, 
                        font_style,
                        uid=label_uid(label)
                    )
                    change_size_window.destroy()
                else:
                    messagebox.showerror("오류", "글자 크기는 8에서 72 사이의 값이어야 합니다.")
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자를 입력하세요.")
        change_size_window = tk.Toplevel(root)
        change_size_window.title("글자 크기 변경")
        change_size_window.geometry("250x150")
        change_size_window.configure(bg="white")  # 흰색 배경 적용
        tk.Label(change_size_window, text="글자 크기 입력 (8-72):", bg="white").pack(pady=10)
        size_entry = tk.Entry(change_size_window, font=("Arial", 12))
        size_entry.pack(pady=5)
        current_size = label.cget("font").split(" ")[1]
        size_entry.insert(0, current_size)
        tk.Button(change_size_window, text="저장", command=save_new_font_size).pack(pady=5)

    def open_calendar(label=None):
        def set_reservation():
            try:
                selected_date = cal.get_date()
                raw_time = f"{hour_entry.get()}:{minute_entry.get()}"
                exec_path = path_entry.get().strip()
                label_text = label_entry.get().strip()

                # 시간 형식 검증
                selected_time = raw_time.replace(":", "")  # 수정: 'selected from' -> 'selected_time'
                datetime.strptime(selected_date + selected_time, "%Y%m%d%H%M")
                hour, minute = map(int, raw_time.split(":"))
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError("시간은 00:00 ~ 23:59 사이여야 합니다.")

                formatted_time = f"{selected_date} {raw_time.replace(':', '')}"
                if exec_path:
                    reservation_text = f"실행 {formatted_time} {exec_path}"
                elif label_text:
                    reservation_text = f"알림 {formatted_time} {label_text}"
                else:
                    messagebox.showwarning("경고", "라벨 텍스트 또는 실행 경로를 입력하세요.")
                    return

                # 라벨 이름은 중복되어도 된다(uid 로 구분한다).
                if label:  # 기존 라벨 수정
                    old_label_name = label.cget("text")
                    label.config(text=reservation_text)
                    save_label_state(
                        reservation_text, 
                        label.winfo_x(), 
                        label.winfo_y(), 
                        label.cget("fg"), 
                        label.cget("font").split(" ")[0], 
                        int(label.cget("font").split(" ")[1]), 
                        label.cget("font").split(" ")[2] if len(label.cget("font").split(" ")) > 2 else "normal",
                        uid=label_uid(label)
                    )
                    print(f"[DEBUG] 라벨 수정됨: {reservation_text}")
                global is_tracking
                if not is_tracking:
                    is_tracking = True
                    check_meeting_time()
                cal_window.destroy()
            except ValueError as e:
                messagebox.showerror("오류", f"잘못된 시간 형식: {e}")
            except Exception as e:
                print(f"[DEBUG] 예약 설정 중 오류: {e}")
                messagebox.showerror("오류", f"예약 설정 중 문제 발생: {e}")

        cal_window = tk.Toplevel(root)
        cal_window.title("예약 설정")
        cal_window.geometry("400x500")
        cal_window.configure(bg="white")
        cal_window.attributes('-topmost', True)
        
        # 이벤트 위치에 따라 창 위치 조정
        cal_x = event.x_root - 200  # 창 중앙을 클릭 위치에 맞춤
        cal_y = event.y_root - 250
        cal_x = max(0, min(cal_x, root.winfo_screenwidth() - 400))
        cal_y = max(0, min(cal_y, root.winfo_screenheight() - 500))
        cal_window.geometry(f"400x500+{cal_x}+{cal_y}")
        print(f"[DEBUG] 캘린더 창 생성됨: 위치 ({cal_x}, {cal_y})")

        try:
            cal = TkCalendar(cal_window, selectmode="day", date_pattern="yyyymmdd")
            cal.pack(pady=10, fill="x")
            print("[DEBUG] 캘린더 위젯 추가됨")
        except Exception as e:
            print(f"[DEBUG] 캘린더 초기화 실패: {e}")
            messagebox.showerror("오류", "캘린더를 로드할 수 없습니다. tkcalendar 모듈을 확인하세요.")
            cal_window.destroy()
            return

        tk.Label(cal_window, text="시간 (HH:MM):", bg="white").pack(pady=5)
        time_frame = tk.Frame(cal_window, bg="white")
        time_frame.pack(pady=5)
        hour_entry = tk.Entry(time_frame, width=2)
        hour_entry.insert(0, "00")
        hour_entry.pack(side=tk.LEFT)
        tk.Label(time_frame, text=":", bg="white").pack(side=tk.LEFT)
        minute_entry = tk.Entry(time_frame, width=2)
        minute_entry.insert(0, "00")
        minute_entry.pack(side=tk.LEFT)
        print("[DEBUG] 시간 입력 필드 추가됨")

        tk.Label(cal_window, text="실행 경로 (파일/폴더/URL, 입력 시 실행 예약):", bg="white").pack(pady=5)
        path_entry = tk.Entry(cal_window, width=40)
        path_entry.pack(pady=5)
        tk.Button(cal_window, text="탐색", command=lambda: path_entry.insert(0, filedialog.askopenfilename() or filedialog.askdirectory())).pack(pady=5)
        print("[DEBUG] 경로 입력 필드 추가됨")

        tk.Label(cal_window, text="알림 텍스트 (선택):", bg="white").pack(pady=5)
        label_entry = tk.Entry(cal_window, width=40)
        if label:  # 기존 라벨 텍스트를 기본값으로 설정
            label_entry.insert(0, label.cget("text"))
        label_entry.pack(pady=5)

        tk.Button(cal_window, text="예약", command=set_reservation).pack(pady=10)
        print("[DEBUG] 예약 버튼 추가됨")

    def get_existing_labels():
        ui_labels = {child.cget("text") for child in content_frame.winfo_children() if isinstance(child, tk.Label)}
        json_labels = {data["text"] for data in load_labels()}  # JSON 파일에서 로드된 라벨 포함
        return ui_labels.union(json_labels)

    # 라벨 복사 기능 추가
    def copy_label():
        global copied_label_data
        copied_label_data = {
            "text": label.cget("text"),
            "color": label.cget("fg"),
            "font_name": label.cget("font").split()[0],
            "font_size": int(label.cget("font").split()[1]),
            "font_style": label.cget("font").split()[2] if len(label.cget("font").split()) > 2 else "normal"
        }
        print(f"[DEBUG] 라벨 복사됨: {copied_label_data}")

    if len(selected_labels) > 1 and label in selected_labels:
        edit_selected_labels()
    else:
        menu = tk.Menu(root, tearoff=0)
        menu.add_command(label="라벨 이름 변경", command=change_label_name)
        menu.add_command(label="라벨 글자색 변경", command=change_label_color)
        menu.add_command(label="글자 두께 변경", command=toggle_bold)
        menu.add_command(label="글자 크기 변경", command=change_label_font_size)
        menu.add_command(label="알림 설정", command=lambda: open_calendar(label))
        menu.add_command(label="실행 설정", command=lambda: open_calendar(label))
        menu.add_command(label="라벨 삭제", command=delete_label)
        menu.add_command(label="라벨 복사", command=copy_label)  # 복사 기능 추가
        menu.post(event.x_root, event.y_root)
        
        
        
        




# MyButton 클래스 수정 (복사된 버튼 실행 방지 강화)
class MyButton(tk.Button):
    _created_buttons = {}  # name -> button 인스턴스 매핑

    def __init__(self, master, path, password="", is_copy=False, *args, **kwargs):
        self.custom_command = kwargs.pop("command", None)
        kwargs.pop("selected", False)
        
        font_tuple = kwargs.pop("font", ("Arial", 10, "normal"))
        if isinstance(font_tuple, tuple) and len(font_tuple) >= 2:
            font_name = font_tuple[0]
            font_size = font_tuple[1]
            font_weight = font_tuple[2] if len(font_tuple) > 2 else "normal"
        else:
            font_name = "Arial"
            font_size = 10
            font_weight = "normal"

        self.is_folder = kwargs.pop("is_folder", os.path.isdir(path))
        display_text = kwargs.get("text", os.path.basename(path))
        if self.is_folder and "📁" not in display_text:
            kwargs["text"] = f"{display_text} 📁"
            kwargs["compound"] = kwargs.get("compound", "right")

        self.path = path
        self.name = kwargs["text"]  # 버튼의 고유 이름
        
        if is_copy:
            base_name = self.name
            counter = 1
            while self.name in MyButton._created_buttons or os.path.exists(os.path.join(json_directory, f"{self.name}_button.json")):
                self.name = f"{base_name}_{counter}"
                kwargs["text"] = f"{self.name} 📁" if self.is_folder else self.name
                counter += 1
                print(f"[DEBUG] 복사로 인한 이름 중복 감지됨: {base_name}, 새 이름 생성: {self.name}")

        # 바로가기 표시는 폴더의 " 📁" 와 똑같이 버튼 "글자"에 붙인다.
        # 자식 위젯으로 얹으면 드롭(모달 루프) 중에 만들어진 것이 그려지지 않아
        # "드롭 직후엔 안 보이고 재실행하면 보이는" 문제가 생긴다. 글자에 넣으면
        # 버튼이 직접 그리므로 항상 보인다.
        # self.name 에는 붙이지 않는다(JSON 파일명·식별자는 그대로 유지).
        if self.is_shortcut() and MyButton.SHORTCUT_MARK not in kwargs.get("text", ""):
            kwargs["text"] = f'{kwargs["text"]} {MyButton.SHORTCUT_MARK}'
            kwargs["compound"] = kwargs.get("compound", "right")

        super().__init__(master, *args, font=(font_name, font_size, font_weight), **kwargs)
        
        MyButton._created_buttons[self.name] = self
        
        self.password = password
        self.selected = False
        self.is_dragging = False
        self.was_dragged = False
        self.long_press_active = False
        self.press_time = 0
        self.press_x = 0
        self.press_y = 0
        self.original_bg = self.cget("bg")
        self.original_fg = self.cget("fg")
        self.original_width = self.cget("width")
        self.original_height = self.cget("height")
        self.relative_x = 0
        self.relative_y = 0
        self._lock = False
        self.font_size = font_size
        self.font_weight = font_weight
        self.hotkey = None
        self.last_drag_time = 0
        self.drag_interval = 50
        self.is_copied_format = False
        self._format_backup = None  # 서식 적용 전 모습(되돌리기용)
        self.drag_direction = None

        self.bind("<ButtonPress-1>", self.on_press)
        self.bind("<B1-Motion>", self.on_drag)
        self.bind("<ButtonRelease-1>", self.on_release)
        self.bind("<Button-3>", partial(on_right_click_button, self))

        self.config(relief="raised", bg=self.original_bg, fg=self.original_fg)
        print(f"[DEBUG] Button {self.cget('text')} initialized: path={self.path}, width={self.cget('width')}, height={self.cget('height')}, is_folder={self.is_folder}, is_shortcut={self.is_shortcut()}")

    # 폴더의 " 📁" 와 같은 방식으로 버튼 글자에 붙이는 바로가기 표시.
    # (자식 위젯으로 얹었다가 드롭 직후 안 그려지는 문제가 있어 되돌렸다)
    SHORTCUT_EXTS = (".lnk", ".url")
    SHORTCUT_MARK = "↲"  # ↲ 아래로 내려갔다 왼쪽으로 꺾이는 화살표

    def is_shortcut(self):
        """이 버튼이 바로가기(.lnk/.url)를 가리키는지."""
        return str(self.path).strip().lower().endswith(MyButton.SHORTCUT_EXTS)

    def refresh_shortcut_mark(self):
        """경로가 바뀌어 바로가기 여부가 달라졌을 때 표시를 다시 맞춘다."""
        text = self.cget("text").replace(" " + MyButton.SHORTCUT_MARK, "")
        if self.is_shortcut():
            text = f"{text} {MyButton.SHORTCUT_MARK}"
            self.config(compound="right")
        self.config(text=text)

    def json_file(self):
        """이 버튼이 사용하는 JSON 파일 경로.

        self.name 은 이 버튼의 고유 키(json_key)다. 드롭으로 만들 때는
        "이름_해시타임스탬프", JSON 에서 불러올 때는 파일명에서 복원한 값이 들어간다.
        예전에는 이동/선택 시 self.cget("text") 로 파일명을 새로 만드는 바람에
        같은 버튼이 JSON 파일 두 개를 갖게 되어, 다음 실행 때 두 개로 늘어났다.
        """
        return os.path.join(json_directory, f"{sanitize_filename(self.name)}_button.json")

    def save_state(self):
        """현재 위치·모양을 자기 JSON 파일에 저장."""
        button_data = {
            "name": self.cget("text").replace(" 📁", "").replace(" ↲", ""),
            "path": self.path,
            "x": self.winfo_x(),
            "y": self.winfo_y(),
            "color": self.original_bg,
            "password": self.password,
            "font_size": self.font_size,
            "fg": self.cget("fg"),
            "font_weight": self.font_weight,
            "width": self.cget("width"),
            "height": self.cget("height"),
            "original_width": self.cget("width"),
            "original_height": self.cget("height"),
            "hotkey": self.hotkey,
            "is_folder": self.is_folder,
            "selected": self.selected
        }
        try:
            with open(self.json_file(), "w", encoding="utf-8") as f:
                json.dump(button_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"[ERROR] 버튼 저장 실패: {self.cget('text')} - {e}")

    def destroy(self):
        if self.name in MyButton._created_buttons:
            del MyButton._created_buttons[self.name]
            print(f"[DEBUG] Button {self.cget('text')} removed from _created_buttons: path={self.path}")
        super().destroy()

    def on_press(self, event):
        global press_time, press_x, press_y, is_dragging, is_button_long_pressed, copied_format
        self.press_time = press_time = time.time() * 1000
        self.press_x = press_x = event.x
        self.press_y = press_y = event.y
        self.was_dragged = False
        self.is_dragging = False
        is_dragging = False
        is_button_long_pressed = False
        self.config(relief="sunken")
        self.drag_direction = None
        
        if copied_format and not self.is_copied_format:
            # 예전에는 미리보기만 하고 Enter 를 눌러야 적용됐다.
            # 클릭한 자리에서 바로 적용하고, 같은 버튼을 다시 누르면 되돌린다.
            self.toggle_selection()
            if self.selected:
                self.apply_copied_format()
            else:
                self.restore_format_backup()
        else:
            if self.selected and len(selected_items()) > 1:
                begin_group_drag(self)
        return "break"

    def on_drag(self, event):
        global is_dragging, is_button_long_pressed
        if copied_format:
            return
        current_time = time.time() * 1000
        throttle_interval = 10 if len(selected_buttons) <= 1 else 50
        if current_time - self.last_drag_time < throttle_interval:
            return
        self.last_drag_time = current_time

        print(f"[DEBUG] on_drag: time={current_time}, x={event.x}, y={event.y}, selected_count={len(selected_buttons)}")

        if (current_time - self.press_time) > 200 and not self.is_dragging:
            self.is_dragging = True
            self.was_dragged = True
            is_dragging = True
            is_button_long_pressed = True
            self.long_press_active = True

        if self.is_dragging:
            dx = event.x - self.press_x
            dy = event.y - self.press_y
            new_x = self.winfo_x() + dx
            new_y = self.winfo_y() + dy
            
            shift_pressed = event.state & 0x0001 != 0
            if shift_pressed:
                if self.drag_direction is None:
                    if abs(dx) > abs(dy):
                        self.drag_direction = "horizontal"
                        print(f"[DEBUG] Shift 눌림 - 방향 결정: 수평")
                    else:
                        self.drag_direction = "vertical"
                        print(f"[DEBUG] Shift 눌림 - 방향 결정: 수직")
                
                if self.drag_direction == "horizontal":
                    new_y = self.winfo_y()
                    print(f"[DEBUG] Shift 눌림 - 정수평 이동: dx={dx}, new_x={new_x}")
                elif self.drag_direction == "vertical":
                    new_x = self.winfo_x()
                    print(f"[DEBUG] Shift 눌림 - 정수직 이동: dy={dy}, new_y={new_y}")
            
            max_w, max_h = content_bounds()
            new_x = max(0, min(new_x, max_w - self.winfo_width()))
            new_y = max(0, min(new_y, max_h - self.winfo_height()))

            if self.selected and len(selected_items()) > 1:
                move_group(self, new_x, new_y)   # 선택된 라벨도 함께 이동
                root.update_idletasks()
            else:
                self.place(x=new_x, y=new_y)

    def on_release(self, event):
        global is_dragging, is_button_long_pressed
        if self._lock:
            return
        self._lock = True
        try:
            ctrl_pressed = event.state & 0x0004 != 0
            current_time = time.time() * 1000
            time_diff = current_time - self.press_time

            if ctrl_pressed and not copied_format:
                self.toggle_selection()
            elif time_diff > 2000 and not self.long_press_active and not copied_format:
                self.long_press_active = True
                is_button_long_pressed = True
                self.config(relief="raised")
                on_right_click_button(self, event)
            elif not self.was_dragged and not self.long_press_active and not copied_format:
                self.animate_click()
                if self.custom_command and not self.is_copied_format:
                    self.custom_command()

            self.is_dragging = False
            is_dragging = False
            is_button_long_pressed = False
            self.long_press_active = False
            self.drag_direction = None

            if self.was_dragged and not copied_format:
                self.config(relief="raised" if not self.selected else "sunken", bg=self.original_bg)
                # 예전에는 끌던 버튼 하나만 저장해서, 여러 개를 옮겨도
                # 다음 실행 때 나머지가 원래 자리로 돌아갔다.
                if self.selected and len(selected_items()) > 1:
                    save_group_positions()
                else:
                    self.save_state()
                print(f"[DEBUG] 버튼 이동 후 저장됨: {self.cget('text')}, 새 위치: ({self.winfo_x()}, {self.winfo_y()})")
        finally:
            self._lock = False
            self.was_dragged = False

    def toggle_selection(self):
        global selected_buttons
        if self.selected:
            self.selected = False
            selected_buttons.remove(self)
            self.config(relief="raised", bg=self.original_bg, fg=self.original_fg)
            print(f"[DEBUG] Button {self.cget('text')} deselected: width={self.cget('width')}, height={self.cget('height')}")
        else:
            self.selected = True
            selected_buttons.append(self)
            self.config(relief="sunken", bg="#ADD8E6" if not copied_format else "#FFA500")
            # 서식 복사 중에는 애니메이션의 after() 콜백이 방금 적용한 서식을
            # 100~300ms 뒤에 덮어써버린다(= "눌러도 변화가 없다"의 원인).
            if not copied_format:
                self.animate_selection()
            print(f"[DEBUG] Button {self.cget('text')} selected: width={self.cget('width')}, height={self.cget('height')}")
        json_file_path = self.json_file()
        with open(json_file_path, "w", encoding="utf-8") as f:
            json.dump({
                "name": self.cget("text").replace(" 📁", "").replace(" ↲", ""),
                "path": self.path,
                "x": self.winfo_x(),
                "y": self.winfo_y(),
                "color": self.original_bg,
                "password": self.password,
                "font_size": self.font_size,
                "fg": self.cget("fg"),
                "font_weight": self.font_weight,
                "width": self.cget("width"),
                "height": self.cget("height"),
                "hotkey": self.hotkey,
                "is_folder": self.is_folder,
                "selected": self.selected
            }, f, ensure_ascii=False, indent=4)
        print(f"[DEBUG] 버튼 상태 저장됨: {json_file_path}")

    def animate_selection(self):
        original_width = self['width']
        original_height = self['height']
        print(f"[DEBUG] Animate selection start for {self.cget('text')}: original_width={original_width}, original_height={original_height}")
        self.config(width=original_width + 4, height=original_height + 2, bg="#FFFF00")
        self.update_idletasks()
        self.after(100, lambda: self.config(width=original_width + 2, height=original_height + 1, bg="#ADD8E6" if not copied_format else "#FFA500"))
        self.after(200, lambda: self.config(width=original_width, height=original_height, bg="#ADD8E6" if not copied_format else "#FFA500"))
        self.after(300, lambda: [
            self.config(width=original_width, height=original_height),
            self.update_idletasks(),
            print(f"[DEBUG] Animate selection forced reset for {self.cget('text')}: width={self.cget('width')}, height={self.cget('height')}")
        ])

    def animate_click(self):
        if self.is_copied_format:
            print(f"[DEBUG] 복사된 버튼 {self.cget('text')} 클릭 실행 차단")
            return
        original_width = self['width']
        original_height = self['height']
        print(f"[DEBUG] Animate click start for {self.cget('text')}: original_width={original_width}, original_height={original_height}")
        self.config(width=original_width + 4, height=original_height + 2, bg="#FF5555", relief="sunken")
        self.update_idletasks()
        self.after(50, lambda: self.config(width=original_width + 2, height=original_height + 1, bg="#FF9999"))
        self.after(100, lambda: [
            self.config(width=original_width, height=original_height, bg=self.original_bg, relief="raised"),
            self.update_idletasks(),
            print(f"[DEBUG] Animate click end for {self.cget('text')}: width={self.cget('width')}, height={self.cget('height')}")
        ])

    def apply_copied_format(self):
        """복사한 서식을 이 버튼에 실제로 적용하고 JSON 에 저장한다.

        선택 상태는 색이 아니라 눌린 모양(relief="sunken")으로 표시한다.
        색으로 표시하면 방금 적용한 배경색이 가려져 아무 변화도 없어 보인다.
        """
        global copied_format
        if not copied_format:
            return
        # Esc / 다시 누르기로 되돌릴 수 있도록 적용 전 모습을 한 번만 보관한다.
        if not getattr(self, "_format_backup", None):
            self._format_backup = {
                "bg": self.original_bg,
                "fg": self.original_fg,
                "font_weight": self.font_weight,
                "width": self.original_width,
                "height": self.original_height
            }
        self.original_bg = copied_format["bg"]
        self.original_fg = copied_format["fg"]
        self.font_weight = copied_format["font_weight"]
        self.original_width = copied_format["width"]
        self.original_height = copied_format["height"]
        self.config(
            bg=self.original_bg,
            fg=self.original_fg,
            font=("Arial", self.font_size, self.font_weight),
            width=self.original_width,
            height=self.original_height,
            relief="sunken"
        )
        self.save_state()
        print(f"[DEBUG] 서식 적용: {self.cget('text')} <- {copied_format}")

    def restore_format_backup(self):
        """서식 적용 전 모습으로 되돌린다(같은 버튼 재클릭 / Esc)."""
        backup = getattr(self, "_format_backup", None)
        if not backup:
            return
        self.original_bg = backup["bg"]
        self.original_fg = backup["fg"]
        self.font_weight = backup["font_weight"]
        self.original_width = backup["width"]
        self.original_height = backup["height"]
        self.config(
            bg=self.original_bg,
            fg=self.original_fg,
            font=("Arial", self.font_size, self.font_weight),
            width=self.original_width,
            height=self.original_height,
            relief="raised"
        )
        self._format_backup = None
        self.save_state()
        print(f"[DEBUG] 서식 되돌림: {self.cget('text')}")
            
            
            
            
            
            
            
            
            
            
# 키보드로 선택된 항목 이동 함수
def move_selected_items(event):
    global selected_buttons, selected_labels
    print(f"[DEBUG] 선택된 버튼: {len(selected_buttons)}, 선택된 라벨: {len(selected_labels)}")
    print(f"[DEBUG] selected_labels: {[label.cget('text') for label in selected_labels]}")
    
    if not selected_buttons and not selected_labels:
        print("[DEBUG] 이동할 항목이 선택되지 않음")
        return

    move_distance = 10
    direction = {
        "Up": (0, -move_distance),
        "Down": (0, move_distance),
        "Left": (-move_distance, 0),
        "Right": (move_distance, 0)
    }.get(event.keysym)

    if direction is None:
        print(f"[DEBUG] 유효하지 않은 키: {event.keysym}")
        return

    dx, dy = direction

    max_w, max_h = content_bounds()

    # 선택된 버튼 이동
    for button in selected_buttons:
        new_x = button.winfo_x() + dx
        new_y = button.winfo_y() + dy
        new_x = max(0, min(new_x, max_w - button.winfo_width()))
        new_y = max(0, min(new_y, max_h - button.winfo_height()))
        button.place(x=new_x, y=new_y)
        # save_button_state 는 "이름" 으로 파일명을 만들어 별도 JSON 이 하나 더 생겼다.
        # 버튼 자신의 고유 키 파일에 저장해야 다음 실행 때 중복되지 않는다.
        button.save_state()
        print(f"[DEBUG] 버튼 이동: {button.cget('text')} -> ({new_x}, {new_y})")

    # 선택된 라벨 이동
    for label in selected_labels:
        new_x = label.winfo_x() + dx
        new_y = label.winfo_y() + dy
        new_x = max(0, min(new_x, max_w - label.winfo_width()))
        new_y = max(0, min(new_y, max_h - label.winfo_height()))
        label.place(x=new_x, y=new_y)
        # font_size와 font_style이 없을 경우 기본값 사용
        font_size = getattr(label, "font_size", 12)
        font_style = getattr(label, "font_style", "normal")
        save_label_state(label.cget("text"), new_x, new_y, label.cget("fg"), 
                         label.cget("font").split()[0], font_size, font_style,
                         uid=label_uid(label))
        print(f"[DEBUG] 라벨 이동: {label.cget('text')} -> ({new_x}, {new_y})")
        

        
  
    
 
    
        
        
        
# ── 선택 항목(버튼+라벨) 공통 처리 ────────────────────────────────────────────
# 원래 selected_buttons 와 selected_labels 가 완전히 분리돼 있어서
# 둘을 섞어 선택하거나 함께 옮기는 것이 불가능했다. 아래 헬퍼로 하나처럼 다룬다.

def selected_items():
    """현재 선택된 버튼 + 라벨."""
    return list(selected_buttons) + list(selected_labels)


def content_bounds():
    """위젯을 놓을 수 있는 영역 크기. 스크롤되므로 창보다 넓을 수 있다."""
    try:
        return (max(_scrollregion_cache[0], canvas.winfo_width(), 1),
                max(_scrollregion_cache[1], canvas.winfo_height(), 1))
    except Exception:
        return max(root.winfo_width(), 1), max(root.winfo_height(), 1)


def begin_group_drag(anchor):
    """드래그 시작 시 선택된 모든 항목의 기준점 대비 상대 위치를 기록."""
    base_x, base_y = anchor.winfo_x(), anchor.winfo_y()
    for item in selected_items():
        item.relative_x = item.winfo_x() - base_x
        item.relative_y = item.winfo_y() - base_y


def move_group(anchor, new_x, new_y):
    """기준 항목을 옮기면서 선택된 나머지(버튼·라벨)도 같은 간격으로 옮긴다."""
    max_w, max_h = content_bounds()
    for item in selected_items():
        ix = new_x + getattr(item, "relative_x", 0)
        iy = new_y + getattr(item, "relative_y", 0)
        ix = max(0, min(int(ix), max_w - item.winfo_width()))
        iy = max(0, min(int(iy), max_h - item.winfo_height()))
        item.place(x=ix, y=iy)


def select_all_items(event=None):
    """Ctrl+A : 모든 버튼과 라벨을 선택한다."""
    # 입력창에 포커스가 있으면 텍스트 전체 선택이 우선이므로 건드리지 않는다.
    try:
        focused = root.focus_get()
        if isinstance(focused, (tk.Entry, tk.Text, ttk.Entry)):
            return
    except Exception:
        pass

    selected_buttons.clear()
    selected_labels.clear()
    for widget in content_frame.winfo_children():
        # 그리드 배경(Canvas)·선택 사각형(Frame) 같은 보조 위젯은 대상이 아니다
        if isinstance(widget, MyButton):
            widget.selected = True
            widget.config(relief="raised", bg="#ADD8E6")
            selected_buttons.append(widget)
        elif isinstance(widget, tk.Label):
            widget.selected = True
            widget.config(relief="raised", bg="#ADD8E6")
            selected_labels.append(widget)

    print(f"[DEBUG] 전체 선택(Ctrl+A): 버튼 {len(selected_buttons)}개, "
          f"라벨 {len(selected_labels)}개")
    return "break"


def save_group_positions():
    """함께 옮긴 항목들의 위치를 각자의 JSON 에 저장."""
    for item in selected_items():
        try:
            if isinstance(item, MyButton):
                item.save_state()
            else:
                parts = item.cget("font").split(" ")
                save_label_state(
                    item.cget("text"), item.winfo_x(), item.winfo_y(), item.cget("fg"),
                    parts[0], int(parts[1]),
                    parts[2] if len(parts) > 2 else "normal",
                    uid=label_uid(item))
        except Exception as e:
            print(f"[DEBUG] 항목 위치 저장 중 오류: {e}")


# 드래그 선택 전역 변수
drag_start_x = 0
drag_start_y = 0
selection_rect = None

class RubberBand:
    """드래그 선택 사각형.

    button_frame 이 캔버스를 완전히 덮기 때문에 canvas.create_rectangle 로 그린
    도형은 화면에 보이지 않는다. 그래서 얇은 Frame 4개로 테두리를 직접 그린다.
    좌표는 button_frame 기준(= 캔버스 월드 좌표)이다.
    """

    def __init__(self, parent, color="blue", thickness=2):
        self.parent = parent
        self.color = color
        self.t = thickness
        self.lines = []
        for _ in range(4):
            line = tk.Frame(parent, bg=color)
            line._ignore_in_scrollregion = True
            self.lines.append(line)

    def show(self, x1, y1, x2, y2):
        x1, x2 = sorted((int(x1), int(x2)))
        y1, y2 = sorted((int(y1), int(y2)))
        w, h = max(x2 - x1, 1), max(y2 - y1, 1)
        t = self.t
        specs = [
            (x1, y1, w, t),              # 위
            (x1, max(y2 - t, y1), w, t), # 아래
            (x1, y1, t, h),              # 왼쪽
            (max(x2 - t, x1), y1, t, h), # 오른쪽
        ]
        for frame, (x, y, ww, hh) in zip(self.lines, specs):
            frame.place(x=x, y=y, width=ww, height=hh)
            frame.lift()

    def hide(self):
        for frame in self.lines:
            frame.place_forget()


_rubber_band = None

def get_rubber_band():
    """button_frame 이 만들어진 뒤에 생성해야 하므로 지연 생성한다."""
    global _rubber_band
    if _rubber_band is None:
        _rubber_band = RubberBand(content_frame)
    return _rubber_band


def start_drag_selection(event):
    global drag_start_x, drag_start_y, selection_rect
    if copied_format:  # 서식 복사 중에는 드래그 선택 비활성화
        return
    drag_start_x = canvas.canvasx(event.x)
    drag_start_y = canvas.canvasy(event.y)
    selection_rect = get_rubber_band()
    selection_rect.show(drag_start_x, drag_start_y, drag_start_x, drag_start_y)
    print(f"[DEBUG] 드래그 시작: ({drag_start_x}, {drag_start_y})")

def update_drag_selection(event):
    global selection_rect
    if copied_format or not selection_rect:
        return
    current_x = canvas.canvasx(event.x)
    current_y = canvas.canvasy(event.y)
    selection_rect.show(drag_start_x, drag_start_y, current_x, current_y)

def end_drag_selection(event):
    global selection_rect, selected_buttons, selected_labels, drag_target_type
    if copied_format or not selection_rect:
        return
    current_x = canvas.canvasx(event.x)
    current_y = canvas.canvasy(event.y)
    selection_rect.hide()
    
    # 선택 영역 계산
    x1, y1 = min(drag_start_x, current_x), min(drag_start_y, current_y)
    x2, y2 = max(drag_start_x, current_x), max(drag_start_y, current_y)
    
    print(f"[DEBUG] 선택 영역 (캔버스 좌표): ({x1}, {y1}) - ({x2}, {y2})")
    
    # 위젯이 button_frame(캔버스 좌표 0,0) 안에 있으므로
    # widget.winfo_x()/y() 가 이미 캔버스 좌표다. 별도 오프셋 보정이 필요 없다.
    
    # 임시 리스트로 선택된 위젯 저장
    temp_buttons = []
    temp_labels = []
    
    # 버튼과 라벨 확인
    for widget in content_frame.winfo_children():
        if isinstance(widget, MyButton) or isinstance(widget, tk.Label):
            w_x = widget.winfo_x()
            w_y = widget.winfo_y()
            w_w = widget.winfo_width()
            w_h = widget.winfo_height()
            
            # 교차 여부 확인
            intersects = (x1 < w_x + w_w and x2 > w_x and y1 < w_y + w_h and y2 > w_y)
            
            print(f"[DEBUG] 위젯: {widget.cget('text')}, "
                  f"캔버스 위치: ({w_x}, {w_y}), 크기: ({w_w}, {w_h}), "
                  f"교차: {intersects}, 타입: {type(widget).__name__}")
            
            if intersects:
                if isinstance(widget, MyButton):
                    temp_buttons.append(widget)
                elif isinstance(widget, tk.Label):
                    temp_labels.append(widget)
    
    # 예전에는 버튼과 라벨이 함께 걸리면 선택을 통째로 취소했다.
    # 이제는 섞여 있어도 그대로 선택한다(함께 옮기기 위해).
    # Ctrl 을 누른 채 드래그하면 기존 선택에 "더한다"(해제하지 않는다).
    ctrl_pressed = bool(event.state & 0x0004)
    if not ctrl_pressed:
        for widget in content_frame.winfo_children():
            if isinstance(widget, MyButton) and widget.selected:
                widget.selected = False
                widget.config(relief="flat", bg=widget.original_bg)
            elif isinstance(widget, tk.Label) and hasattr(widget, 'selected') and widget.selected:
                widget.selected = False
                widget.config(relief="flat", bg="white")
        selected_buttons.clear()
        selected_labels.clear()

    # 선택 적용 (이미 선택된 것은 건너뛴다)
    for widget in temp_buttons:
        if widget not in selected_buttons:
            widget.selected = True
            widget.config(relief="raised", bg="#ADD8E6")
            selected_buttons.append(widget)
    for widget in temp_labels:
        if widget not in selected_labels:
            widget.selected = True
            widget.config(relief="raised", bg="#ADD8E6")
            selected_labels.append(widget)
    print(f"[DEBUG] 드래그 선택 완료 (Ctrl={ctrl_pressed}): "
          f"버튼 {len(temp_buttons)}개, 라벨 {len(temp_labels)}개 걸림")
    
    selection_rect = None
    drag_target_type = None
    print(f"[DEBUG] 선택된 버튼: {[btn.cget('text') for btn in selected_buttons]}")
    print(f"[DEBUG] 선택된 라벨: {[lbl.cget('text') for lbl in selected_labels]}")
    
    
    
    
    

        
        
# NOTE: 아래 clear_selection_with_paste 는 현재 어떤 이벤트에도 바인딩되어 있지 않다.
# 예전에는 clear_selection 이라는 같은 이름으로 정의되어 있었으나,
# 바로 아래에서 동일한 이름이 다시 정의되면서 항상 덮어써져 실행되지 않는 상태였다.
# (라벨 붙여넣기는 우클릭 메뉴의 paste_label 로 동작한다.)
def clear_selection_with_paste(event):
    global selected_buttons, selected_labels, copied_label_data
    
    widget = event.widget
    print(f"[DEBUG] clear_selection_with_paste 호출됨: widget={widget}, x={event.x}, y={event.y}")

    # 캔버스 또는 root에서 클릭 시 붙여넣기 및 선택 해제 처리
    if widget in (canvas, root) and copied_label_data:
        # 라벨 이름은 중복되어도 되므로(uid 로 구분) 복사한 이름을 그대로 쓴다.
        new_name = copied_label_data["text"]

        # 라벨 붙여넣기
        new_label = MyLabel(
            content_frame,
            text=new_name,
            fg=copied_label_data["color"],  # 로그에서 'color' 사용
            font=(copied_label_data["font_name"], copied_label_data["font_size"], copied_label_data["font_style"]),
            bg="white"
        )
        new_label.place(x=event.x, y=event.y)
        new_label.bind("<ButtonPress-1>", partial(on_label_press, new_label))
        new_label.bind("<B1-Motion>", partial(on_label_drag, new_label))
        new_label.bind("<ButtonRelease-1>", partial(on_label_release, new_label))
        new_label.bind("<Button-3>", partial(on_right_click_label, new_label))
        save_label_state(
            new_label.cget("text"), event.x, event.y, new_label.cget("fg"),
            new_label.cget("font").split()[0], int(new_label.cget("font").split()[1]),
            new_label.cget("font").split()[2] if len(new_label.cget("font").split()) > 2 else "normal",
            uid=label_uid(new_label)
        )
        print(f"[DEBUG] 라벨 붙여넣기 완료: {new_name} at ({event.x}, {event.y})")
        copied_label_data = None  # 붙여넣기 후 초기화

    # 선택 해제 로직
    if widget == canvas:  # 캔버스에서만 선택 해제
        for btn in selected_buttons:
            btn.selected = False
            btn.config(relief="flat", bg=btn.original_bg)
        for lbl in selected_labels:
            lbl.selected = False
            lbl.config(relief="flat", bg="white")
        selected_buttons.clear()
        selected_labels.clear()
        print(f"[DEBUG] 모든 선택 해제됨")

        
        

def clear_selection(event=None):
    global selected_buttons, selected_labels
    if event:
        ctrl_pressed = event.state & 0x0004 != 0
        print(f"[DEBUG] clear_selection 호출됨: ctrl_pressed={ctrl_pressed}, event={event}")
        if ctrl_pressed:
            print(f"[DEBUG] Ctrl 눌림: 선택 유지")
            return  # Ctrl+클릭 시 선택 해제하지 않음
    else:
        print(f"[DEBUG] clear_selection 호출됨: event=None")

    # 선택 해제 로직
    if selected_buttons:
        for button in selected_buttons[:]:
            button.selected = False
            button.config(relief="raised", bg=button.original_bg)
            selected_buttons.remove(button)
        print(f"[DEBUG] 버튼 선택 해제 완료: {len(selected_buttons)}개 남음")
    if selected_labels:
        for label in selected_labels[:]:
            label.selected = False
            label.config(relief="flat", bg="white")
            selected_labels.remove(label)
        print(f"[DEBUG] 라벨 선택 해제 완료: {len(selected_labels)}개 남음")

create_window = None

def on_right_click(event):
    print(f"[DEBUG] 우측 단일 클릭 감지: {time.time() * 1000}")  # 단일 클릭 디버깅

def on_right_double_click(event):
    global last_right_double_click_time, create_window, last_right_single_click_time
    current_time = time.time() * 1000  # 밀리초 단위로 현재 시간
    time_diff = current_time - last_right_double_click_time
    
    # 첫 클릭인지 확인 (초기값이 0일 경우)
    if last_right_double_click_time == 0:
        print(f"[DEBUG] 우측 더블 클릭 감지: {current_time}, 첫 클릭으로 간주")
        last_right_double_click_time = current_time
        return  # 첫 클릭은 무시하고 두 번째 클릭 대기
    
    # 500ms 이내에 두 번째 클릭이 발생했는지 확인
    if time_diff < 500:
        print(f"[DEBUG] 우측 더블 클릭 확인: {current_time}, 이전 클릭과의 간격: {time_diff:.2f}ms")
    else:
        print(f"[DEBUG] 우측 더블 클릭 감지: {current_time}, 이전 클릭과의 간격: {time_diff:.2f}ms (500ms 초과, 단일 클릭으로 간주)")
        last_right_double_click_time = current_time
        return  # 500ms 초과 시 단일 클릭으로 간주하고 GUI 표시 안 함

    last_right_double_click_time = current_time

    # 조건 확인 시작
    # 1. 선택된 라벨 편집 GUI가 있는지 확인
    if 'edit_window' in globals() and edit_window is not None and edit_window.winfo_exists() and selected_labels:
        print("[DEBUG] 선택된 라벨 편집 GUI가 열려 있어 실행 버튼 추가 GUI를 표시하지 않음")
        return

    # 2. 선택된 버튼 편집 GUI가 있는지 확인
    if 'edit_window' in globals() and edit_window is not None and edit_window.winfo_exists() and selected_buttons:
        print("[DEBUG] 선택된 버튼 편집 GUI가 열려 있어 실행 버튼 추가 GUI를 표시하지 않음")
        return

    # 3. Ctrl 키가 눌러져 있는지 확인
    ctrl_pressed = event.state & 0x0004 != 0  # Ctrl 키 상태 비트 확인
    if ctrl_pressed:
        print("[DEBUG] Ctrl 키가 눌려 있어 실행 버튼 추가 GUI를 표시하지 않음")
        return

    # 4. 선택된 실행 버튼이 하나라도 있는지 확인 (Ctrl 키와 관계없이)
    if len(selected_buttons) > 0:
        print("[DEBUG] 선택된 실행 버튼이 하나 이상 있어 실행 버튼 추가 GUI를 표시하지 않음")
        return

    # 5. 선택된 라벨이 하나라도 있는지 확인 (Ctrl 키와 관계없이)
    if len(selected_labels) > 0:
        print("[DEBUG] 선택된 라벨이 하나 이상 있어 실행 버튼 추가 GUI를 표시하지 않음")
        return

    # 기존 버튼 생성 창이 존재하면 활성화
    if create_window and create_window.winfo_exists():
        create_window.lift()
        create_window.focus_set()
        print("[DEBUG] 기존 버튼 생성 창 활성화")
        return

    # 조건을 모두 통과하면 실행 버튼 추가 GUI 표시
    def create_button():
        name = name_input.get()
        path = path_input.get()
        path = auto_change_slash(path)
        existing_buttons = [btn.cget("text") for btn in content_frame.winfo_children() if isinstance(btn, tk.Button)]
        if name and name in existing_buttons:
            messagebox.showerror("오류", "중복된 버튼 이름입니다.")
            return
        color = "#4CAF50"
        button = MyButton(
            content_frame, path=path, password="", text=name, command=partial(run_file, path),
            width=15, height=1, bg=color, fg="white"
        )
        button.place(x=event.x, y=event.y)
        button.original_bg = color
        adjust_font_size(button, name if name else "", max_width=150)
        button.bind("<Button-3>", partial(on_right_click_button, button))
        new_button_data = {
            "name": name if name else "",
            "path": path,
            "x": event.x,
            "y": event.y,
            "color": color,
            "password": "",
            "font_size": 10,
            "fg": "white",
            "font_weight": "normal",
            "width": 15,
            "height": 1
        }
        update_json_file(name if name else "", json_directory, new_button_data)
        create_window.destroy()
        globals()['create_window'] = None
        print("[DEBUG] 새 버튼 추가됨")

    create_window = tk.Toplevel(root)
    create_window.title("실행 버튼 추가")
    create_window_width = 300
    create_window_height = 200  # 높이 증가
    create_window.configure(bg="white")
    create_window.attributes('-topmost', True)  # 항상 화면 맨 앞에 유지
    root.update_idletasks()
    pos_x = root.winfo_x() + (root.winfo_width() - create_window_width) // 2
    pos_y = root.winfo_y() + (root.winfo_height() - create_window_height) // 2
    create_window.geometry(f"{create_window_width}x{create_window_height}+{pos_x}+{pos_y}")

    tk.Label(create_window, text="버튼 이름:", bg="white").pack(pady=5)
    name_input = tk.Entry(create_window)
    name_input.pack(pady=5)

    tk.Label(create_window, text="파일 경로 (.py 포함):", bg="white").pack(pady=5)
    path_input = tk.Entry(create_window)
    path_input.pack(pady=5)
    tk.Button(create_window, text="파일 선택", command=lambda: path_input.insert(0, filedialog.askopenfilename(filetypes=[("All Files", "*.*"), ("Python Files", "*.py")]))).pack(pady=5)

    tk.Button(create_window, text="추가", command=create_button).pack(pady=5)
    name_input.focus()
    print("[DEBUG] 실행 버튼 추가 GUI 표시됨")

    create_window.protocol("WM_DELETE_WINDOW", lambda: [create_window.destroy(), globals().__setitem__('create_window', None)])
    
    
    
    


label_window = None
last_left_double_click_time = 0
last_right_double_click_time = 0

def on_double_click(event):
    global last_left_double_click_time, label_window
    current_time = time.time() * 1000
    if current_time - last_left_double_click_time < 500:
        print(f"[DEBUG] 더블 클릭 간격이 500ms 미만 ({current_time - last_left_double_click_time:.2f}ms), 무시됨")
        return
    last_left_double_click_time = current_time

    ctrl_pressed = event.state & 0x0004 != 0
    if ctrl_pressed:
        print("[DEBUG] Ctrl 키가 눌린 상태에서 더블 클릭 감지됨, 라벨 입력 무시")
        return

    # 기존 라벨 입력 창이 열려 있으면 닫기
    if label_window and label_window.winfo_exists():
        label_window.destroy()
        print("[DEBUG] 기존 라벨 입력 창 닫음")

    # 이벤트가 발생한 위젯 확인 및 좌표 보정
    widget = event.widget
    print(f"[DEBUG] 더블 클릭 발생: 위젯={widget}, 원본 좌표=({event.x}, {event.y})")

    # 어느 위젯에서 발생했든 화면 절대좌표를 button_frame 기준으로 환산한다.
    # (캔버스 오프셋과 현재 스크롤 위치가 자동으로 반영된다)
    x = event.x_root - content_frame.winfo_rootx()
    y = event.y_root - content_frame.winfo_rooty()

    # 음수만 막는다. 스크롤이 되므로 창 너비를 넘는 위치도 유효하다.
    x = max(0, x)
    y = max(0, y)
    print(f"[DEBUGмин] 보정된 좌표: x={x}, y={y}")

    def get_existing_labels():
        ui_labels = {child.cget("text") for child in content_frame.winfo_children() if isinstance(child, tk.Label)}
        json_labels = {data["text"] for data in load_labels()}  # JSON 파일에서 로드된 라벨 포함
        return ui_labels.union(json_labels)

    def create_label(text, x, y):
        label = MyLabel(content_frame, text=text, font=("Arial", 12, "normal"), bg="white", fg="black")
        label.place(x=x, y=y)
        label.bind("<ButtonPress-1>", partial(on_label_press, label))
        label.bind("<B1-Motion>", partial(on_label_drag, label))
        label.bind("<ButtonRelease-1>", partial(on_label_release, label))
        label.bind("<Button-3>", partial(on_right_click_label, label))
        label.executed = False
        save_label_state(text, x, y, "black", "Arial", 12, "normal", uid=label_uid(label))
        print(f"[DEBUG] 새 라벨 생성: {text} at ({x}, {y})")
        return label

    def update_label(label_text):
        if not label_text:
            print("[DEBUG] 빈 텍스트 입력, 라벨 생성 취소")
            return
        # 같은 이름의 라벨을 여러 개 만들 수 있다(각각 uid 로 구분된다).
        create_label(label_text, x, y)

    def open_calendar(label=None):
        def set_reservation():
            try:
                selected_date = cal.get_date()
                raw_time = f"{hour_entry.get()}:{minute_entry.get()}"
                exec_path = path_entry.get().strip()
                label_text = label_entry.get().strip()

                # 시간 형식 검증
                selected_time = raw_time.replace(":", "")
                datetime.strptime(selected_date + selected_time, "%Y%m%d%H%M")
                hour, minute = map(int, raw_time.split(":"))
                if not (0 <= hour <= 23 and 0 <= minute <= 59):
                    raise ValueError("시간은 00:00 ~ 23:59 사이여야 합니다.")

                formatted_time = f"{selected_date} {raw_time.replace(':', '')}"
                if exec_path:
                    reservation_text = f"실행 {formatted_time} {exec_path}"
                elif label_text:
                    reservation_text = f"알림 {formatted_time} {label_text}"
                else:
                    messagebox.showwarning("경고", "라벨 텍스트 또는 실행 경로를 입력하세요.")
                    return

                # 라벨 이름은 중복되어도 된다(uid 로 구분한다).
                if label:  # 기존 라벨 수정
                    label.config(text=reservation_text)
                    parts = label.cget("font").split(" ")
                    save_label_state(reservation_text, label.winfo_x(), label.winfo_y(),
                                     label.cget("fg"), parts[0], int(parts[1]),
                                     parts[2] if len(parts) > 2 else "normal",
                                     uid=label_uid(label))
                    print(f"[DEBUG] 라벨 수정됨: {reservation_text}")
                else:  # 새 라벨 생성
                    create_label(reservation_text, x, y)
                global is_tracking
                if not is_tracking:
                    is_tracking = True
                    check_meeting_time()
                cal_window.destroy()
            except ValueError as e:
                messagebox.showerror("오류", f"잘못된 시간 형식: {e}")
            except Exception as e:
                print(f"[DEBUG] 예약 설정 중 오류: {e}")
                messagebox.showerror("오류", f"예약 설정 중 문제 발생: {e}")

        cal_window = tk.Toplevel(root)
        cal_window.title("예약 설정")
        cal_window.geometry("400x500")
        cal_window.configure(bg="white")
        cal_window.attributes('-topmost', True)
        
        # 이벤트 위치에 따라 창 위치 조정
        cal_x = event.x_root - 200  # 창 중앙을 클릭 위치에 맞춤
        cal_y = event.y_root - 250
        cal_x = max(0, min(cal_x, root.winfo_screenwidth() - 400))
        cal_y = max(0, min(cal_y, root.winfo_screenheight() - 500))
        cal_window.geometry(f"400x500+{cal_x}+{cal_y}")
        print(f"[DEBUG] 캘린더 창 생성됨: 위치 ({cal_x}, {cal_y})")

        try:
            cal = TkCalendar(cal_window, selectmode="day", date_pattern="yyyymmdd")
            cal.pack(pady=10, fill="x")
            print("[DEBUG] 캘린더 위젯 추가됨")
        except Exception as e:
            print(f"[DEBUG] 캘린더 초기화 실패: {e}")
            messagebox.showerror("오류", "캘린더를 로드할 수 없습니다. tkcalendar 모듈을 확인하세요.")
            cal_window.destroy()
            return

        tk.Label(cal_window, text="시간 (HH:MM):", bg="white").pack(pady=5)
        time_frame = tk.Frame(cal_window, bg="white")
        time_frame.pack(pady=5)
        hour_entry = tk.Entry(time_frame, width=2)
        hour_entry.insert(0, "00")
        hour_entry.pack(side=tk.LEFT)
        tk.Label(time_frame, text=":", bg="white").pack(side=tk.LEFT)
        minute_entry = tk.Entry(time_frame, width=2)
        minute_entry.insert(0, "00")
        minute_entry.pack(side=tk.LEFT)
        print("[DEBUG] 시간 입력 필드 추가됨")

        tk.Label(cal_window, text="실행 경로 (파일/폴더/URL, 입력 시 실행 예약):", bg="white").pack(pady=5)
        path_entry = tk.Entry(cal_window, width=40)
        path_entry.pack(pady=5)
        tk.Button(cal_window, text="탐색", command=lambda: path_entry.insert(0, filedialog.askopenfilename() or filedialog.askdirectory())).pack(pady=5)
        print("[DEBUG] 경로 입력 필드 추가됨")

        tk.Label(cal_window, text="알림 텍스트 (선택):", bg="white").pack(pady=5)
        label_entry = tk.Entry(cal_window, width=40)
        if label:  # 기존 라벨 텍스트를 기본값으로 설정
            label_entry.insert(0, label.cget("text"))
        label_entry.pack(pady=5)

        tk.Button(cal_window, text="예약", command=set_reservation).pack(pady=10)
        print("[DEBUG] 예약 버튼 추가됨")

    # 캔버스에서 바로 텍스트 입력 (네모 박스 제거)
    # x, y 는 content_frame 기준 좌표다. 라벨과 같은 부모에 놓아야
    # 입력칸이 마우스로 클릭한 자리에 그대로 나타난다.
    # (root 에 놓으면 캔버스 오프셋만큼 위쪽으로 치우쳐 보인다)
    entry = tk.Entry(content_frame, bg="white", fg="black", font=("Arial", 12), borderwidth=0, highlightthickness=0)
    entry._ignore_in_scrollregion = True  # 임시 입력칸은 스크롤 영역 계산에서 제외
    entry.place(x=x, y=y)
    entry.lift()
    entry.focus_set()

    def finish_label(event=None):
        label_text = entry.get().strip()
        if label_text:
            update_label(label_text)
        entry.destroy()

    # Enter 키로 입력 완료
    entry.bind("<Return>", finish_label)
    # Ctrl 키와 함께 누르면 캘린더 열기 (Ctrl+Enter로 동작)
    entry.bind("<Control-Return>", lambda e: [entry.destroy(), open_calendar()])
    # 포커스 잃으면 자동 저장
    entry.bind("<FocusOut>", finish_label)
    # Esc 키로 취소
    entry.bind("<Escape>", lambda e: entry.destroy())
    
    
    
    
    
    

# 전역 변수 정의
press_x = 0
press_y = 0
is_dragging = False
selected_labels = []
selection_rect = None
last_drag_time = 0
drag_threshold = 0.03  # 30ms 유지

def on_label_press(label, event):
    global press_x, press_y, is_dragging
    press_x = event.x
    press_y = event.y
    is_dragging = False
    
    if selection_rect is not None:
        return "break"
    
    ctrl_pressed = event.state & 0x0004 != 0
    print(f"[DEBUG] on_label_press: 라벨={label.cget('text')}, Ctrl={ctrl_pressed}, press_x={press_x}, press_y={press_y}")
    
    if ctrl_pressed:
        toggle_label_selection(label)
        if selected_labels:
            label.focus_set()
            root.bind("<MouseWheel>", on_mouse_wheel)
            root.bind("<Button-4>", partial(on_mouse_wheel, delta=120))
            root.bind("<Button-5>", partial(on_mouse_wheel, delta=-120))
    else:
        if label not in selected_labels:
            for lbl in selected_labels[:]:
                toggle_label_selection(lbl)
            toggle_label_selection(label)
        
        if label in selected_labels and len(selected_items()) > 1:
            begin_group_drag(label)
    
    return "break"

def toggle_label_selection(label):
    global selected_labels
    if not hasattr(label, "selected"):
        label.selected = False
    if label.selected:
        label.selected = False
        selected_labels.remove(label)
        label.config(relief="flat")
        label.config(bg="#FF9999")
        label.after(50, lambda: label.config(bg="#FFFFFF", relief="flat"))
        print(f"[DEBUG] toggle_label_selection: 라벨={label.cget('text')} 선택 해제")
    else:
        label.selected = True
        selected_labels.append(label)
        label.config(relief="raised")
        label.config(bg="#FF5555", relief="raised")
        label.after(50, lambda: label.config(bg="#FF9999"))
        label.after(100, lambda: label.config(bg="#ADD8E6", relief="raised"))
        print(f"[DEBUG] toggle_label_selection: 라벨={label.cget('text')} 선택됨")

def on_label_drag(label, event):
    global is_dragging, last_drag_time
    if not is_dragging:
        is_dragging = True
    
    current_time = time.time()
    if current_time - last_drag_time < drag_threshold:
        return
    
    last_drag_time = current_time
    dx = event.x - press_x
    dy = event.y - press_y
    
    max_w, max_h = content_bounds()
    new_x = int(label.winfo_x() + dx)
    new_y = int(label.winfo_y() + dy)
    new_x = max(0, min(new_x, max_w - label.winfo_width()))
    new_y = max(0, min(new_y, max_h - label.winfo_height()))

    if hasattr(label, "selected") and label.selected and len(selected_items()) > 1:
        move_group(label, new_x, new_y)   # 선택된 버튼도 함께 이동
    else:
        label.place(x=new_x, y=new_y)
    
    root.update()  # 강제 화면 갱신

def on_label_release(label, event):
    global is_dragging
    if is_dragging:
        is_dragging = False
        for lbl in selected_labels:
            lbl.config(relief="raised")
        # 함께 옮긴 항목(버튼 포함)의 위치를 모두 저장
        if selected_items():
            save_group_positions()
        else:
            parts = label.cget("font").split(" ")
            save_label_state(label.cget("text"), label.winfo_x(), label.winfo_y(),
                             label.cget("fg"), parts[0], int(parts[1]),
                             parts[2] if len(parts) > 2 else "normal",
                             uid=label_uid(label))
    else:
        ctrl_pressed = event.state & 0x0004 != 0
        if not ctrl_pressed and hasattr(label, "selected") and label.selected:
            toggle_label_selection(label)
            if not selected_labels:
                root.unbind("<MouseWheel>")
                root.unbind("<Button-4>")
                root.unbind("<Button-5>")
                print(f"[DEBUG] on_label_release: MouseWheel 및 Button-4/5 이벤트 해제 (모두 선택 해제)")
    print(f"[DEBUG] on_label_release: 라벨={label.cget('text')}, 선택 상태={label.selected if hasattr(label, 'selected') else False}, 선택된 라벨 수={len(selected_labels)}")
    
    
    
    
    
def on_mouse_wheel(event, delta=None):
    if not selected_labels:
        print("[DEBUG] on_mouse_wheel: 선택된 라벨 없음, 무시됨")
        return
    
    delta = event.delta if delta is None else delta
    print(f"[DEBUG] on_mouse_wheel 호출: delta={delta}, 위치=({event.x}, {event.y}), 선택된 라벨 수={len(selected_labels)}")
    
    # 글자 크기 변화량 설정 (3으로 변경, 원하는 값으로 조정 가능)
    size_change = 5
    
    for label in selected_labels:
        # 현재 폰트 정보 가져오기
        current_font = label.cget("font").split()
        font_name = current_font[0]
        font_size = int(current_font[1])
        font_style = current_font[2] if len(current_font) > 2 else "normal"
        print(f"[DEBUG] 라벨={label.cget('text')}, 현재 폰트: 이름={font_name}, 크기={font_size}, 스타일={font_style}")

        # 글자 크기 조정
        old_size = font_size
        if delta > 0:
            font_size += size_change  # 변화량 증가
            font_size = min(72, font_size)  # 최대 크기 제한 (예: 50)
            print(f"[DEBUG] 휠 위로: 라벨={label.cget('text')}, 크기 {old_size} -> {font_size}")
        elif delta < 0:
            font_size -= size_change  # 변화량 증가
            font_size = max(6, font_size)  # 최소 크기 제한 (예: 6)
            print(f"[DEBUG] 휠 아래로: 라벨={label.cget('text')}, 크기 {old_size} -> {font_size}")

        # 새로운 폰트 적용
        label.config(font=(font_name, font_size, font_style), width=0, height=0)
        label.update_idletasks()
        print(f"[DEBUG] 라벨={label.cget('text')}, 폰트 적용 후: 크기={label.winfo_width()}x{label.winfo_height()}")

        # JSON 파일에 상태 저장
        save_label_state(
            label.cget("text"), 
            label.winfo_x(), 
            label.winfo_y(), 
            label.cget("fg"), 
            font_name, 
            font_size, 
            font_style,
            uid=label_uid(label)
        )
        print(f"[DEBUG] 라벨 상태 저장 완료: {label.cget('text')} -> 크기={font_size}")
        
    

def adjust_transparency(value):
    root.attributes('-alpha', float(value))

def add_version_info_to_title(root, version):
    global PROGRAM_VERSION
    PROGRAM_VERSION = version  # 전역 변수에 버전 저장
    
    def show_info(event):
        print("[DEBUG] 툴팁 표시 시작 - 이벤트 감지됨")
        # 기존 툴팁 제거 (중복 생성 방지)
        if hasattr(event.widget, 'tooltip') and event.widget.tooltip.winfo_exists():
            print("[DEBUG] 기존 툴팁 제거")
            event.widget.tooltip.destroy()
            del event.widget.tooltip
        
        tooltip = tk.Toplevel(root)
        tooltip.wm_overrideredirect(True)
        tooltip.wm_attributes("-topmost", True)  # 툴팁을 항상 최상단에

        # 툴팁 프레임 생성
        tooltip_frame = tk.Frame(tooltip, background="#D3D3D3")
        tooltip_frame.pack(padx=5, pady=5)

        # 기본 툴팁 텍스트
        tooltip_text = (
            "제작자: 금성환\n"
            "https://quicklaunchpadoriginal.tistory.com/1\n"
        )
        label = tk.Label(
            tooltip_frame,
            text=tooltip_text,
            background="#D3D3D3",
            relief="flat",
            justify='left',
            font=("Arial", 8, "bold"),
        )
        label.pack(anchor='w')

        # 유튜버 채널 강조
        youtube_label = tk.Label(
            tooltip_frame,
            text="유튜버 채널: 코드담다",
            background="#D3D3D3",
            foreground="#0066CC",  # 진한 파랑으로 강조
            relief="flat",
            justify='left',
            font=("Arial", 8, "bold"),
        )
        youtube_label.pack(anchor='w')

        # 추가 텍스트
        disclaimer_text = (
            "※ 본 프로그램의 사용으로 인한 모든 결과에 대한 책임은 \n"
            "사용자 본인에게 있으며 제작자는 어떠한 법적 책임도 지지 않습니다"
        )
        disclaimer_label = tk.Label(
            tooltip_frame,
            text=disclaimer_text,
            background="#D3D3D3",
            relief="flat",
            justify='left',
            font=("Arial", 8, "bold"),
        )
        disclaimer_label.pack(anchor='w')

        # 툴팁 크기 계산 (화면 밖에서 먼저 렌더링)
        tooltip.wm_geometry("+0+0")  # 임시 위치
        tooltip.update()  # 렌더링 강제
        tooltip_width = tooltip_frame.winfo_width()
        tooltip_height = tooltip_frame.winfo_height()
        print(f"[DEBUG] 툴팁 크기: width={tooltip_width}, height={tooltip_height}")
        
        # 창 크기
        root.update()  # 창 크기 동기화
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()
        root_width = root.winfo_width()
        root_height = root.winfo_height()
        print(f"[DEBUG] 창 크기: x={root_x}, y={root_y}, width={root_width}, height={root_height}")

        # 기본 위치: title_label 아래
        x = event.widget.winfo_rootx()
        y = event.widget.winfo_rooty() + event.widget.winfo_height()

        # 경계 조정
        if x + tooltip_width > root_x + root_width:  # 오른쪽 경계 초과
            x = root_x + root_width - tooltip_width - 5
        if x < root_x:  # 왼쪽 경계 초과
            x = root_x + 5
        if y + tooltip_height > root_y + root_height:  # 아래쪽 경계 초과
            y = event.widget.winfo_rooty() - tooltip_height - 5
        if y < root_y:  # 위쪽 경계 초과
            y = root_y + 5

        tooltip.wm_geometry(f"+{int(x)}+{int(y)}")  # 정수 좌표로 설정
        event.widget.tooltip = tooltip
        print(f"[DEBUG] 툴팁 표시 완료 - 위치: x={x}, y={y}, width={tooltip_width}, height={tooltip_height}")

    def hide_info(event):
        print("[DEBUG] 툴팁 숨김 시작")
        try:
            if hasattr(event.widget, 'tooltip') and event.widget.tooltip.winfo_exists():
                event.widget.tooltip.destroy()
                print("[DEBUG] 툴팁 제거 성공")
                del event.widget.tooltip
            else:
                print("[DEBUG] 제거할 툴팁 없음 또는 이미 제거됨")
        except Exception as e:
            print(f"[DEBUG] 툴팁 제거 중 오류: {e}")
        print("[DEBUG] 툴팁 숨김 완료")

    # control_frame에 title_label 배치
    title_label = tk.Label(
        control_frame,
        text=f"ver {version}",
        fg="gray",  # 글자 색상 회색
        font=("Helvetica", 10, "bold"),  # title_label은 Helvetica 유지
        relief="raised",  # 입체적 테두리
        borderwidth=2,  # 테두리 두께
        padx=4,  # 크기 조정
        pady=2
    )
    title_label.pack(side=tk.RIGHT, padx=10, pady=5)  # 오른쪽 상단 배치

    # 다중 그림자 효과
    shadow1 = tk.Label(
        control_frame,
        text=f"ver {version}",
        fg="gray70",
        font=("Helvetica", 10, "bold"),
        relief="flat",
    )
    shadow1.place(in_=title_label, x=2, y=2)  # 첫 번째 그림자

    shadow2 = tk.Label(
        control_frame,
        text=f"ver {version}",
        fg="gray50",
        font=("Helvetica", 10, "bold"),
        relief="flat",
    )
    shadow2.place(in_=title_label, x=3, y=3)  # 두 번째 그림자

    # title_label을 최상단으로
    title_label.lift()
    
    # 이벤트 바인딩
    title_label.bind("<Enter>", show_info)
    title_label.bind("<Leave>", hide_info)
    
    # 그림자 레이블은 이벤트를 받지 않도록 설정
    shadow1.event_generate("<Enter>", when="tail")  # 이벤트 무시
    shadow1.event_generate("<Leave>", when="tail")
    shadow2.event_generate("<Enter>", when="tail")
    shadow2.event_generate("<Leave>", when="tail")
    
    # 추가: root의 <Leave> 이벤트로 툴팁 제거 (안정성 강화)
    def hide_tooltip_on_root_leave(event):
        print("[DEBUG] root에서 마우스 떠남 감지")
        if hasattr(title_label, 'tooltip') and title_label.tooltip.winfo_exists():
            title_label.tooltip.destroy()
            del title_label.tooltip
            print("[DEBUG] root 떠남으로 툴팁 제거 완료")

    root.bind("<Leave>", hide_tooltip_on_root_leave)

    print("[DEBUG] title_label 및 그림자 생성 및 배치 완료")
    root.update_idletasks()  # 즉시 렌더링
    print(f"[DEBUG] title_label 위치: x={title_label.winfo_x()}, y={title_label.winfo_y()}, "
          f"width={title_label.winfo_width()}, height={title_label.winfo_height()}")
    
    
    
    
    
    
    
    
    
    
    
    
    
    

def load_window_state():
    file_path = os.path.join(json_directory, "window_state.json")
    if os.path.exists(file_path):
        try:
            with open(file_path, "r") as file:
                data = json.load(file)
                geometry = data.get("geometry", "800x600")
                state = data.get("state", "normal")
                return geometry, state
        except (json.JSONDecodeError, FileNotFoundError, KeyError) as e:
            print(f"창 상태 로드 중 오류 발생: {e}. 기본값 사용.")
            return "800x600", "normal"
    print("창 상태 파일이 없음. 기본값 사용.")
    return "800x600", "normal"

def save_window_state():
    file_path = os.path.join(json_directory, "window_state.json")
    state = "zoomed" if root.state() == "zoomed" else "normal"
    # 최대화 상태가 아니면 현재 창 크기를 저장
    geometry = root.geometry() if state != "zoomed" else root.geometry()  # 최대화 시에도 현재 geometry 저장
    data = {"geometry": geometry, "state": state}
    try:
        with open(file_path, "w") as file:
            json.dump(data, file, indent=4)
        print(f"창 상태 저장됨: {data}")  # 디버깅용 출력
    except Exception as e:
        print(f"창 상태 저장 중 오류 발생: {e}")





def drop(event, x=None, y=None, processed_folders=None, existing_labels=None, total_items=0):
    if processed_folders is None:
        processed_folders = set()
    if existing_labels is None:
        existing_labels = set()

    files = root.tk.splitlist(event.data) if x is None else [event]
    drop_x = event.x_root - content_frame.winfo_rootx() if x is None else x
    drop_y = event.y_root - content_frame.winfo_rooty() if y is None else y
    # 상단 컨트롤 영역 위에 떨어뜨리면 음수가 되어 화면 밖에 배치된다
    drop_x = max(0, drop_x)
    drop_y = max(0, drop_y)
    spacing = 2
    current_y = drop_y

    def _make_json_key(file_path, base_name):
        """같은 파일/폴더를 여러 번 드롭해도 각각 고유한 JSON 파일을 갖도록
        경로 해시 + 타임스탬프로 키를 생성한다."""
        path_hash = hashlib.md5(file_path.encode()).hexdigest()[:6]
        ts = str(int(time.time() * 1000))[-6:]  # 밀리초 끝 6자리
        sanitized = sanitize_filename(base_name)
        return f"{sanitized}_{path_hash}{ts}"

    def _create_button_from_drop(file_path, display_name, drop_x, current_y, is_folder):
        """버튼 생성 + JSON 저장 공통 로직.
        MyButton.__init__에서 이벤트 바인딩이 자동으로 이루어지므로
        별도 bind 호출 불필요."""
        color = "#4CAF50"
        button = MyButton(
            content_frame,
            path=file_path,
            password="",
            text=display_name,
            command=partial(run_file, file_path),
            width=15,
            height=1,
            bg=color,
            fg="white",
            is_folder=is_folder
        )
        button.place(x=drop_x, y=current_y)
        root.update_idletasks()
        button.original_bg = color
        label_for_font = f"{display_name} 📁" if is_folder else display_name
        adjust_font_size(button, label_for_font, max_width=150)
        print(f"[DEBUG] {'폴더' if is_folder else '파일'} 버튼 생성됨: {display_name}, "
              f"위치: ({drop_x}, {current_y}), 높이: {button.winfo_height()}")

        # JSON 저장 — 키를 경로 해시 기반으로 만들어 이름 충돌 방지
        json_key = _make_json_key(file_path, display_name)
        json_path = os.path.join(json_directory, f"{json_key}_button.json")
        button_data = {
            "name": display_name,
            "path": file_path,
            "x": drop_x,
            "y": current_y,
            "color": color,
            "password": "",
            "font_size": 10,
            "fg": "white",
            "font_weight": "normal",
            "width": 15,
            "height": 1,
            "is_folder": is_folder
        }
        try:
            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(button_data, f, indent=4, ensure_ascii=False)
            print(f"[DEBUG] JSON 저장됨: {json_path}")
        except Exception as e:
            print(f"[ERROR] JSON 저장 실패: {json_path} - {e}")

        # button.name을 json_key와 일치시켜 나중에 삭제 시 JSON도 함께 지워지게 함
        button.name = json_key
        MyButton._created_buttons[json_key] = button

        return button

    for i, file_path in enumerate(files):
        file_path = auto_change_slash(file_path)
        # 확장자 제거한 순수 이름 (원본 이름 그대로 유지)
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        # 바로가기는 파일명 끝에 "- 바로가기" 가 붙는다. 버튼 이름에서는 뺀다.
        if file_path.strip().lower().endswith(MyButton.SHORTCUT_EXTS):
            base_name = clean_shortcut_name(base_name)
        print(f"[DEBUG] 드롭된 항목: {file_path}, 기본 이름: {base_name}, "
              f"현재 y: {current_y}, 총 항목: {total_items}")

        if os.path.isdir(file_path):
            if file_path in processed_folders:
                print(f"[DEBUG] 이미 처리된 폴더 건너뜀: {file_path}")
                continue
            processed_folders.add(file_path)

            button = _create_button_from_drop(file_path, base_name, drop_x, current_y, is_folder=True)
            button_height = button.winfo_height()
            existing_labels.add(base_name)
            total_items += 1

        elif os.path.isfile(file_path):
            button = _create_button_from_drop(file_path, base_name, drop_x, current_y, is_folder=False)
            button_height = button.winfo_height()
            existing_labels.add(base_name)
            total_items += 1

        else:
            print(f"[DEBUG] 알 수 없는 항목, 건너뜀: {file_path}")
            continue

        # 단축키 할당
        if 1 <= total_items <= 9:
            shortcut_key = f"<Control-{total_items}>"
        elif 10 <= total_items <= 35:
            shortcut_key = f"<Control-{chr(total_items + 87)}>"
        else:
            current_y += button_height + spacing
            continue

        try:
            root.bind(shortcut_key, partial(run_file, file_path))
            print(f"[DEBUG] 단축키 {shortcut_key}가 {base_name} 버튼에 할당됨.")
        except Exception as e:
            print(f"[ERROR] 단축키 할당 오류: {e}")

        current_y += button_height + spacing

    return current_y






            
            

def on_closing():
    for widget in content_frame.winfo_children():
        if isinstance(widget, MyButton):
            # button.name 을 그대로 파일명으로 쓰면 폴더 버튼의 " 📁" 가 파일명에 들어가
            # "이름 📁_button.json" 과 "이름_button.json" 두 파일이 생겨 버튼이 2개로 늘어났다.
            # json_file() 이 항상 정규화된 경로를 돌려준다.
            clean_text = widget.cget("text").replace(" 📁", "").replace(" ↲", "")
            button_data = {
                "name": clean_text,
                "path": widget.path,
                "x": widget.winfo_x(),
                "y": widget.winfo_y(),
                "color": widget.original_bg,
                "password": widget.password,
                "font_size": widget.font_size,
                "fg": widget.cget("fg"),
                "font_weight": widget.font_weight,
                "width": widget.cget("width"),
                "height": widget.cget("height"),
                "original_width": widget.original_width,
                "original_height": widget.original_height,
                "is_folder": widget.is_folder,
                "selected": widget.selected,
                "hotkey": getattr(widget, "hotkey", None)
            }
            file_name = widget.json_file()
            try:
                with open(file_name, "w", encoding="utf-8") as f:
                    json.dump(button_data, f, indent=4, ensure_ascii=False)
                print(f"[DEBUG] 버튼 상태 저장됨: {file_name}")
            except OSError as e:
                print(f"[ERROR] 버튼 상태 저장 실패: {file_name} - {e}")
    save_window_state()
    root.destroy()
    
    
    
    
def show_button_info_and_set_hotkey(event):
    # 선택된 버튼 찾기
    selected_button = None
    for widget in content_frame.winfo_children():
        if isinstance(widget, MyButton) and widget.selected:
            selected_button = widget
            break
    
    if not selected_button:
        messagebox.showinfo("알림", "선택된 버튼이 없습니다.", parent=root)
        return

    # 버튼 정보 창 생성
    button_info_window = tk.Toplevel(root)
    button_info_window.title("버튼 정보 및 핫키 설정")
    button_info_window.geometry("500x200")
    button_info_window.configure(bg="#ffffff")
    button_info_window.attributes('-topmost', True)
    button_info_window.resizable(False, False)

    # 상단 제목
    title_frame = tk.Frame(button_info_window, bg="#34495e")
    title_frame.pack(fill="x")
    tk.Label(
        title_frame, 
        text="버튼 정보", 
        bg="#34495e", 
        fg="white", 
        font=("Arial", 14, "bold"), 
        pady=10
    ).pack()

    # 버튼 이름과 링크 표시
    info_frame = tk.Frame(button_info_window, bg="#ffffff")
    info_frame.pack(pady=10, padx=15, fill="x")

    tk.Label(info_frame, text="버튼 이름:", bg="#ffffff", fg="#2c3e50", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
    tk.Label(info_frame, text=selected_button.cget("text"), bg="#ffffff", fg="#2c3e50", font=("Arial", 10)).grid(row=0, column=1, sticky="w", pady=5)

    tk.Label(info_frame, text="링크:", bg="#ffffff", fg="#2c3e50", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
    link_label = tk.Label(info_frame, text=selected_button.path, bg="#ffffff", fg="#2980b9", font=("Arial", 10), cursor="hand2")
    link_label.grid(row=1, column=1, sticky="w", pady=5)
    link_label.bind("<Button-1>", lambda e: run_file(selected_button.path))

    # 핫키 설정 프레임
    hotkey_frame = tk.Frame(button_info_window, bg="#ffffff")
    hotkey_frame.pack(pady=10, padx=15, fill="x")

    tk.Label(hotkey_frame, text="핫키 설정:", bg="#ffffff", fg="#2c3e50", font=("Arial", 10, "bold")).pack(side="left")
    hotkey_entry = tk.Entry(hotkey_frame, width=20)
    hotkey_entry.pack(side="left", padx=5)
    hotkey_entry.insert(0, selected_button.hotkey if selected_button.hotkey else "None")

    # 핫키 입력 감지
    def capture_hotkey(event):
        key = event.keysym
        modifier = ""
        if event.state & 0x0004:  # Ctrl 키
            modifier = "Control-"
        elif event.state & 0x0008:  # Alt 키
            modifier = "Alt-"
        
        if key in ["Control_L", "Control_R", "Alt_L", "Alt_R"]:  # 단독 제어 키 무시
            return
        
        hotkey = f"<{modifier}{key}>"
        hotkey_entry.delete(0, tk.END)
        hotkey_entry.insert(0, hotkey)

    hotkey_entry.bind("<KeyPress>", capture_hotkey)

    # 핫키 저장 함수
    def save_hotkey():
        new_hotkey = hotkey_entry.get()
        if new_hotkey == "None" or not new_hotkey:
            if selected_button.hotkey:
                root.unbind(selected_button.hotkey)
            selected_button.hotkey = None
        else:
            # 기존 핫키 해제
            if selected_button.hotkey:
                root.unbind(selected_button.hotkey)
            # 새 핫키 바인딩
            try:
                root.bind(new_hotkey, lambda e: run_file(selected_button.path))
                selected_button.hotkey = new_hotkey
            except tk.TclError:
                messagebox.showerror("오류", "유효하지 않은 핫키입니다.", parent=button_info_window)
                return

        # JSON 업데이트
        button_data = {
            "hotkey": selected_button.hotkey,
            "name": selected_button.cget("text"),
            "path": selected_button.path,
            "x": selected_button.winfo_x(),
            "y": selected_button.winfo_y(),
            "color": selected_button.original_bg,
            "password": selected_button.password,
            "font_size": selected_button.font_size,
            "fg": selected_button.cget("fg"),
            "font_weight": selected_button.font_weight,
            "width": selected_button.cget("width"),
            "height": selected_button.cget("height"),
            "selected": selected_button.selected
        }
        update_json_file(selected_button.cget("text"), json_directory, button_data)
        messagebox.showinfo("성공", "핫키가 저장되었습니다.", parent=button_info_window)

    # 버튼 프레임
    button_frame = tk.Frame(button_info_window, bg="#ffffff")
    button_frame.pack(pady=10, fill="x")
    save_btn = tk.Button(
        button_frame, 
        text="저장", 
        command=save_hotkey, 
        bg="#2ecc71", 
        fg="white", 
        font=("Arial", 10, "bold"), 
        relief="flat", 
        padx=12, 
        pady=6
    )
    save_btn.pack(side="left", padx=5)
    close_btn = tk.Button(
        button_frame, 
        text="닫기", 
        command=button_info_window.destroy, 
        bg="#bdc3c7", 
        fg="#2c3e50", 
        font=("Arial", 10, "bold"), 
        relief="flat", 
        padx=12, 
        pady=6
    )
    close_btn.pack(side="left", padx=5)

    # 버튼 호버 효과
    save_btn.bind("<Enter>", lambda e: save_btn.config(bg="#27ae60"))
    save_btn.bind("<Leave>", lambda e: save_btn.config(bg="#2ecc71"))
    close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#95a5a6"))
    close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#bdc3c7"))    
    

def show_hotkey_info(event):
    # 고정 핫키 정의
    FIXED_HOTKEYS = ["<Control-c>", "<Control-C>", "<Control-g>", "<Control-G>"]
    
    # 핫키 수집 (고정 핫키 제외)
    hotkeys = []
    for widget in content_frame.winfo_children():
        if (isinstance(widget, MyButton) and 
            hasattr(widget, "hotkey") and 
            widget.hotkey and 
            widget.hotkey != "None" and 
            widget.hotkey not in FIXED_HOTKEYS):
            hotkeys.append((widget.cget('text'), widget.hotkey, widget.path))
    
    # 고정 핫키 항목 정의
    fixed_hotkey_entries = [
        ("캘린더 토글", "<Control-c>", "달력 표시/숨김 토글"),
        ("그리드 토글", "<Control-g>", "그리드 표시/숨김 토글"),
        ("전체 선택", "<Control-a>", "모든 버튼과 라벨 선택")
    ]
    
    if not hotkeys and not fixed_hotkey_entries:
        messagebox.showinfo("핫키 정보", "설정된 핫키가 없습니다.")
        return

    # 핫키를 Alt와 Ctrl로 분류
    alt_hotkeys = [hk for hk in hotkeys if hk[1].lower().startswith("<alt-")]
    ctrl_hotkeys = [hk for hk in hotkeys if hk[1].lower().startswith("<control-")]

    # 핫키 정보 창 생성
    hotkey_window = tk.Toplevel(root)
    hotkey_window.title("핫키 정보")
    hotkey_window.geometry("900x450")
    hotkey_window.configure(bg="#ffffff")
    hotkey_window.attributes('-topmost', True)
    hotkey_window.resizable(True, True)

    # 상단 제목
    title_frame = tk.Frame(hotkey_window, bg="#34495e")
    title_frame.pack(fill="x")
    tk.Label(
        title_frame, 
        text="핫키 정보", 
        bg="#34495e", 
        fg="white", 
        font=("Arial", 14, "bold"), 
        pady=10
    ).pack()

    # 메인 프레임 (2열 레이아웃)
    main_frame = tk.Frame(hotkey_window, bg="#ffffff")
    main_frame.pack(fill="both", expand=True, padx=15, pady=10)

    # 체크박스 상태를 저장할 딕셔너리
    check_vars = {}

    # Alt 핫키 열 (좌측)
    alt_frame = tk.Frame(main_frame, bg="#ffffff")
    alt_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
    alt_canvas = tk.Canvas(alt_frame, bg="#ffffff", highlightthickness=0)
    alt_scrollbar = ttk.Scrollbar(alt_frame, orient="vertical", command=alt_canvas.yview)
    alt_scrollable_frame = tk.Frame(alt_canvas, bg="#ffffff")

    alt_scrollable_frame.bind("<Configure>", lambda e: alt_canvas.configure(scrollregion=alt_canvas.bbox("all")))
    alt_canvas.create_window((0, 0), window=alt_scrollable_frame, anchor="nw")
    alt_canvas.configure(yscrollcommand=alt_scrollbar.set)

    alt_canvas.pack(side="left", fill="both", expand=True)
    alt_scrollbar.pack(side="right", fill="y")

    # Ctrl 핫키 열 (우측)
    ctrl_frame = tk.Frame(main_frame, bg="#ffffff")
    ctrl_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
    ctrl_canvas = tk.Canvas(ctrl_frame, bg="#ffffff", highlightthickness=0)
    ctrl_scrollbar = ttk.Scrollbar(ctrl_frame, orient="vertical", command=ctrl_canvas.yview)
    ctrl_scrollable_frame = tk.Frame(ctrl_canvas, bg="#ffffff")

    ctrl_scrollable_frame.bind("<Configure>", lambda e: ctrl_canvas.configure(scrollregion=ctrl_canvas.bbox("all")))
    ctrl_canvas.create_window((0, 0), window=ctrl_scrollable_frame, anchor="nw")
    ctrl_canvas.configure(yscrollcommand=ctrl_scrollbar.set)

    ctrl_canvas.pack(side="left", fill="both", expand=True)
    ctrl_scrollbar.pack(side="right", fill="y")

    # GUI 갱신 함수
    def update_hotkey_display():
        # 기존 항목 삭제
        for widget in alt_scrollable_frame.winfo_children():
            widget.destroy()
        for widget in ctrl_scrollable_frame.winfo_children():
            widget.destroy()
        check_vars.clear()

        # 최신 핫키 수집 (고정 핫키 제외)
        current_hotkeys = []
        for widget in content_frame.winfo_children():
            if (isinstance(widget, MyButton) and 
                hasattr(widget, "hotkey") and 
                widget.hotkey and 
                widget.hotkey != "None" and 
                widget.hotkey not in FIXED_HOTKEYS):
                current_hotkeys.append((widget.cget('text'), widget.hotkey, widget.path))
        
        alt_hotkeys = [hk for hk in current_hotkeys if hk[1].lower().startswith("<alt-")]
        ctrl_hotkeys = [hk for hk in current_hotkeys if hk[1].lower().startswith("<control-")]

        # Alt 핫키 표시
        if alt_hotkeys:
            tk.Label(alt_scrollable_frame, text="Alt 핫키", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 12, "bold"), pady=5).pack(fill="x")
            alt_header = tk.Frame(alt_scrollable_frame, bg="#ecf0f1")
            alt_header.pack(fill="x", pady=(0, 5))
            tk.Label(alt_header, text="선택", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=5).pack(side="left", padx=5)
            tk.Label(alt_header, text="라벨 이름", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=5)
            tk.Label(alt_header, text="핫키", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=5)
            tk.Label(alt_header, text="링크", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=20).pack(side="left", padx=5)

            for btn_name, hotkey, path in alt_hotkeys:
                var = tk.BooleanVar(value=False)
                check_vars[btn_name] = var
                row_frame = tk.Frame(alt_scrollable_frame, bg="white", bd=1, relief="flat")
                row_frame.pack(fill="x", pady=2, padx=5)

                tk.Checkbutton(row_frame, variable=var, bg="white").pack(side="left", padx=5)
                tk.Label(row_frame, text=btn_name, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                tk.Label(row_frame, text=hotkey, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                
                font_size = 10
                max_width = 150
                link_label = tk.Label(
                    row_frame, 
                    text=path, 
                    bg="white", 
                    fg="#2980b9", 
                    font=("Arial", font_size), 
                    anchor="w", 
                    cursor="hand2"
                )
                link_label.pack(side="left", padx=5, fill="x", expand=True)
                while link_label.winfo_reqwidth() > max_width and font_size > 6:
                    font_size -= 1
                    link_label.config(font=("Arial", font_size))
                link_label.bind("<Button-1>", lambda e, p=path: run_file(p))

        # Ctrl 핫키 표시
        if ctrl_hotkeys or fixed_hotkey_entries:
            tk.Label(ctrl_scrollable_frame, text="Ctrl 핫키", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 12, "bold"), pady=5).pack(fill="x")
            ctrl_header = tk.Frame(ctrl_scrollable_frame, bg="#ecf0f1")
            ctrl_header.pack(fill="x", pady=(0, 5))
            tk.Label(ctrl_header, text="선택", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=5).pack(side="left", padx=5)
            tk.Label(ctrl_header, text="라벨 이름", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=5)
            tk.Label(ctrl_header, text="핫키", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=15).pack(side="left", padx=5)
            tk.Label(ctrl_header, text="링크/설명", bg="#ecf0f1", fg="#2c3e50", font=("Arial", 10, "bold"), width=20).pack(side="left", padx=5)

            # 고정 핫키 표시 (체크박스 없이)
            for btn_name, hotkey, desc in fixed_hotkey_entries:
                row_frame = tk.Frame(ctrl_scrollable_frame, bg="white", bd=1, relief="flat")
                row_frame.pack(fill="x", pady=2, padx=5)

                # 체크박스 대신 빈 공간 추가
                tk.Label(row_frame, text="", bg="white", width=5).pack(side="left", padx=5)
                tk.Label(row_frame, text=btn_name, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                tk.Label(row_frame, text=hotkey, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                tk.Label(row_frame, text=desc, bg="white", fg="#2c3e50", font=("Arial", 10), anchor="w").pack(side="left", padx=5)

            # 일반 Ctrl 핫키 표시
            for btn_name, hotkey, path in ctrl_hotkeys:
                var = tk.BooleanVar(value=False)
                check_vars[btn_name] = var
                row_frame = tk.Frame(ctrl_scrollable_frame, bg="white", bd=1, relief="flat")
                row_frame.pack(fill="x", pady=2, padx=5)

                tk.Checkbutton(row_frame, variable=var, bg="white").pack(side="left", padx=5)
                tk.Label(row_frame, text=btn_name, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                tk.Label(row_frame, text=hotkey, bg="white", fg="#2c3e50", font=("Arial", 10), width=15, anchor="w").pack(side="left", padx=5)
                
                font_size = 10
                max_width = 150
                link_label = tk.Label(
                    row_frame, 
                    text=path, 
                    bg="white", 
                    fg="#2980b9", 
                    font=("Arial", font_size), 
                    anchor="w", 
                    cursor="hand2"
                )
                link_label.pack(side="left", padx=5, fill="x", expand=True)
                while link_label.winfo_reqwidth() > max_width and font_size > 6:
                    font_size -= 1
                    link_label.config(font=("Arial", font_size))
                link_label.bind("<Button-1>", lambda e, p=path: run_file(p))

        # 스크롤 영역 갱신
        alt_canvas.configure(scrollregion=alt_canvas.bbox("all"))
        ctrl_canvas.configure(scrollregion=ctrl_canvas.bbox("all"))

    # 초기 표시
    update_hotkey_display()

    # 열 가중치 설정
    main_frame.grid_columnconfigure(0, weight=1)
    main_frame.grid_columnconfigure(1, weight=1)
    main_frame.grid_rowconfigure(0, weight=1)

    # 삭제 버튼 기능
    def delete_selected_hotkeys():
        deleted = False
        for btn_name, var in list(check_vars.items()):  # 리스트로 변환하여 반복 중 수정 문제 방지
            if var.get():
                for widget in content_frame.winfo_children():
                    if (isinstance(widget, MyButton) and 
                        widget.cget("text") == btn_name and 
                        hasattr(widget, "hotkey") and 
                        widget.hotkey and 
                        widget.hotkey != "None"):
                        root.unbind(widget.hotkey)
                        widget.hotkey = None
                        button_data = {
                            "hotkey": None,
                            "name": widget.cget("text"),
                            "path": widget.path,
                            "x": widget.winfo_x(),
                            "y": widget.winfo_y(),
                            "color": widget.original_bg,
                            "password": widget.password,
                            "font_size": widget.font_size,
                            "fg": widget.cget("fg"),
                            "font_weight": widget.font_weight,
                            "width": widget.cget("width"),
                            "height": widget.cget("height"),
                            "selected": widget.selected
                        }
                        update_json_file(btn_name, json_directory, button_data)
                        deleted = True
        if deleted:
            update_hotkey_display()  # GUI 갱신
            messagebox.showinfo("성공", "선택된 핫키가 삭제되었습니다.", parent=hotkey_window)
        else:
            messagebox.showinfo("알림", "삭제할 핫키가 선택되지 않았습니다.", parent=hotkey_window)

    # 버튼 프레임
    button_frame = tk.Frame(hotkey_window, bg="#ffffff")
    button_frame.pack(pady=10, fill="x")
    delete_btn = tk.Button(
        button_frame, 
        text="선택된 핫키 삭제", 
        command=delete_selected_hotkeys, 
        bg="#e74c3c", 
        fg="white", 
        font=("Arial", 10, "bold"), 
        relief="flat", 
        padx=12, 
        pady=6
    )
    delete_btn.pack(side="left", padx=5)
    close_btn = tk.Button(
        button_frame, 
        text="닫기", 
        command=hotkey_window.destroy, 
        bg="#bdc3c7", 
        fg="#2c3e50", 
        font=("Arial", 10, "bold"), 
        relief="flat", 
        padx=12, 
        pady=6
    )
    close_btn.pack(side="left", padx=5)

    # 버튼 호버 효과
    delete_btn.bind("<Enter>", lambda e: delete_btn.config(bg="#c0392b"))
    delete_btn.bind("<Leave>", lambda e: delete_btn.config(bg="#e74c3c"))
    close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#95a5a6"))
    close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#bdc3c7"))

    # 창 크기 조절 시 캔버스 업데이트
    def resize_canvas(event):
        alt_canvas.configure(scrollregion=alt_canvas.bbox("all"))
        ctrl_canvas.configure(scrollregion=ctrl_canvas.bbox("all"))

    hotkey_window.bind("<Configure>", resize_canvas)
    
    
    
    
    
    
    
    
    
# F2 키 토글 함수 추가 (show_hotkey_info 아래)
# 삽입 위치: def show_hotkey_info(event): 아래
def toggle_cmd_visibility(event=None):
    global show_cmd
    show_cmd = not show_cmd
    state = "보임" if show_cmd else "숨김"
    print(f"[DEBUG] F2 키 눌림 - CMD 창 표시 상태: {state} (show_cmd={show_cmd})")
    return "break"  # 이벤트 전파 중단


class StdoutRedirector:
    """표준 출력을 Tkinter 텍스트 위젯으로 리디렉션"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        if self.text_widget and self.text_widget.winfo_exists():
            self.text_widget.insert(tk.END, message)
            self.text_widget.see(tk.END)  # 스크롤을 맨 아래로 유지
        else:
            # 디버깅 창이 없으면 원래 stdout으로 출력
            sys.__stdout__.write(message)

    def flush(self):
        pass  # flush는 Tkinter에서 필요 없음    
    

def start_tracking():
    global is_tracking
    is_tracking = True
    check_meeting_time()

is_tracking = False

def check_meeting_time():
    # 이전 라벨 텍스트를 저장하기 위한 딕셔너리
    label_previous_text = {}

    def execute_action(label, action_type, content):
        """스레드에서 실행할 작업 정의"""
        try:
            if action_type == "실행" and content:
                print(f"[DEBUG] 스레드에서 실행: {content}")
                run_file(content)
                # 실행 완료 후 상태만 업데이트 (초기화하지 않음)
                root.after(0, lambda: mark_executed(label, action_type))
            elif action_type == "알림":
                print(f"[DEBUG] 스레드에서 알림: {content}")
                # Tkinter는 메인 스레드에서만 UI 작업을 처리해야 하므로 after 사용
                root.after(0, lambda: messagebox.showinfo("알림", f"예약된 시간입니다: {content}"))
                # 알림 표시 후 상태 업데이트
                root.after(0, lambda: mark_executed(label, action_type))
        except Exception as e:
            print(f"[DEBUG] 스레드 실행 중 오류: {e}")
            # 오류 발생 시에도 상태 업데이트
            root.after(0, lambda: mark_executed(label, action_type))

    def mark_executed(label, action_type):
        """라벨 실행 상태를 영구적으로 표시"""
        try:
            label.executed = True
            if action_type == "알림" and hasattr(label, "alarm_shown"):
                label.alarm_shown = True
            print(f"[DEBUG] 라벨 실행 완료: {label.cget('text')}")
        except Exception as e:
            print(f"[DEBUG] 라벨 상태 업데이트 중 오류: {e}")

    def run_periodically():
        current_time = datetime.now().strftime("%Y%m%d %H%M")  # 현재 시간 (YYYYMMDD HHMM 형식)
        #print(f"[DEBUG] 현재 시간: {current_time}")
        labels = [child for child in content_frame.winfo_children() if isinstance(child, tk.Label)]
        
        # 실행할 작업 목록 수집
        tasks = []
        for label in labels:
            text_parts = label.cget("text").split(" ")
            if len(text_parts) >= 3 and (text_parts[0] in ["실행", "알림"]):
                scheduled_time = f"{text_parts[1]} {text_parts[2]}"  # 라벨에 설정된 시간
                current_text = label.cget("text")

                # 라벨 텍스트 변경 감지
                label_id = id(label)
                if label_id in label_previous_text and label_previous_text[label_id] != current_text:
                    label.executed = False
                    label.alarm_shown = False
                    print(f"[DEBUG] 라벨 변경 감지, 상태 초기화: {current_text}")
                label_previous_text[label_id] = current_text

                try:
                    sched_dt = datetime.strptime(scheduled_time, "%Y%m%d %H%M")
                    curr_dt = datetime.strptime(current_time, "%Y%m%d %H%M")
                    time_diff = (sched_dt - curr_dt).total_seconds() / 60  # 분 단위 차이

                    if time_diff < 0:
                        print(f"[DEBUG] 과거 시간 감지됨, 실행 건너뜀: {scheduled_time}")
                        continue
                    if time_diff > 0:
                        print(f"[DEBUG] 미래 시간 대기 중: {scheduled_time}, 시간 차이: {time_diff}분")
                        continue

                    if time_diff == 0:
                        print(f"[DEBUG] 실행 준비: {current_text}, executed={hasattr(label, 'executed') and label.executed}")
                        if not hasattr(label, "executed"):
                            label.executed = False
                        if not hasattr(label, "alarm_shown"):
                            label.alarm_shown = False
                        if not label.executed:
                            action_type = text_parts[0]
                            content = " ".join(text_parts[3:]) if len(text_parts) > 3 else ""
                            tasks.append((label, action_type, content))
                            # 여기서는 executed를 바로 설정하지 않음 (execute_action에서 처리)
                except ValueError as e:
                    print(f"[DEBUG] 시간 변환 오류: {e}")

        # 수집된 작업을 병렬로 실행
        threads = []
        for label, action_type, content in tasks:
            thread = threading.Thread(target=execute_action, args=(label, action_type, content))
            threads.append(thread)
            thread.start()

        # thread.join()을 제거하여 메인 스레드가 차단되지 않도록 함

        # 다음 체크 예약
        if any(isinstance(child, tk.Label) for child in content_frame.winfo_children()):
            root.after(5000, run_periodically)  # 1초마다 체크

    if not is_tracking:
        return
    run_periodically()
    
    
    
   
    


# 현재 사용자 이름 자동으로 가져오기
username = os.getenv('USERNAME')  # Windows 환경 변수에서 사용자 이름 가져옴
startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
exe_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
shortcut_path = os.path.join(startup_folder, f"{os.path.basename(sys.argv[0])}.lnk")  # .exe 포함 이름 사용
exe_path = os.path.abspath(sys.argv[0])  # 현재 EXE 파일의 절대 경로
json_directory_shared = "C:/json"
json_directory = os.path.join(json_directory_shared, exe_name)
if not os.path.exists(json_directory):
    os.makedirs(json_directory)  # 디렉토리 생성 보장
LAST_OPEN_FILE = os.path.join(json_directory, "last_youtube_open.json")
# GUI 설정
# GUI 설정
# 전역 변수
grid_on = False  # 그리드 상태를 추적하기 위한 전역 변수 추가

# GUI 설정
# GUI 설정
print("[DEBUG] GUI 설정 시작")
root = TkinterDnD.Tk()
root.title(exe_name)
root.configure(bg="white")
# 창을 초기에 숨김 처리하여 불완전한 상태로 표시되지 않도록 함
root.withdraw()
print("[DEBUG] root 창 초기화 완료")

# 창 크기와 상태를 먼저 설정해 초기 렌더링 안정화
geometry, state = load_window_state()
root.geometry(geometry)
if state == "zoomed":
    root.state("zoomed")
else:
    root.state("normal")
print(f"[DEBUG] 창 상태 로드됨: geometry={geometry}, state={state}")

control_frame = tk.Frame(root)
control_frame.pack(side=tk.TOP, fill=tk.X)
transparency_slider = tk.Scale(control_frame, from_=0.3, to=1.0, resolution=0.05, orient="horizontal", command=adjust_transparency)
transparency_slider.set(1.0)
transparency_slider.pack(side=tk.LEFT, padx=5)
print("[DEBUG] control_frame 및 투명도 슬라이더 설정 완료")

# 스크롤 가능한 캔버스 설정
canvas = tk.Canvas(root, bg="white", highlightthickness=0)

# 스크롤바 추가
# 주의: 스크롤바를 캔버스보다 "먼저" pack 해야 한다.
# 캔버스를 expand=True 로 먼저 pack 하면 남는 공간을 전부 차지해서
# 가로 스크롤바가 오른쪽 구석에 수십 px 크기로 찌그러진다.
scrollbar_x = tk.Scrollbar(root, orient="horizontal", command=canvas.xview)
scrollbar_x.pack(side=tk.BOTTOM, fill="x")
scrollbar_y = tk.Scrollbar(root, orient="vertical", command=canvas.yview)
scrollbar_y.pack(side=tk.RIGHT, fill="y")
canvas.pack(side=tk.LEFT, fill="both", expand=True)

canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

# 버튼과 라벨을 담을 프레임 (스크롤 대상). 캔버스 좌표 (0,0) 에 고정되므로
# 이 프레임 안 위젯의 winfo_x()/winfo_y() 가 곧 "스크롤 월드 좌표" 가 된다.
content_frame = tk.Frame(canvas, bg="white")
canvas.create_window((0, 0), window=content_frame, anchor="nw")

# 스크롤 영역 업데이트
_scrollregion_cache = [0, 0]

def update_scrollregion(event=None):
    """place() 로 배치된 자식은 부모의 요청 크기에 반영되지 않으므로
    (canvas.bbox("all") 가 항상 1x1 로 나온다) 자식 좌표에서 직접 계산한다."""
    try:
        width = height = 0
        for child in content_frame.winfo_children():
            # 그리드 배경·선택 사각형 같은 보조 위젯은 스크롤 영역 크기의 근거가 아니다.
            # (포함시키면 매번 40px 씩 커지며 무한히 늘어난다)
            if getattr(child, "_ignore_in_scrollregion", False):
                continue
            try:
                width = max(width, child.winfo_x() + child.winfo_reqwidth())
                height = max(height, child.winfo_y() + child.winfo_reqheight())
            except Exception:
                continue
        # 여백을 조금 두고, 최소한 보이는 영역만큼은 확보
        width = max(width + 40, canvas.winfo_width())
        height = max(height + 40, canvas.winfo_height())
        if [width, height] == _scrollregion_cache:
            return  # 변화 없으면 재설정하지 않음 (<Configure> 무한 루프 방지)
        _scrollregion_cache[:] = [width, height]
        content_frame.configure(width=width, height=height)
        canvas.configure(scrollregion=(0, 0, width, height))
        if grid_on:
            draw_grid()  # 넓어진 영역까지 그리드를 다시 그린다
    except Exception as e:
        print(f"[DEBUG] 스크롤 영역 갱신 중 오류: {e}")

# 버튼/라벨은 place() 로 추가·이동되므로 <Configure> 만으로는 감지되지 않는다.
# 주기적으로 스크롤 영역을 재계산한다.
def _scrollregion_poll():
    update_scrollregion()
    root.after(400, _scrollregion_poll)

print("[DEBUG] 스크롤 가능한 캔버스 설정 완료")

# 창 크기 변경 시 스크롤 영역 재계산
# (캔버스 크기는 pack 의 expand/fill 이 알아서 맞춰주므로 직접 지정하지 않는다.
#  직접 지정하면 스크롤바 자리를 덮어버린다.)
def resize_canvas(event):
    if event.widget is root:
        update_scrollregion()
root.bind("<Configure>", resize_canvas)

# 초기 렌더링 완료 후 창 표시
root.update_idletasks()  # 레이아웃 계산 완료
root.deiconify()  # 창 표시
root.focus_force()
print("[DEBUG] root에 포커스 강제 설정 완료")

# 버튼과 라벨 로딩을 비동기적으로 처리
def load_widgets_async():
    print(f"[DEBUG] 비동기 위젯 로딩 시작")
    # load_button_states()/load_label_states() 가 "기존 것을 지우고 새로 만드는" 함수다.
    # 예전에는 여기서 반환된 데이터로 버튼을 한 번 더 만들어서, 같은 버튼이
    # 똑같은 좌표에 두 개씩 겹쳐 생겼다(옮기면 밑에서 하나가 더 나오는 원인).
    load_button_states()
    load_label_states()
    print(f"[DEBUG] 비동기 위젯 로딩 완료")
    canvas.update_idletasks()  # 로딩 후 스크롤 영역 갱신

# 버튼과 라벨 로딩을 창 표시 후 약간 지연시켜 실행
root.after(100, load_widgets_async)










# 그리드 토글 함수
def toggle_grid(event=None):
    global grid_on
    grid_on = not grid_on
    if grid_on:
        draw_grid()
    else:
        clear_grid()
    print(f"[DEBUG] 그리드 상태: {grid_on}")

# 그리드 전용 배경 캔버스
# content_frame 이 메인 캔버스를 덮기 때문에 canvas.create_line 으로 그린 그리드는
# 화면에 보이지 않는다. 그래서 content_frame 안에 배경 캔버스를 따로 두고 거기에 그린다.
grid_canvas = tk.Canvas(content_frame, bg="white", highlightthickness=0, bd=0)
grid_canvas._ignore_in_scrollregion = True

# 그리드 그리기 함수
def draw_grid():
    width, height = _scrollregion_cache
    width = max(width, canvas.winfo_width())
    height = max(height, canvas.winfo_height())
    grid_canvas.delete("grid_line")
    grid_canvas.place(x=0, y=0, width=width, height=height)
    # Canvas.lower() 는 캔버스 "도형"을 내리는 tag_lower 라서 인자가 필요하다.
    # 위젯 자체의 쌓임 순서를 내리려면 tk.Misc.lower 를 직접 호출해야 한다.
    tk.Misc.lower(grid_canvas)  # 버튼/라벨 아래로 (형제 위젯 중 맨 뒤)
    for i in range(0, width, 10):
        grid_canvas.create_line(i, 0, i, height, tag="grid_line", fill="gray")
    for i in range(0, height, 10):
        grid_canvas.create_line(0, i, width, i, tag="grid_line", fill="gray")
    print(f"[DEBUG] 그리드 표시됨 ({width}x{height})")

# 그리드 지우기 함수
def clear_grid():
    grid_canvas.delete("grid_line")
    grid_canvas.place_forget()
    print("[DEBUG] 그리드 지워짐")

# 드래그 선택 이벤트 바인딩
canvas.bind("<ButtonPress-1>", start_drag_selection)
canvas.bind("<B1-Motion>", update_drag_selection)
canvas.bind("<ButtonRelease-1>", end_drag_selection)

# button_frame 이 캔버스를 덮으므로 빈 공간 클릭은 캔버스가 아니라 프레임이 받는다.
# 같은 핸들러를 프레임에도 연결하되, 좌표를 캔버스 기준으로 환산해서 넘긴다.
def _as_canvas_event(handler):
    def wrapper(event):
        event.x = event.x_root - canvas.winfo_rootx()
        event.y = event.y_root - canvas.winfo_rooty()
        return handler(event)
    return wrapper

content_frame.bind("<ButtonPress-1>", _as_canvas_event(start_drag_selection))
content_frame.bind("<B1-Motion>", _as_canvas_event(update_drag_selection))
content_frame.bind("<ButtonRelease-1>", _as_canvas_event(end_drag_selection))

# 그리드(Ctrl+G)를 켜면 grid_canvas 가 content_frame 을 덮어 마우스를 가로챈다.
# 같은 핸들러를 grid_canvas 에도 연결해야 그리드 상태에서도 드래그 선택이 된다.
grid_canvas.bind("<ButtonPress-1>", _as_canvas_event(start_drag_selection))
grid_canvas.bind("<B1-Motion>", _as_canvas_event(update_drag_selection))
grid_canvas.bind("<ButtonRelease-1>", _as_canvas_event(end_drag_selection))

add_version_info_to_title(root, "12.8")
print("[DEBUG] 버전 정보 추가 완료")

geometry, state = load_window_state()
root.geometry(geometry)
if state == "zoomed":
    root.state("zoomed")
else:
    root.state("normal")
print(f"[DEBUG] 창 상태 로드됨: geometry={geometry}, state={state}")

root.focus_force()
print("[DEBUG] root에 포커스 강제 설정 완료")

# 서식 복사 및 드래그 선택 전역 변수
copied_format = None
selected_buttons = []
drag_start_x = 0
drag_start_y = 0
selection_rect = None

# 버튼 및 라벨 로드
# 버튼·라벨 로딩은 아래 두 함수가 전부 담당한다(멱등).
# 여기서 반환값으로 위젯을 또 만들면 중복 생성되므로 호출만 한다.
load_button_states()
load_label_states()

# 서식 적용 및 취소 함수
def apply_format(event=None):
    """Enter: 서식 복사 모드를 끝낸다.

    서식은 버튼을 누르는 순간 이미 적용되므로(MyButton.apply_copied_format)
    여기서는 혹시 빠진 버튼만 마무리하고 모드를 해제한다.
    """
    global copied_format, selected_buttons
    if event and event.keysym == "Return" and copied_format:
        for button in list(selected_buttons):
            if not button.is_copied_format:
                button.apply_copied_format()
        clear_copied_format()

def cancel_format(event=None):
    """Esc: 이번 서식 복사로 바뀐 버튼들을 원래 모습으로 되돌리고 모드 해제."""
    global copied_format
    if event and event.keysym == "Escape" and copied_format:
        for button in content_frame.winfo_children():
            if isinstance(button, MyButton) and not button.is_copied_format:
                button.restore_format_backup()
        # 서식을 복사한 원본 버튼 복구(색·실행 명령)와 선택 해제는 여기서 함께 처리한다.
        clear_copied_format()

def clear_copied_format():
    global copied_format
    for button in content_frame.winfo_children():  # root 대신 button_frame
        if isinstance(button, MyButton):
            if button.is_copied_format:
                button.config(
                    bg=button.original_bg,
                    fg=button.original_fg,
                    font=("Arial", button.font_size, button.font_weight),
                    command=partial(run_file, button.path + ("|" + button.password if button.password else "")),
                    relief="raised"
                )
                button.is_copied_format = False
                button.custom_command = partial(run_file, button.path + ("|" + button.password if button.password else ""))
            elif button.selected:
                button.config(
                    fg=button.original_fg,
                    font=("Arial", button.font_size, button.font_weight),
                    relief="raised"
                )
            button.selected = False
            button._format_backup = None  # 적용 확정: 되돌리기 정보 폐기
    selected_buttons.clear()
    copied_format = None

# 드래그 선택 함수
drag_start_x = 0
drag_start_y = 0
selection_rect = None

# 이벤트 바인딩
print("[DEBUG] 이벤트 바인딩 시작")
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', drop)
root.bind("<Button-1>", clear_selection)
root.bind("<Double-Button-1>", on_double_click)
root.bind("<Double-Button-3>", on_right_double_click)
root.bind("<Delete>", delete_selected_items)
root.bind("<Control-g>", toggle_grid)  # Ctrl+G를 toggle_grid에 직접 연결
root.bind("<Control-G>", toggle_grid)  # 대문자 G도 동일하게 처리
root.bind("<Control-c>", edit_selected_buttons)
root.bind("<Control-C>", edit_selected_buttons)
root.bind("<F1>", show_hotkey_info)
root.bind("<Return>", apply_format)  # Enter 키로 서식 적용
root.bind("<Escape>", cancel_format)  # Esc 키로 서식 취소
root.bind("<F2>", lambda event: [toggle_cmd_visibility(event), toggle_debug_window(event)])
root.bind("<Button-3>", on_right_click_root)  # 빈 공간 우클릭 이벤트
root.bind("<Double-Button-1>", on_double_click)  # 왼쪽 더블 클릭 이벤트
root.bind("<Control-v>", on_paste_button)
root.bind("<Control-V>", on_paste_button)
root.bind("<Control-a>", select_all_items)  # 전체 선택 (버튼+라벨)
root.bind("<Control-A>", select_all_items)  # 대문자 A도 동일
root.protocol("WM_DELETE_WINDOW", on_closing)

# 방향키로 선택된 항목 이동
root.bind("<Up>", move_selected_items)
root.bind("<Down>", move_selected_items)
root.bind("<Left>", move_selected_items)
root.bind("<Right>", move_selected_items)

print("[DEBUG] 모든 이벤트 바인딩 완료")
bindings = root.bind()
print(f"[DEBUG] 현재 root에 바인딩된 이벤트: {bindings}")

# 시간 추적 시작
start_tracking()
# print("[DEBUG] 시간 추적 시작")

# 기존 코드의 마지막 부분
print("[DEBUG] 현재 root에 바인딩된 이벤트:")
for event in root.bind():
    print(f"  {event}: {root.bind(event)}")

# 응원 메시지 리스트
encouraging_messages = [
    "지금 숨이 막혀도, 곧 다시 바람이 불 거야.",
    "작은 씨앗 하나가 언젠가 숲을 만들 거야.",
    "여기 서 있는 것만으로도 이미 큰 힘이야.",
    "폭풍이 몰아쳐도 결국엔 햇살이 비칠 거야.",
    "흔들리는 마음 속에도 불꽃이 살아있어.",
    "오늘은 멈춰도, 내일은 다시 날아갈 거야.",
    "멀리 보이지 않아도, 길은 열려 있어.",
    "넘어진 자리에 꽃이 피어날 거야.",
    "쌓인 땀방울이 언젠가 빛을 낼 거야.",
    "지금 이 순간이 너를 새롭게 만들어가고 있어.",
    "어둠이 깊을수록 별빛이 더 선명해져.",
    "지친 발걸음에도 희망이 숨어 있어.",
    "한 번의 선택이 세상을 뒤흔들 수 있어.",
    "끝없는 하늘처럼 가능성은 열려 있어.",
    "쓰러져도 다시 피어나는 꽃잎이야.",
    "혼자 걷는 길에도 손 내밀 누군가가 있어.",
    "오늘의 눈물이 내일의 웃음을 키워줄 거야.",
    "빛은 언제나 어둠을 뚫고 나와.",
    "아직 손에 잡히지 않아도, 꿈은 가까워.",
    "부서진 조각들도 언젠가 그림이 될 거야。",
    "한 번의 용기가 운명을 바꿀 수 있어。",
    "가슴이 뛰는 곳으로 한 발짝 내디뎌봐。",
    "작은 떨림이 큰 파도를 일으킬 거야。",
    "꺼지지 않는 불씨가 마음 속에 있어。",
    "이 순간이 새로운 날을 열어줄 거야。",
    "어떤 벽 앞에서도 무너지지 않을 거야。",
    "끝나지 않은 이야기가 아직 남아 있어。",
    "매일 아침이 새로운 날개를 줄 거야。",
    "두려움이 밀려와도, 다시 날아오를 거야。",
    "걸음마다 세상이 조금씩 바뀌고 있어."
]

# 희망을 주는 배경색 리스트
hopeful_bg_colors = [
    "#FFD700", "#98FB98", "#FFA07A", "#DDA0DD", "#F0E68C",
    "#FFB6C1", "#FFDEAD", "#87CEEB", "#ADD8E6", "#E6E6FA"
]

# 보색 계산 함수
def get_complementary_color(hex_color):
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    return f"#{255-r:02x}{255-g:02x}{255-b:02x}"

# 랜덤 메시지 선택 함수
def get_random_message():
    return random.choice(encouraging_messages)

# 실행 이력 확인 함수
def check_execution_history():
    current_date = datetime.now().strftime("%Y-%m-%d")
    json_base_dir = "C:/json"  # C:/json 폴더 전체 검색
    
    print(f"[DEBUG] check_execution_history 시작, 검색 경로: {json_base_dir}")
    for root_dir, _, files in os.walk(json_base_dir):
        print(f"[DEBUG] 검색 중인 디렉토리: {root_dir}")
        for file in files:
            if file == "last_youtube_open.json":
                file_path = os.path.join(root_dir, file)
                print(f"[DEBUG] 발견된 파일: {file_path}")
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        last_open_date = data.get("last_open_date")
                        print(f"[DEBUG] 파일 내용: {data}, last_open_date: {last_open_date}")
                        if last_open_date == current_date:
                            print(f"[DEBUG] 오늘 실행 이력 발견: {file_path}")
                            return True
                except (json.JSONDecodeError, FileNotFoundError) as e:
                    print(f"[DEBUG] 파일 읽기 오류: {file_path}, 오류: {e}")
                    continue
                except Exception as e:
                    print(f"[DEBUG] 기타 오류: {file_path}, 오류: {e}")
                    continue
    print("[DEBUG] 오늘 실행 이력 없음")
    return False

# 메시지 표시 및 유튜브 열기 함수
def show_message_and_open_link():
    youtube_link = "https://www.youtube.com/?hl=ko&gl=KR&app=desktop"
    current_date = datetime.now().strftime("%Y-%m-%d")
    LAST_OPEN_FILE = os.path.join(json_directory, "last_youtube_open.json")
    
    print(f"[DEBUG] show_message_and_open_link 시작 - 현재 날짜: {current_date}")
    print(f"[DEBUG] LAST_OPEN_FILE 경로 설정: {LAST_OPEN_FILE}")
    print(f"[DEBUG] json_directory 확인: {json_directory}")

    # 실행 이력 확인
    try:
        print(f"[DEBUG] 실행 이력 확인 시작: check_execution_history 호출")
        if check_execution_history():
            print(f"[DEBUG] 오늘 ({current_date}) 이미 실행됨, 메시지 표시 및 링크 열기 건너뜀")
            return
        print(f"[DEBUG] 실행 이력 없음, 메시지 표시 및 링크 열기 진행")
    except Exception as e:
        print(f"[ERROR] 실행 이력 확인 중 오류: {e}")

    # 메시지 창 생성
    try:
        print(f"[DEBUG] 메시지 창 생성 시작")
        message_window = tk.Toplevel(root)
        message_window.overrideredirect(True)
        message_window.attributes("-alpha", 0.0)
        message_window.attributes("-topmost", True)
        print(f"[DEBUG] 메시지 창 속성 설정 완료: alpha=0.0, topmost=True")
    except Exception as e:
        print(f"[ERROR] 메시지 창 생성 실패: {e}")
        return

    # 배경색 및 글자색 설정
    try:
        random_bg_color = random.choice(hopeful_bg_colors)
        random_fg_color = get_complementary_color(random_bg_color)
        print(f"[DEBUG] 색상 선택 - 배경색: {random_bg_color}, 글자색: {random_fg_color}")
        
        frame = tk.Frame(message_window, bg=random_bg_color)
        frame.pack(expand=True, fill="both")
        print(f"[DEBUG] 프레임 생성 및 배치 완료")
    except Exception as e:
        print(f"[ERROR] 색상 설정 또는 프레임 생성 실패: {e}")
        message_window.destroy()
        return

    # 메시지 라벨 생성
    try:
        message = get_random_message()
        label = tk.Label(
            frame,
            text=message,
            font=("Helvetica", 22, "bold"),
            justify="center",
            bg=random_bg_color,
            fg=random_fg_color
        )
        label.pack(expand=True, pady=20)
        print(f"[DEBUG] 메시지 라벨 생성 완료 - 메시지: {message}")
    except Exception as e:
        print(f"[ERROR] 메시지 라벨 생성 실패: {e}")
        message_window.destroy()
        return

    # 창 크기 및 위치 설정
    try:
        message_window.update_idletasks()
        text_width = max(label.winfo_reqwidth() + 40, 300)
        text_height = max(label.winfo_reqheight() + 40, 100)
        print(f"[DEBUG] 메시지 창 크기 계산 - 너비: {text_width}, 높이: {text_height}")

        root.update_idletasks()
        root_x = root.winfo_x()
        root_y = root.winfo_y()
        root_width = root.winfo_width()
        root_height = root.winfo_height()
        message_x = root_x + (root_width - text_width) // 2
        message_y = root_y + (root_height - text_height) // 2
        message_window.geometry(f"{text_width}x{text_height}+{message_x}+{message_y}")
        print(f"[DEBUG] 메시지 창 위치 설정 완료 - x: {message_x}, y: {message_y}")
    except Exception as e:
        print(f"[ERROR] 메시지 창 크기 설정 실패: {e}")
        message_window.destroy()
        return

    # 파일 쓰기 함수
    def write_last_open_date():
        try:
            print(f"[DEBUG] last_youtube_open.json 쓰기 시작: {LAST_OPEN_FILE}")
            os.makedirs(json_directory, exist_ok=True)
            with open(LAST_OPEN_FILE, "w", encoding="utf-8") as f:
                json.dump({"last_open_date": current_date}, f, indent=4)
            print(f"[DEBUG] last_youtube_open.json 쓰기 성공: {LAST_OPEN_FILE}")
        except PermissionError as e:
            print(f"[ERROR] last_youtube_open.json 쓰기 실패 - 권한 오류: {LAST_OPEN_FILE}, 오류: {e}")
        except Exception as e:
            print(f"[ERROR] last_youtube_open.json 쓰기 실패: {LAST_OPEN_FILE}, 오류: {e}")

    # 페이드인 효과
    def fade_in(step=0):
        try:
            alpha = step / 20
            message_window.attributes("-alpha", alpha)
            print(f"[DEBUG] 페이드인 진행 - alpha: {alpha:.2f}, step: {step}")
            if step < 20:
                message_window.after(100, fade_in, step + 1)
            else:
                print(f"[DEBUG] 페이드인 완료, 페이드아웃 예약")
                message_window.after(1000, fade_out, 10)
        except Exception as e:
            print(f"[ERROR] 페이드인 효과 적용 실패: {e}")
            message_window.destroy()

    # 페이드아웃 효과
    def fade_out(step=10):
        try:
            alpha = step / 10
            message_window.attributes("-alpha", alpha)
            print(f"[DEBUG] 페이드아웃 진행 - alpha: {alpha:.2f}, step: {step}")
            if step > 0:
                message_window.after(100, fade_out, step - 1)
            else:
                print(f"[DEBUG] 페이드아웃 완료, 유튜브 링크 열기 시작: {youtube_link}")
                webbrowser.open(youtube_link)
                print(f"[DEBUG] 유튜브 링크 열기 완료")
                write_last_open_date()
                message_window.destroy()
                print(f"[DEBUG] 메시지 창 닫힘")
        except Exception as e:
            print(f"[ERROR] 페이드아웃 효과 적용 또는 유튜브 링크 열기 실패: {e}")
            message_window.destroy()

    print(f"[DEBUG] 메시지 창 표시 시작")
    fade_in(0)

# STARTUP 정리 함수
def cleanup_startup_folder():
    """STARTUP 폴더에서 더 이상 존재하지 않는 타겟을 가진 바로가기를 정리"""
    try:
        startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        if not os.path.exists(startup_folder):
            print(f"[DEBUG] STARTUP 폴더가 존재하지 않음: {startup_folder}")
            return

        for file in os.listdir(startup_folder):
            if file.endswith('.lnk'):
                shortcut_path = os.path.join(startup_folder, file)
                try:
                    shell = Dispatch('WScript.Shell')
                    shortcut = shell.CreateShortCut(shortcut_path)
                    target_path = os.path.abspath(shortcut.Targetpath)  # 절대 경로로 확인
                    if not os.path.exists(target_path):
                        os.remove(shortcut_path)
                        print(f"[DEBUG] 존재하지 않는 타겟의 바로가기 삭제됨: {shortcut_path}, 타겟: {target_path}")
                    else:
                        print(f"[DEBUG] 유효한 바로가기 확인됨: {shortcut_path}, 타겟: {target_path}")
                except Exception as e:
                    print(f"[DEBUG] 바로가기 처리 중 오류: {shortcut_path}, 오류: {e}")
    except Exception as e:
        print(f"[DEBUG] STARTUP 폴더 정리 중 오류: {e}")

# STARTUP에 추가 함수
def add_to_startup():
    """STARTUP 폴더에 현재 실행 파일의 바로가기를 추가하고 기존 정리"""
    try:
        # 먼저 STARTUP 폴더 정리
        cleanup_startup_folder()

        # 현재 실행 파일 경로와 이름
        current_exe_path = os.path.abspath(sys.argv[0])
        current_exe_name = os.path.basename(current_exe_path)

        # .py 파일은 STARTUP에 추가하지 않음
        if current_exe_name.lower().endswith('.py'):
            print(f"[DEBUG] .py 파일은 STARTUP에 추가되지 않음: {current_exe_name}")
            return

        # STARTUP 폴더 경로
        startup_folder = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        shortcut_path = os.path.join(startup_folder, f"{current_exe_name}.lnk")

        # 기존 바로가기 확인 및 삭제 로직
        if os.path.exists(shortcut_path):
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(shortcut_path)
            existing_target = os.path.abspath(shortcut.Targetpath)  # 절대 경로로 비교

            if existing_target != current_exe_path:
                # 타겟 경로가 다르면 기존 바로가기 삭제
                os.remove(shortcut_path)
                print(f"[DEBUG] 기존 바로가기 삭제됨 (타겟 불일치): {shortcut_path}, 기존 타겟: {existing_target}")
            else:
                # 타겟이 동일하면 유지
                print(f"[DEBUG] 이미 최신 바로가기 존재함: {shortcut_path}")
                return

        # 새 바로가기 생성
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = current_exe_path
        shortcut.WorkingDirectory = os.path.dirname(current_exe_path)
        shortcut.save()
        print(f"[DEBUG] STARTUP에 새 바로가기 추가됨: {shortcut_path}")

    except Exception as e:
        print(f"[DEBUG] STARTUP 추가 중 오류: {e}")

# 수정된 send_to_nateon 함수

        
        

# 메인 실행 블록
if __name__ == "__main__":
    # json_directory 동적 설정
    exe_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    json_directory = f"C:/json/{exe_name}"
    try:
        os.makedirs(json_directory, exist_ok=True)
    except PermissionError as e:
        print(f"[ERROR] json_directory 생성 실패 - 권한 오류: {json_directory}, 오류: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"[ERROR] json_directory 생성 실패: {json_directory}, 오류: {e}")
        sys.exit(1)

    # 네이트온 전송 진행
    # 네트워크 호출 3회(최대 20초)를 메인 스레드에서 하면 GUI가 뜨기 전에 멈추므로
    # 백그라운드 데몬 스레드에서 처리한다. (UI 작업이 전혀 없어 스레드 사용이 안전함)
    def _notify_access():
        try:
            username = os.getlogin() or "Unknown User"
            program_name = exe_name
            external_ip = "Unknown IP"
            try:
                external_ip = requests.get('https://api.ipify.org', timeout=5).text
            except requests.RequestException:
                pass

            location = "Unknown Location"
            try:
                ip_info = requests.get(f'https://ipinfo.io/{external_ip}/json', timeout=5).json()
                city = ip_info.get('city', 'Unknown City')
                region = ip_info.get('region', 'Unknown Region')
                country = ip_info.get('country', 'Unknown Country')
                location = f"{city}, {region}, {country}"
            except requests.RequestException:
                pass

            access_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            program_version = "12.3"
            message = (
                f"User: {username}, Program: {program_name}, "
                f"Location: {location}, Time: {access_time}, Version: {program_version}"
            )
            webhook_url = "https://teamroom.nate.com/api/webhook/be7e8ea5/C53VcbV4FeBhLDtCkw2EvH6t"
            requests.post(webhook_url, json={"text": message}, timeout=10)
        except Exception:
            pass

    threading.Thread(target=_notify_access, daemon=True).start()

    # STARTUP 정리 및 추가
    try:
        add_to_startup()
    except Exception as e:
        print(f"[WARNING] 스타트업 추가 중 오류: {e}")

    # 메시지 표시 및 유튜브 링크 열기
    try:
        show_message_and_open_link()
    except Exception as e:
        print(f"[WARNING] 메시지 표시 및 유튜브 링크 열기 중 오류: {e}")

    # 버튼 상태 로드
    try:
        load_button_states()
    except Exception as e:
        print(f"[ERROR] 버튼 상태 로드 중 오류: {e}")

    # 라벨 로드 (기존 라벨 정리 포함 — load_label_states 가 전부 처리)
    try:
        load_label_states()
    except Exception as e:
        print(f"[ERROR] 라벨 데이터 로드 중 오류: {e}")

    # 스크롤 영역 주기적 갱신 시작 (버튼/라벨이 place() 로 추가·이동되므로)
    try:
        _scrollregion_poll()
    except Exception as e:
        print(f"[WARNING] 스크롤 영역 갱신 시작 실패: {e}")

    # 메인 루프 실행
    try:
        root.mainloop()
    except Exception as e:
        print(f"[ERROR] Tkinter 메인 루프 실행 중 오류: {e}")