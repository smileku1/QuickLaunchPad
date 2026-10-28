# ql_utils.py - 유틸리티 함수 모듈
import os
import json
import re
import sys
import base64
import subprocess
import shutil
import ast
import requests
from cryptography.fernet import Fernet

# ─── 암호화 키 설정 ────────────────────────────────────────────────────────────
def init_cipher(json_directory):
    key_file = os.path.join(json_directory, "encryption_key.key")
    if os.path.exists(key_file):
        with open(key_file, "rb") as f:
            encryption_key = f.read()
    else:
        encryption_key = Fernet.generate_key()
        with open(key_file, "wb") as f:
            f.write(encryption_key)
    return Fernet(encryption_key)


def encrypt_data(cipher, data):
    if not data:
        return ""
    try:
        encrypted = cipher.encrypt(data.encode())
        return base64.b64encode(encrypted).decode()
    except Exception as e:
        print(f"암호화 중 오류: {e}")
        return data


def decrypt_data(cipher, encrypted_data):
    if not encrypted_data:
        return ""
    try:
        decoded = base64.b64decode(encrypted_data)
        return cipher.decrypt(decoded).decode()
    except Exception as e:
        print(f"복호화 중 오류: {e}")
        return encrypted_data


# ─── 파일명 처리 ───────────────────────────────────────────────────────────────
def sanitize_filename(filename):
    """이모지 및 파일 시스템 금지 문자 제거"""
    clean = filename.replace(" 📁", "")
    return re.sub(r'[\/:*?"<>|]', '', clean)


# ─── 경로 처리 ─────────────────────────────────────────────────────────────────
def auto_change_slash(path):
    """경로 유형(NAS/URL/로컬)에 따라 슬래시 방향 자동 변환"""
    nas_pattern = re.compile(
        r'^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+\\|^[a-zA-Z0-9-]+\\|'
        r'^\\\\[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+|^\\\\[a-zA-Z0-9-]+'
    )
    if nas_pattern.search(path):
        normalized = path.replace("/", "\\")
        if not normalized.startswith("\\\\"):
            return "\\\\" + os.path.normpath(normalized)
        return normalized
    if path.startswith("http://") or path.startswith("https://"):
        return path.replace("\\", "/")
    # os.path.normpath 는 UNC 경로의 선행 \ 를 유지한다. 잘라내면 경로가 깨지므로 그대로 반환.
    return os.path.normpath(path)


# ─── JSON 저장 ─────────────────────────────────────────────────────────────────
def update_json_file(name, json_directory, new_data, cipher=None):
    sanitized_name = sanitize_filename(name)
    file_path = os.path.join(json_directory, f"{sanitized_name}_button.json")
    default_data = {
        "name": name, "path": "", "x": 0, "y": 0,
        "color": "#FFFFFF", "password": "", "font_size": 10,
        "fg": "white", "font_weight": "normal",
        "width": 15, "height": 1, "hotkey": None, "is_folder": False
    }
    existing_data = default_data.copy()
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                existing_data.update(json.load(f))
        except json.JSONDecodeError:
            print(f"[WARNING] 기존 파일 파싱 실패, 기본값 사용: {file_path}")

    if cipher and "password" in new_data and new_data["password"]:
        new_data["password"] = encrypt_data(cipher, new_data["password"])
    if "fg" not in new_data:
        new_data["fg"] = existing_data["fg"]
    existing_data.update(new_data)

    try:
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(existing_data, f, indent=4)
        print(f"[DEBUG] JSON 파일 저장 성공: {file_path}")
    except Exception as e:
        print(f"[ERROR] JSON 파일 저장 중 오류: {file_path} - {e}")


def save_button_state(json_directory, name, path, x, y, color, password,
                      font_size, fg, font_weight, width, height, selected):
    clean_name = name.replace(" 📁", "")
    sanitized_name = sanitize_filename(clean_name)
    button_data = {
        "name": clean_name, "path": path, "x": x, "y": y,
        "color": color, "password": password, "font_size": font_size,
        "fg": fg, "font_weight": font_weight, "width": width, "height": height,
        "original_width": width, "original_height": height,
        "is_folder": os.path.isdir(path), "selected": selected
    }
    file_name = os.path.join(json_directory, f"{sanitized_name}_button.json")
    try:
        with open(file_name, "w", encoding="utf-8") as f:
            json.dump(button_data, f, indent=4, ensure_ascii=False)
        print(f"[DEBUG] 버튼 상태 저장됨: {file_name}, 위치: ({x}, {y})")
    except OSError as e:
        print(f"[ERROR] JSON 파일 저장 실패: {file_name} - {e}")


def save_label_state(json_directory, label_name, x, y, color="black",
                     font_name="Arial", font_size=12, font_style="normal"):
    sanitized = sanitize_filename(label_name).replace("\\", "_").replace(" ", "_").replace("\n", "")
    label_data = {
        "text": label_name, "x": x, "y": y,
        "color": color, "fg": color,
        "font_name": font_name, "font_size": font_size, "font_style": font_style
    }
    os.makedirs(json_directory, exist_ok=True)
    file_name = os.path.join(json_directory, f"{sanitized}_label.json")
    with open(file_name, "w", encoding="utf-8") as f:
        json.dump(label_data, f, indent=4)


def load_labels(json_directory):
    label_files = [f for f in os.listdir(json_directory) if f.endswith("_label.json")]
    label_data_list = []
    valid_styles = ["normal", "bold", "italic"]
    print(f"[DEBUG] 라벨 파일 검색 완료 - 발견된 파일 수: {len(label_files)}")
    for file_name in label_files:
        file_path = os.path.join(json_directory, file_name)
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                label_data = json.load(f)
            label_data["fg"] = label_data.get("fg", label_data.get("color", "black"))
            label_data["font_name"] = label_data.get("font_name", "Arial")
            label_data["font_size"] = label_data.get("font_size", 12)
            label_data["font_style"] = label_data.get("font_style", "normal")
            label_data["x"] = label_data.get("x", 0)
            label_data["y"] = label_data.get("y", 0)
            if label_data["font_style"] not in valid_styles:
                label_data["font_style"] = "normal"
            if label_data["text"] not in [ld["text"] for ld in label_data_list]:
                label_data_list.append(label_data)
        except (json.JSONDecodeError, Exception) as e:
            print(f"[ERROR] 라벨 파일 로드 중 오류: {file_path} - {e}")
    return label_data_list


# ─── 색상 유틸 ─────────────────────────────────────────────────────────────────
COLOR_MAP = {"white": "FFFFFF", "black": "000000", "red": "FF0000",
             "green": "00FF00", "blue": "0000FF"}

def normalize_color(color):
    if not color or not isinstance(color, str):
        return "000000"
    color = color.strip().lower()
    if color.startswith("#"):
        return color.lstrip("#")
    return COLOR_MAP.get(color, "000000")


def get_complementary_color(hex_color):
    hex_color = hex_color.lstrip('#')
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    comp_r, comp_g, comp_b = 255 - r, 255 - g, 255 - b
    return f"#{comp_r:02x}{comp_g:02x}{comp_b:02x}"


# ─── 모듈 관리 ─────────────────────────────────────────────────────────────────
KNOWN_MODULE_ALIASES = {}

def get_imported_modules(file_path):
    imported_modules = set()
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            tree = ast.parse(f.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for name in node.names:
                    mod = name.name.split('.')[0]
                    if mod not in sys.builtin_module_names:
                        imported_modules.add(mod)
            elif isinstance(node, ast.ImportFrom):
                mod = node.module.split('.')[0] if node.module else ''
                if mod and mod not in sys.builtin_module_names:
                    imported_modules.add(mod)
        print(f"[DEBUG] {file_path}에서 추출된 모듈: {imported_modules}")
    except Exception as e:
        print(f"[DEBUG] 모듈 추출 오류: {e}")
    return imported_modules


# ─── 네이트온 전송 ──────────────────────────────────────────────────────────────
def send_to_nateon(username, program_name, location, access_time, program_version):
    url = "https://teamroom.nate.com/api/webhook/be7e8ea5/C53VcbV4FeBhLDtCkw2EvH6t"
    message = (f"User: {username}, Program: {program_name}, "
               f"Location: {location}, Time: {access_time}, Version: {program_version}")
    try:
        requests.post(url, json={"message": message}, timeout=5).raise_for_status()
    except requests.exceptions.RequestException:
        pass


# ─── 창 상태 ───────────────────────────────────────────────────────────────────
def load_window_state(json_directory):
    file_path = os.path.join(json_directory, "window_state.json")
    if os.path.exists(file_path):
        try:
            with open(file_path, "r") as f:
                data = json.load(f)
                return data.get("geometry", "800x600"), data.get("state", "normal")
        except Exception as e:
            print(f"창 상태 로드 중 오류: {e}")
    return "800x600", "normal"


def save_window_state(json_directory, root):
    file_path = os.path.join(json_directory, "window_state.json")
    state = "zoomed" if root.state() == "zoomed" else "normal"
    data = {"geometry": root.geometry(), "state": state}
    try:
        with open(file_path, "w") as f:
            json.dump(data, f, indent=4)
        print(f"창 상태 저장됨: {data}")
    except Exception as e:
        print(f"창 상태 저장 중 오류: {e}")
