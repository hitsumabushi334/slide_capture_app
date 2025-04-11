import sys
from cx_Freeze import setup, Executable
import os

# --- アプリケーション情報 ---
app_name = "SlideCaptureApp"
version = "1.0"
description = "スライドショーのスクリーンショットを自動保存するアプリケーション"
author = "Your Name" # 必要に応じて変更してください
icon_path = "app_icon.ico" # アイコンファイルのパス。不要な場合はNoneまたはコメントアウト

# --- ビルドオプション ---
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# --- 実行可能ファイルの定義 ---
executables = [
    Executable(
        "main.py",
        base=base,
        icon=icon_path if os.path.exists(icon_path) else None, # アイコンが存在する場合のみ設定
        target_name="SlideCaptureApp.exe"
    )
]

# --- セットアップ設定 ---
# build_exe オプション:
#   packages: cx_Freezeが自動検出できない場合に含めるパッケージ
#   includes: 含めるモジュール
#   excludes: 除外するモジュール
#   include_files: 追加ファイル (アイコンなど)
#   zip_include_packages: zipファイルに含めるパッケージ (デフォルトは*)
#   zip_exclude_packages: zipファイルから除外するパッケージ (例: tkinter)
build_exe_options = {
    "packages": ["tkinter", "cv2", "numpy", "PIL"], # PillowはPILとして認識されることがある
    "includes": [],
    "excludes": ["tkinter.test", "tkinter.tix", "distutils", "unittest"], # 不要なものを除外
    "include_files": [icon_path] if os.path.exists(icon_path) else [], # アイコンが存在する場合のみ含める
    "zip_include_packages": ["*"],
    "zip_exclude_packages": ["tkinter"], # tkinterはzipに含めない方が良い場合がある
}

setup(
    name=app_name,
    version=version,
    description=description,
    author=author,
    options={"build_exe": build_exe_options},
    executables=executables,
)

# setup.pyの先頭に os をインポートするのを忘れないように
