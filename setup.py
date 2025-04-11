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
base_type = None
if sys.platform == "win32":
    base_type = "gui"  # 文字列リテラルを使用

# --- 実行可能ファイルの定義 ---
executables = [
    Executable(
        "main.py",
        base=base_type,  # 変数名をbase_typeに変更
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
    "packages": ["tkinter", "cv2", "numpy"], # Pillowを削除し、cx_Freezeの自動検出に期待
    "includes": [],
    "excludes": ["tkinter.test", "tkinter.tix", "distutils"], # unittestを除外リストから削除
    "include_files": [icon_path] if os.path.exists(icon_path) else [], # 環境固有のパス指定を削除
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
