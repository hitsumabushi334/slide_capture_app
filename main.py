# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from datetime import datetime
import os
import logging  # logging モジュールをインポート
import traceback  # スタックトレース取得のため

import cv2
from PIL import Image, ImageGrab, UnidentifiedImageError
import numpy as np

# --- ロギング設定 ---
log_filename = "slide_capture_app.log"
log_format = "%(asctime)s - %(levelname)s - %(threadName)s - %(message)s"
logging.basicConfig(
    level=logging.INFO,
    format=log_format,
    handlers=[
        logging.FileHandler(log_filename, encoding="utf-8"),  # ファイル出力
        logging.StreamHandler(),  # コンソールにも出力 (デバッグ用)
    ],
)
logger = logging.getLogger(__name__)


class SlideCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("スライドキャプチャ")
        # UIの高さを少し増やしてエラーメッセージ表示スペースを確保
        self.root.geometry("400x200")

        self.is_capturing = False
        self.capture_thread = None
        self.last_image = None
        self.saved_count = 0
        self.last_saved_filename = ""
        self.save_folder_name = tk.StringVar()
        self.error_occurred_in_thread = False  # スレッド内エラーフラグ

        # --- UI要素の作成 ---
        folder_frame = ttk.Frame(root, padding="10")
        folder_frame.pack(fill=tk.X)
        folder_label = ttk.Label(folder_frame, text="保存フォルダ名:")
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(
            folder_frame, textvariable=self.save_folder_name, width=35
        )
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        button_frame = ttk.Frame(root, padding="10")
        button_frame.pack(fill=tk.X)
        self.start_button = ttk.Button(
            button_frame, text="開始", command=self.start_capture
        )
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(
            button_frame, text="終了", command=self.stop_capture, state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)

        status_frame = ttk.Frame(root, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)
        # wraplengthでテキストの折り返しを設定
        self.status_label = ttk.Label(
            status_frame, text="待機中...", anchor=tk.W, justify=tk.LEFT, wraplength=380
        )
        self.status_label.pack(fill=tk.BOTH, expand=True)

        self.root.bind("<Escape>", lambda e: self.stop_capture())
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        logger.info("アプリケーションを初期化しました。")

    def update_status(self):
        """ステータスラベルを更新する"""
        if self.is_capturing:
            status_text = f"キャプチャ中...\n保存枚数: {self.saved_count}\n最終保存: {self.last_saved_filename}"
            if self.error_occurred_in_thread:
                # エラー発生時はログファイル参照を促すメッセージを追加
                status_text += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.status_label.config(foreground="red")  # エラー時は赤文字
            else:
                self.status_label.config(foreground="black")  # 通常時は黒文字
            self.status_label.config(text=status_text)
            # is_capturing が True の間だけ再スケジュール
            if self.is_capturing:
                self.root.after(1000, self.update_status)
        else:
            final_status = f"停止しました。\n合計保存枚数: {self.saved_count}"
            if self.last_saved_filename:
                final_status += f"\n最終保存: {self.last_saved_filename}"
            if self.error_occurred_in_thread:
                final_status += f"\n警告: エラー発生。詳細はログファイル\n({log_filename})を確認してください。"
                self.status_label.config(foreground="red")
            else:
                self.status_label.config(foreground="black")
            self.status_label.config(text=final_status)

    def start_capture(self):
        """キャプチャ処理を開始する"""
        if self.is_capturing:
            return

        folder_name_input = self.save_folder_name.get().strip()
        if not folder_name_input:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            folder_name = f"saved_screenshots_{timestamp}"
            self.save_folder_name.set(folder_name)
            logger.info(f"フォルダ名が未入力のため、デフォルト名を設定: {folder_name}")
        else:
            # ファイル名として不適切な文字を置換 (簡易的な対策)
            invalid_chars = '<>:"/\\|?*'
            folder_name = "".join(
                c if c not in invalid_chars else "_" for c in folder_name_input
            )
            if folder_name != folder_name_input:
                self.save_folder_name.set(folder_name)
                warning_msg = f"フォルダ名に使用できない文字が含まれていたため、'{folder_name}' に修正しました。"
                logger.warning(warning_msg)
                messagebox.showwarning("フォルダ名修正", warning_msg)

        try:
            self.capture_save_path = os.path.abspath(folder_name)
            logger.info(f"保存先フォルダの絶対パス: {self.capture_save_path}")

            if not os.path.exists(self.capture_save_path):
                logger.info(
                    f"フォルダが存在しないため作成します: {self.capture_save_path}"
                )
                os.makedirs(self.capture_save_path, exist_ok=True)
                logger.info(f"フォルダを作成しました: {self.capture_save_path}")
            elif not os.path.isdir(self.capture_save_path):
                error_msg = (
                    f"指定されたパスはフォルダではありません: {self.capture_save_path}"
                )
                logger.error(error_msg)
                messagebox.showerror("エラー", error_msg)
                return
            # 書き込み権限チェック (Windows用簡易チェック)
            elif not os.access(self.capture_save_path, os.W_OK):
                error_msg = f"指定されたフォルダへの書き込み権限がありません:\n{self.capture_save_path}\n別のフォルダを指定するか、権限を確認してください。"
                logger.error(error_msg)
                messagebox.showerror("権限エラー", error_msg)
                return
            else:
                logger.info(
                    f"保存先フォルダへの書き込み権限を確認しました: {self.capture_save_path}"
                )

        except OSError as e:
            error_detail = (
                f"フォルダの作成/アクセス中にOSエラーが発生しました。\n"
                f"エラータイプ: {type(e).__name__}\n"
                f"エラーコード: {e.errno}\n"
                f"メッセージ: {e.strerror}\n"
                f"パス: {getattr(e, 'filename', 'N/A')}"
            )
            # スタックトレース付きでログ記録
            logger.error(f"{error_detail}\n{traceback.format_exc()}")
            messagebox.showerror(
                "フォルダエラー",
                f"{error_detail}\n詳細はログファイル ({log_filename}) を確認してください。",
            )
            return
        except Exception as e:
            error_detail = (
                f"フォルダ処理中に予期せぬエラーが発生しました:\n"
                f"エラータイプ: {type(e).__name__}\n"
                f"メッセージ: {e}"
            )
            logger.exception(error_detail)  # exceptionはスタックトレースを自動で記録
            messagebox.showerror(
                "予期せぬエラー",
                f"{error_detail}\n詳細はログファイル ({log_filename}) を確認してください。",
            )
            return

        # エラーフラグをリセット
        self.error_occurred_in_thread = False
        self.status_label.config(foreground="black")  # ステータス色をリセット

        self.is_capturing = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.folder_entry.config(state=tk.DISABLED)
        self.saved_count = 0
        self.last_saved_filename = ""
        self.last_image = None

        # スレッドを開始
        self.capture_thread = threading.Thread(
            target=self.capture_loop, name="CaptureThread", daemon=True
        )
        self.capture_thread.start()
        self.update_status()  # 定期的なステータス更新を開始
        logger.info(f"キャプチャを開始しました。保存先: {self.capture_save_path}")

    def stop_capture(self):
        """キャプチャ処理を停止する"""
        if not self.is_capturing:
            return

        logger.info("キャプチャ停止処理を開始します。")
        self.is_capturing = False
        # スレッドが終了するのを少し待つ
        if self.capture_thread and self.capture_thread.is_alive():
            logger.info("キャプチャループの終了を待っています...")
            self.capture_thread.join(timeout=1.5)  # 少し長めに待つ
            if self.capture_thread.is_alive():
                logger.warning("警告: キャプチャスレッドが時間内に終了しませんでした。")

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.folder_entry.config(state=tk.NORMAL)
        logger.info("キャプチャを停止しました。")
        # 停止後にもう一度ステータスを更新して最終結果を表示
        # is_capturingがFalseなのでupdate_statusは再スケジュールされない
        self.root.after(100, self.update_status)

    def capture_loop(self):
        """定期的にスクリーンショットを取得し、比較・保存するループ"""
        logger.info("キャプチャループを開始します。")
        while self.is_capturing:
            try:
                # 1. スクリーンショット取得
                screenshot = ImageGrab.grab()
                if screenshot is None:
                    logger.error(
                        "エラー: ImageGrab.grab() が None を返しました。スクリーンショットを取得できませんでした。"
                    )
                    self.error_occurred_in_thread = True
                    time.sleep(2.0)  # 少し待ってリトライ
                    continue

                # 2. 画像形式変換 (PIL -> OpenCV)
                current_image_pil = screenshot.convert("RGB")
                current_image_cv = np.array(current_image_pil)
                current_image_cv = cv2.cvtColor(current_image_cv, cv2.COLOR_RGB2BGR)

                if current_image_cv is None or current_image_cv.size == 0:
                    logger.error(
                        "エラー: 画像データの変換に失敗しました (Noneまたはサイズ0)。"
                    )
                    self.error_occurred_in_thread = True
                    time.sleep(2.0)
                    continue

                # 3. 前回の画像と比較
                if self.last_image is None:
                    logger.info("最初の画像を取得しました。保存します。")
                    self.save_image(current_image_cv)
                    self.last_image = current_image_cv
                else:
                    if not self.is_similar(current_image_cv, self.last_image):
                        logger.info(
                            "新しい画像または類似していない画像を検出しました。保存します。"
                        )
                        self.save_image(current_image_cv)
                        self.last_image = current_image_cv
                    else:
                        # logger.debug("類似画像のためスキップ") # DEBUGレベルに変更
                        pass

            except (OSError, UnidentifiedImageError) as e:
                logger.exception(f"エラー (キャプチャ/変換): {type(e).__name__} - {e}")
                self.error_occurred_in_thread = True
                time.sleep(5.0)
            except cv2.error as e:
                logger.exception(f"エラー (OpenCV): {type(e).__name__} - {e}")
                self.error_occurred_in_thread = True
                time.sleep(5.0)
            except Exception as e:
                logger.exception(
                    f"エラー (キャプチャループ): 予期せぬエラーが発生しました - {type(e).__name__}: {e}"
                )
                self.error_occurred_in_thread = True
                time.sleep(5.0)

            # 次のキャプチャまでの待機時間
            wait_start_time = time.time()
            while self.is_capturing and time.time() - wait_start_time < 2.0:
                time.sleep(0.1)

        logger.info("キャプチャループが終了しました。")

    def is_similar(self, img1_cv, img2_cv, threshold=0.95):
        """2つの画像の類似度を計算する (差分ベースの簡易比較)"""
        try:
            gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)
            h1, w1 = gray1.shape
            h2, w2 = gray2.shape

            if h1 != h2 or w1 != w2:
                # logger.debug(f"画像サイズが異なるためリサイズします: ({w1}x{h1}) vs ({w2}x{h2})")
                if h1 * w1 < h2 * w2:
                    gray2 = cv2.resize(gray2, (w1, h1), interpolation=cv2.INTER_AREA)
                else:
                    gray1 = cv2.resize(gray1, (w2, h2), interpolation=cv2.INTER_AREA)

            diff = cv2.absdiff(gray1, gray2)
            non_zero_count = np.count_nonzero(diff)
            total_pixels = gray1.size
            if total_pixels == 0:
                logger.warning("警告: 類似度計算中の画像サイズが0です。")
                return False

            similarity = 1.0 - (non_zero_count / total_pixels)
            # logger.debug(f"類似度: {similarity:.4f}")
            return similarity >= threshold

        except cv2.error as e:
            logger.exception(f"エラー (類似度計算 - OpenCV): {type(e).__name__} - {e}")
            self.error_occurred_in_thread = True
            return False
        except Exception as e:
            logger.exception(
                f"エラー (類似度計算): 予期せぬエラー - {type(e).__name__}: {e}"
            )
            self.error_occurred_in_thread = True
            return False

    def save_image(self, image_cv):
        """画像をファイルに保存する。エラー発生時はログに記録。"""
        save_path = None
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            filename = f"screenshot_{timestamp}.png"
            save_path = os.path.join(self.capture_save_path, filename)

            if image_cv is None or image_cv.size == 0:
                logger.warning(
                    f"警告: 保存しようとした画像データが無効です (Noneまたはサイズ0)。パス: {save_path}"
                )
                self.error_occurred_in_thread = True
                return

            success = cv2.imwrite(save_path, image_cv, [cv2.IMWRITE_PNG_COMPRESSION, 3])

            if success:
                self.saved_count += 1
                self.last_saved_filename = os.path.basename(save_path)
                logger.info(f"画像を保存しました: {save_path}")
            else:
                # imwriteがFalseを返した場合
                error_msg = (
                    f"エラー (保存): cv2.imwrite が False を返しました。ファイル書き込みに失敗した可能性があります。\n"
                    f" - 保存試行パス: {save_path}\n"
                    f" - 画像サイズ: {image_cv.shape if image_cv is not None else 'None'}\n"
                    f" - 考えられる原因: 書き込み権限不足、ディスク容量不足、パス名の問題、画像データ破損など"
                )
                logger.error(error_msg)
                self.error_occurred_in_thread = True

        except cv2.error as e:
            logger.exception(
                f"エラー (保存 - OpenCV): {type(e).__name__} - {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_thread = True
        except OSError as e:
            logger.exception(
                f"エラー (保存 - OS): {type(e).__name__} - {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_thread = True
        except Exception as e:
            logger.exception(
                f"エラー (保存): 予期せぬエラー - {type(e).__name__}: {e}\n - 保存試行パス: {save_path}"
            )
            self.error_occurred_in_thread = True

    def on_closing(self):
        """ウィンドウが閉じられたときの処理"""
        if self.is_capturing:
            if messagebox.askokcancel(
                "確認", "キャプチャ処理が実行中です。\n本当に終了しますか？"
            ):
                logger.info("ユーザー操作により終了します...")
                self.stop_capture()
                self.root.destroy()
                logger.info("アプリケーションを終了しました。")
            else:
                logger.info("終了操作がキャンセルされました。")
                return
        else:
            logger.info("アプリケーションを終了しました。")
            self.root.destroy()


if __name__ == "__main__":
    # Tkinterのルートウィンドウを作成
    root = tk.Tk()
    # アプリケーションクラスのインスタンスを作成
    app = SlideCaptureApp(root)
    # Tkinterのイベントループを開始
    root.mainloop()
