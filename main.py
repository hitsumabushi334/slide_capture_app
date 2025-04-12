# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time
from datetime import datetime
import os

import cv2
from PIL import Image, ImageGrab
import numpy as np

class SlideCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("スライドキャプチャ")
        self.root.geometry("400x180")

        self.is_capturing = False
        self.capture_thread = None
        self.last_image = None
        self.saved_count = 0
        self.last_saved_filename = ""
        self.last_saved_filepath = None # 前回保存したファイルのフルパス
        self.save_folder_name = tk.StringVar()

        # --- UI要素の作成 ---
        folder_frame = ttk.Frame(root, padding="10")
        folder_frame.pack(fill=tk.X)
        folder_label = ttk.Label(folder_frame, text="保存フォルダ名:")
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.save_folder_name, width=35)
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)

        button_frame = ttk.Frame(root, padding="10")
        button_frame.pack(fill=tk.X)
        self.start_button = ttk.Button(button_frame, text="開始", command=self.start_capture)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(button_frame, text="終了", command=self.stop_capture, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        status_frame = ttk.Frame(root, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)
        self.status_label = ttk.Label(status_frame, text="待機中...", anchor=tk.W, justify=tk.LEFT)
        self.status_label.pack(fill=tk.BOTH, expand=True)

        self.root.bind('<Escape>', lambda e: self.stop_capture())
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def update_status(self):
        """ステータスラベルを更新する"""
        if self.is_capturing:
            status_text = f"キャプチャ中...\n保存枚数: {self.saved_count}\n最終保存: {self.last_saved_filename}"
            self.status_label.config(text=status_text)
            self.root.after(1000, self.update_status)
        else:
            final_status = f"停止しました。\n合計保存枚数: {self.saved_count}"
            if self.last_saved_filename:
                final_status += f"\n最終保存: {self.last_saved_filename}"
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
        else:
            folder_name = folder_name_input

        self.capture_save_path = os.path.abspath(folder_name)

        try:
            if not os.path.exists(self.capture_save_path):
                os.makedirs(self.capture_save_path)
                print(f"フォルダを作成しました: {self.capture_save_path}")
            elif not os.path.isdir(self.capture_save_path):
                 messagebox.showerror("エラー", f"指定された名前はファイルとして存在します: {self.capture_save_path}")
                 return
        except OSError as e:
            messagebox.showerror("エラー", f"フォルダの作成/アクセスに失敗しました: {e}")
            return
        except Exception as e:
             messagebox.showerror("エラー", f"予期せぬエラーが発生しました: {e}")
             return

        self.is_capturing = True
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.folder_entry.config(state=tk.DISABLED)
        self.saved_count = 0
        self.last_saved_filename = ""
        self.last_image = None

        self.capture_thread = threading.Thread(target=self.capture_loop, daemon=True)
        self.capture_thread.start()
        self.update_status()
        print(f"キャプチャを開始しました。保存先: {self.capture_save_path}")

    def stop_capture(self):
        """キャプチャ処理を停止する"""
        if not self.is_capturing:
            return

        self.is_capturing = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.folder_entry.config(state=tk.NORMAL)
        print("キャプチャを停止しました。")
        self.root.after(100, self.update_status)

    def capture_loop(self):
        """定期的にスクリーンショットを取得し、比較・保存するループ"""
        while self.is_capturing:
            try:
                screenshot = ImageGrab.grab()
                current_image_pil = screenshot.convert('RGB')
                current_image_cv = np.array(current_image_pil)
                current_image_cv = cv2.cvtColor(current_image_cv, cv2.COLOR_RGB2BGR)

                if self.last_image is None or not self.is_similar(current_image_cv, self.last_image):
                    self.save_image(current_image_cv)
                    self.last_image = current_image_cv
                else:
                    # 画像が類似している場合、品質を比較して置き換えるか判断
                    current_quality = self.evaluate_image_quality(current_image_cv)
                    # last_image は前回保存または高品質と判断された画像データ
                    last_quality = self.evaluate_image_quality(self.last_image)

                    print(f"品質比較 - 現在: {current_quality}, 前回: {last_quality}")

                    # ファイルサイズが大きい方を優先、同じならシャープネスが高い方
                    if current_quality[1] > last_quality[1] or \
                       (current_quality[1] == last_quality[1] and current_quality[0] > last_quality[0]):
                        print("新規画像の方が高品質なため、古い画像を置き換えます。")
                        # 古いファイルを削除
                        if self.last_saved_filepath and os.path.exists(self.last_saved_filepath):
                            try:
                                os.remove(self.last_saved_filepath)
                                print(f"古い画像を削除: {self.last_saved_filepath}")
                            except OSError as e:
                                print(f"古い画像の削除に失敗: {e}")
                        # 新しい画像を保存
                        self.save_image(current_image_cv)
                        self.last_image = current_image_cv # last_image を高品質な新しい画像に更新
                    else:
                        print("新規画像の方が品質が低い、または同等なためスキップ")
                        pass # 新規画像を破棄

            except Exception as e:
                print(f"キャプチャ中にエラーが発生しました: {e}")

            wait_start_time = time.time()
            while self.is_capturing and time.time() - wait_start_time < 2.0:
                time.sleep(0.1)

    def is_similar(self, img1_cv, img2_cv, threshold=0.85):
        """2つの画像の類似度を計算する (SSIMを使用) - 修正版"""
        try:
            gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)

            h1, w1 = gray1.shape
            h2, w2 = gray2.shape

            # サイズが異なる場合は、img1のサイズにimg2をリサイズ
            if h1 != h2 or w1 != w2:
                print(f"画像サイズが異なるためリサイズします: ({w1}x{h1}) vs ({w2}x{h2})")
                gray2 = cv2.resize(gray2, (w1, h1), interpolation=cv2.INTER_AREA)

            # SSIM計算 (簡易差分ベース)
            diff = cv2.absdiff(gray1, gray2)
            non_zero_count = np.count_nonzero(diff)
            if gray1.size == 0:
                print("警告: 画像サイズが0です。")
                return False
            similarity = 1.0 - (non_zero_count / gray1.size)

            print(f"類似度: {similarity:.4f}")
            return similarity >= threshold

        except cv2.error as e:
             print(f"OpenCVエラー（類似度計算中）: {e}")
             return False
        except Exception as e:
            print(f"類似度計算中に予期せぬエラー: {e}")
            return False

    def save_image(self, image_cv):
        """画像をファイルに保存する"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            filename = f"screenshot_{timestamp}.png"
            save_path = os.path.join(self.capture_save_path, filename)

            cv2.imwrite(save_path, image_cv)

            self.saved_count += 1
            self.last_saved_filename = filename
            self.last_saved_filepath = save_path # 保存したファイルのフルパスを記録
            # print(f"画像を保存しました: {save_path}")
        except Exception as e:
            print(f"画像保存中にエラー: {e}")

    def evaluate_image_quality(self, image_cv):
        """画像の品質を評価する (シャープネスとファイルサイズ)"""
        try:
            # 1. シャープネス (ラプラシアン分散)
            gray = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
            laplacian_var = cv2.Laplacian(gray, cv2.CV_64F).var()

            # 2. ファイルサイズ (メモリ上でのPNGエンコードサイズ)
            #    品質比較のため、圧縮率の低いPNGで比較する
            result, encoded_image = cv2.imencode('.png', image_cv)
            if not result:
                print("警告: 画像のエンコードに失敗しました。")
                file_size = 0
            else:
                file_size = len(encoded_image)

            return laplacian_var, file_size
        except cv2.error as e:
            print(f"OpenCVエラー（品質評価中）: {e}")
            return 0.0, 0 # エラー時は最低品質とする
        except Exception as e:
            print(f"品質評価中に予期せぬエラー: {e}")
            return 0.0, 0 # エラー時は最低品質とする

    def on_closing(self):
        """ウィンドウが閉じられたときの処理"""
        if self.is_capturing:
            if messagebox.askokcancel("確認", "キャプチャ中に終了しますか？"):
                self.stop_capture()
                self.root.destroy()
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SlideCaptureApp(root)
    root.mainloop()