# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox # filedialogを削除
import threading
import time
from datetime import datetime
import os
# import keyboard # keyboardライブラリは現在未使用のためコメントアウト (Escキー処理はTkinterのbindで実装)
from PIL import ImageGrab, Image
import cv2
import numpy as np

class SlideCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("スライドキャプチャ")
        self.root.geometry("400x180") # ウィンドウサイズ調整 (参照ボタン削除分)

        self.is_capturing = False
        self.capture_thread = None
        self.last_image = None
        self.saved_count = 0
        self.last_saved_filename = ""
        self.save_folder_name = tk.StringVar() # 変数名を変更して意図を明確化

        # --- UI要素の作成 ---
        # 保存先フォルダ名指定
        folder_frame = ttk.Frame(root, padding="10")
        folder_frame.pack(fill=tk.X)
        folder_label = ttk.Label(folder_frame, text="保存フォルダ名:") # ラベルテキスト変更
        folder_label.pack(side=tk.LEFT, padx=(0, 5))
        self.folder_entry = ttk.Entry(folder_frame, textvariable=self.save_folder_name, width=35) # 幅調整
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        # browse_button を削除

        # 開始・終了ボタン
        button_frame = ttk.Frame(root, padding="10")
        button_frame.pack(fill=tk.X)
        self.start_button = ttk.Button(button_frame, text="開始", command=self.start_capture)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.stop_button = ttk.Button(button_frame, text="終了", command=self.stop_capture, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)

        # ステータス表示
        status_frame = ttk.Frame(root, padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True)
        self.status_label = ttk.Label(status_frame, text="待機中...", anchor=tk.W, justify=tk.LEFT)
        self.status_label.pack(fill=tk.BOTH, expand=True)

        # Escキーでの終了設定
        self.root.bind('<Escape>', lambda e: self.stop_capture())

        # ウィンドウが閉じられたときの処理
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # browse_folder メソッドを削除

    def update_status(self):
        """ステータスラベルを更新する"""
        if self.is_capturing:
            status_text = f"キャプチャ中...\n保存枚数: {self.saved_count}\n最終保存: {self.last_saved_filename}"
            self.status_label.config(text=status_text)
            self.root.after(1000, self.update_status)
        else:
            # キャプチャ停止後、最終状態を表示
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
            # デフォルトフォルダ名生成
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            folder_name = f"saved_screenshots_{timestamp}"
            self.save_folder_name.set(folder_name) # UIにも反映
        else:
            # ユーザー指定のフォルダ名を使用
            # ファイル名として不正な文字をチェック・置換する処理を追加するとより堅牢になる
            folder_name = folder_name_input

        # 保存先フォルダ作成 (常にカレントディレクトリ直下に作成)
        # 注意: ここでは './' を起点とする相対パスで作成します。
        #       もし絶対パスを指定したい場合は、ユーザー入力の解釈方法を変更する必要があります。
        self.capture_save_path = os.path.abspath(folder_name) # 絶対パスに変換

        try:
            # フォルダが存在しない場合のみ作成
            if not os.path.exists(self.capture_save_path):
                os.makedirs(self.capture_save_path)
                print(f"フォルダを作成しました: {self.capture_save_path}")
            # フォルダが既に存在する場合でもエラーとしない
            elif not os.path.isdir(self.capture_save_path):
                 messagebox.showerror("エラー", f"指定された名前はファイルとして存在します: {self.capture_save_path}")
                 return

        except OSError as e:
            messagebox.showerror("エラー", f"フォルダの作成/アクセスに失敗しました: {e}")
            return
        except Exception as e: # その他の予期せぬエラー
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
        # スレッドの終了待機は不要 (daemon=True)
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.folder_entry.config(state=tk.NORMAL)
        print("キャプチャを停止しました。")
        # update_statusが最終状態を表示するように呼び出す
        self.root.after(100, self.update_status) # 少し遅延させてから最終ステータス更新

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
                    # TODO: 要件定義書の「よりきれいで情報量が多い画像」の判断基準を実装
                    print("類似画像のためスキップ")
                    pass

            except Exception as e:
                print(f"キャプチャ中にエラーが発生しました: {e}")
                # 必要に応じてエラー時の処理を追加 (例: ログ記録、UI通知)

            wait_start_time = time.time()
            while self.is_capturing and time.time() - wait_start_time < 2.0:
                time.sleep(0.1)

    def is_similar(self, img1_cv, img2_cv, threshold=0.9):
        """2つの画像の類似度を計算する (SSIMを使用) - 改善版"""
        try:
            gray1 = cv2.cvtColor(img1_cv, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2_cv, cv2.COLOR_BGR2GRAY)

            h1, w1 = gray1.shape
            h2, w2 = gray2.shape

            # サイズが異なる場合は、小さい方のサイズに大きい方をリサイズ
            if h1 != h2 or w1 != w2:
                if h1 * w1 < h2 * w2: # img1の方が小さい
                    gray2 = cv2.resize(gray2, (w1, h1), interpolation=cv2.INTER_AREA)
                else: # img2の方が小さい or 同じ
                    gray1 = cv2.resize(gray1, (w2, h2), interpolation=cv2.INTER_AREA)
                # リサイズ後のサイズを再取得
                h1, w1 = gray1.shape


            # SSIM計算 (OpenCVには直接SSIMを計算する関数がないため、差分ベースの簡易類似度を使用)
            # より正確な比較には skimage.metrics.structural_similarity や他の手法を検討
            diff = cv2.absdiff(gray1, gray2)
            non_zero_count = np.count_nonzero(diff)
            # 画像サイズが0でないことを確認
            if gray1.size == 0:
                print("警告: 画像サイズが0です。")
                return False
            similarity = 1.0 - (non_zero_count / gray1.size)

            print(f"類似度: {similarity:.4f}")
            return similarity >= threshold

        except cv2.error as e:
             print(f"OpenCVエラー（類似度計算中）: {e}")
             return False # OpenCVエラー時は類似していないと判断
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
            # print(f"画像を保存しました: {save_path}") # ログは任意
        except Exception as e:
            print(f"画像保存中にエラー: {e}")

    def on_closing(self):
        """ウィンドウが閉じられたときの処理"""
        if self.is_capturing:
            if messagebox.askokcancel("確認", "キャプチャ中に終了しますか？"):
                self.stop_capture()
                # stop_capture内で最終ステータス更新が予約される
                self.root.destroy() # 少し待ってから破棄した方が良い場合もある
        else:
            self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SlideCaptureApp(root)
    root.mainloop()