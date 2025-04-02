#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import logging

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("target_files.log"),
        logging.StreamHandler()
    ]
)

def create_target_files():
    """
    目標値を保存するテキストファイルが存在しない場合に作成します。
    """
    # 週間目標ファイル
    if not os.path.exists("week_target_values.txt"):
        try:
            with open("week_target_values.txt", "w", encoding="utf-8") as file:
                file.write("0\n0")
            print("week_target_values.txt ファイルを作成しました")
        except Exception as e:
            logging.error(f"週間目標ファイルの作成でエラー: {str(e)}")
    
    # 月間目標ファイル
    if not os.path.exists("month_target_values.txt"):
        try:
            with open("month_target_values.txt", "w", encoding="utf-8") as file:
                file.write("0\n0")
            print("month_target_values.txt ファイルを作成しました")
        except Exception as e:
            logging.error(f"月間目標ファイルの作成でエラー: {str(e)}")
    
    # 最終目標ファイル
    if not os.path.exists("last_target_values.txt"):
        try:
            with open("last_target_values.txt", "w", encoding="utf-8") as file:
                file.write("0\n0\n0")
            print("last_target_values.txt ファイルを作成しました")
        except Exception as e:
            logging.error(f"最終目標ファイルの作成でエラー: {str(e)}")

if __name__ == "__main__":
    create_target_files()
    print("すべての目標値ファイルを作成しました。") 