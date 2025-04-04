import tkinter as tk
from tkinter import simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
import keyboard
import re
import time
import os
from datetime import datetime, timedelta
import sys
import logging
import shutil
import tempfile

# ログ設定
log_file = 'weekgoal_error.log'
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def calculate_week_number():
    """
    2024年12月30日を1週目として、現在の週数を計算します。
    """
    start_date = datetime(2024, 12, 30)  # 1週目の開始日
    current_date = datetime.now()
    days_diff = (current_date - start_date).days
    week_number = (days_diff // 7) + 1
    return week_number


def recalc_total_week():
    """
    apolloサイトの指定された週の [実] を合計して返します。
    """
    total = 0
    try:
        # ヘッダー行から日付を取得
        date_headers = WebDriverWait(extraction_driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//thead//th[contains(@class, 'sticky')]"))
        )
        
        # 合計行を取得
        totals_row = WebDriverWait(extraction_driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//tr[td[@class='total']]"))
        )
        
        # 合計行のセルを取得（最初と最後のセルを除く）
        total_cells = totals_row.find_elements(By.XPATH, ".//td[@class='total sticky']")
        
        # 各日の [実] の値を合計
        for cell in total_cells:
            spans = cell.find_elements(By.TAG_NAME, "span")
            for span in spans:
                match = re.search(r"\[実\](\d+)", span.text)
                if match:
                    total += int(match.group(1))
                    break
    except Exception as e:
        print(f"データ取得エラー (apollo): {e}")
    return total

def recalc_total_month():
    """
    apolloサイトの22クリニックの [実] を合計して返します。（A残込）
    """
    total = 0
    target_clinics = [
        "札幌院", "仙台院", "宇都宮", "高崎院", "大宮院", "柏院", "船橋院",
        "新宿院", "新橋院", "神田院", "立川院", "横浜院", "静岡院", "名古屋院", "梅田院",
        "心斎橋", "なんば院", "神戸院", "高松", "広島院", "博多", "天神"
    ]
    try:
        WebDriverWait(extraction_driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "clinic")))
        clinic_cells = extraction_driver.find_elements(By.CLASS_NAME, "clinic")
        for cell in clinic_cells:
            try:
                clinic_name = cell.find_element(By.TAG_NAME, "a").text.strip()
            except NoSuchElementException:
                continue
            if clinic_name in target_clinics:
                spans = cell.find_elements(By.TAG_NAME, "span")
                for span in spans:
                    match = re.search(r"\[実\](\d+)", span.text)
                    if match:
                        total += int(match.group(1))
                        break
    except Exception as e:
        print(f"データ取得エラー (apollo): {e}")
    return total

def get_oguchi_value(driver, page_url):
    """
    tomatoサイトの指定ページにアクセスし、
    ① 3/3-9の範囲で実質大口をカウント
    ② <a>タグをすべて調べ、outerHTMLに "fa fa-wifi" または "大[残未計]" が含まれる場合は除外
    戻り値は (total_oguchi, exclude_count, net_value) です。
    """
    driver.get(page_url)
    time.sleep(2)  # ページ読み込み待ち

    total_oguchi = 0
    exclude_count = 0

    try:
        # 現在の週の月曜日と日曜日を取得
        monday, sunday = get_current_week_range()
        
        # 日付範囲内のセルを取得
        for day in range(7):  # 月曜から日曜まで
            target_date = monday + timedelta(days=day)
            date_str = target_date.strftime('%Y-%m-%d')
            
            # その日付のセルを取得
            cell_xpath = f"//td[.//a[contains(@href, '{date_str}')]]"
            try:
                cell = driver.find_element(By.XPATH, cell_xpath)
                
                # 予約を取得（class="rest"で始まるp要素内のaタグ）
                reservation_links = cell.find_elements(By.CSS_SELECTOR, 'p[class^="rest"] > a')
                for a in reservation_links:
                    text = a.get_attribute("innerText")
                    outer_html = a.get_attribute("outerHTML")
                    
                    # 術検のみの場合はスキップ
                    if "術検" in text and not any(k in text for k in ["包", "大", "長", "非長", "陰大"]):
                        continue
                        
                    # EDのみの場合はスキップ
                    if "ED" in text and not any(k in text for k in ["包", "大", "長", "非長", "陰大"]):
                        continue
                    
                    # 実質大口のカウント
                    if any(keyword in text for keyword in ["包", "大", "長", "非長", "陰大"]):
                        total_oguchi += 1
                    
                    # 除外カウント（wifi または 大[残未計]）
                    if "fa fa-wifi" in outer_html or "大[残未計]" in outer_html:
                        exclude_count += 1
                        
            except NoSuchElementException:
                continue

    except Exception as e:
        print(f"データ取得エラー (tomato): {e}")

    net_value = total_oguchi - exclude_count
    return total_oguchi, exclude_count, net_value

def get_oguchi_value_month(driver, page_url):
    """
    tomatoサイトの指定ページにアクセスし、
      ① ページ全体から "[大]XX" のXX（大口件数）を取得
      ② ページ内の<a>タグをすべて調べ、outerHTMLに "fa fa-wifi" または "大[残未計]" が含まれる場合は除外カウントを加算
      ③ 大口件数から除外カウントを引いたネット値を返す
    戻り値は (big_count, exclude_count, net_value) です。
    """
    driver.get(page_url)
    time.sleep(2)  # ページ読み込み待ち
    page_source = driver.page_source
    big_match = re.search(r"\[大\](\d+)", page_source)
    if big_match:
        big_count = int(big_match.group(1))
    else:
        big_count = 0

    exclude_count = 0
    anchors = driver.find_elements(By.TAG_NAME, "a")
    for a in anchors:
        outer_html = a.get_attribute("outerHTML")
        if ("fa fa-wifi" in outer_html) or ("大[残未計]" in outer_html):
            exclude_count += 1

    net_value = big_count - exclude_count
    return big_count, exclude_count, net_value

def read_target_values():
    """
    目標値をテキストファイルから読み取ります。
    ファイルが存在しない場合はデフォルト値を使用します。
    """
    try:
        with open("week_target_values.txt", "r", encoding="utf-8") as file:
            lines = file.readlines()
            a_target = int(lines[0].strip())
            k_target = int(lines[1].strip())
            return a_target, k_target
    except Exception as e:
        print(f"目標値ファイルの読み込みエラー: {e}")
        # デフォルト値を返す
        return 0, 0

def save_target_values(a_target, k_target):
    """
    週間目標値をテキストファイルに保存します。
    """
    try:
        with open("week_target_values.txt", "w", encoding="utf-8") as file:
            file.write(f"{a_target}\n{k_target}")
        logging.info(f"週間目標値を保存: A={a_target}, K={k_target}")
        return True
    except Exception as e:
        logging.error(f"週間目標値の保存でエラー: {str(e)}")
        return False

def read_month_target_values():
    """
    月間目標値をテキストファイルから読み取ります。
    ファイルが存在しない場合はデフォルト値を使用します。
    """
    try:
        with open("month_target_values.txt", "r", encoding="utf-8") as file:
            lines = file.readlines()
            a_target = int(lines[0].strip())
            k_target = int(lines[1].strip())
            return a_target, k_target
    except Exception as e:
        print(f"月間目標値ファイルの読み込みエラー: {e}")
        # デフォルト値を返す
        return 0, 0

def save_month_target_values(a_target, k_target):
    """
    月間目標値をテキストファイルに保存します。
    """
    try:
        with open("month_target_values.txt", "w", encoding="utf-8") as file:
            file.write(f"{a_target}\n{k_target}")
        logging.info(f"月間目標値を保存: A={a_target}, K={k_target}")
        return True
    except Exception as e:
        logging.error(f"月間目標値の保存でエラー: {str(e)}")
        return False

def read_last_target_values():
    """
    最終目標値をテキストファイルから読み取ります。
    ファイルが存在しない場合はデフォルト値を使用します。
    """
    try:
        with open("last_target_values.txt", "r", encoding="utf-8") as file:
            lines = file.readlines()
            if len(lines) >= 3:  # 少なくとも3行あることを確認
                target1 = int(lines[0].strip())
                target3 = int(lines[1].strip())
                target5 = int(lines[2].strip())
                return target1, target3, target5
            else:
                logging.warning("last_target_values.txtの行数が不足しています")
                return 0, 0, 0
    except Exception as e:
        logging.error(f"最終目標値ファイルの読み込みエラー: {e}")
        return 0, 0, 0

def save_last_target_values(target1, target3, target5):
    """
    最終目標値をテキストファイルに保存します。
    """
    try:
        with open("last_target_values.txt", "w", encoding="utf-8") as file:
            file.write(f"{target1}\n{target3}\n{target5}")
        logging.info(f"最終目標値を保存: 最低={target1}, 第二={target3}, 最高={target5}")
        return True
    except Exception as e:
        logging.error(f"最終目標値の保存でエラー: {str(e)}")
        return False

def get_target_values_from_storage(driver):
    """
    HTMLの属性から更新された目標値を取得し、対応するテキストファイルに保存します。
    """
    try:
        # LocalStorageの値を確認
        # 週間目標値の取得と保存
        week_a = driver.execute_script("return localStorage.getItem('week_a_target')")
        week_k = driver.execute_script("return localStorage.getItem('week_k_target')")
        if week_a and week_k:
            save_target_values(int(week_a), int(week_k))

        # 月間目標値の取得と保存
        month_a = driver.execute_script("return localStorage.getItem('month_a_target')")
        month_k = driver.execute_script("return localStorage.getItem('month_k_target')")
        if month_a and month_k:
            save_month_target_values(int(month_a), int(month_k))

        # 最終目標値の取得と保存
        last_target1 = driver.execute_script("return localStorage.getItem('last_target1')")
        last_target3 = driver.execute_script("return localStorage.getItem('last_target3')")
        last_target5 = driver.execute_script("return localStorage.getItem('last_target5')")
        if last_target1 and last_target3 and last_target5:
            save_last_target_values(int(last_target1), int(last_target3), int(last_target5))

        return True
    except Exception as e:
        logging.error(f"目標値の取得でエラー: {str(e)}")
        return False

def check_for_target_updates(driver):
    """
    HTML属性から目標値の即時更新をチェックします。
    更新があった場合は、すぐにテキストファイルに保存し、HTMLを再生成します。
    """
    try:
        # 更新フラグの確認
        is_updated = driver.execute_script("return document.body.getAttribute('data-target-updated')")
        if is_updated == 'true':
            # 更新タイプの取得
            target_type = driver.execute_script("return document.body.getAttribute('data-target-type')")
            
            if target_type == 'week':
                # 週間目標値の更新
                a_target = int(driver.execute_script("return document.body.getAttribute('data-target-a_target')"))
                k_target = int(driver.execute_script("return document.body.getAttribute('data-target-k_target')"))
                save_target_values(a_target, k_target)
                logging.info(f"週間目標値を即時更新: A={a_target}, K={k_target}")
            
            elif target_type == 'month':
                # 月間目標値の更新
                a_target = int(driver.execute_script("return document.body.getAttribute('data-target-a_target')"))
                k_target = int(driver.execute_script("return document.body.getAttribute('data-target-k_target')"))
                save_month_target_values(a_target, k_target)
                logging.info(f"月間目標値を即時更新: A={a_target}, K={k_target}")
            
            elif target_type == 'last':
                # 最終目標値の更新
                target1 = int(driver.execute_script("return document.body.getAttribute('data-target-target1')"))
                target3 = int(driver.execute_script("return document.body.getAttribute('data-target-target3')"))
                target5 = int(driver.execute_script("return document.body.getAttribute('data-target-target5')"))
                save_last_target_values(target1, target3, target5)
                logging.info(f"最終目標値を即時更新: 最低={target1}, 第二={target3}, 最高={target5}")
            
            # 更新フラグをリセット
            driver.execute_script("document.body.removeAttribute('data-target-updated')")
            return True
        
        return False
    except Exception as e:
        logging.error(f"目標値の即時更新チェックでエラー: {str(e)}")
        return False

def create_html_content(a_total, k_total, is_week=True, is_last=False):
    """
    HTMLコンテンツを生成します。
    is_week: Trueの場合はweek_result.html用、Falseの場合はmonth_result.html用
    is_last: Trueの場合はlast_result.html用
    """
    # 目標値を読み取り（週間/月間/最終で分ける）
    if is_week:
        a_target, k_target = read_target_values()
    elif is_last:
        target1, target3, target5 = read_last_target_values()  # 3つの値を正しく受け取る
        a_target = target1  # 一時的な対応として最初の値をa_targetに使用
        k_target = 0  # 一時的な対応として0をk_targetに使用
    else:
        a_target, k_target = read_month_target_values()
    
    ak_total = a_total + k_total
    
    if is_week or not is_last:
        remainder_a = a_target - a_total
        remainder_k = k_target - k_total
    
    # 現在の週番号または月を取得
    current_date = datetime.now()
    current_month = current_date.month
    week_number = calculate_week_number()
    
    # 画像の相対パスを使用
    achievement_img_path = "achievement.png"

    # タイトル設定
    if is_week:
        title = f"{week_number}w目標！"
    elif is_last:
        title = "最終目標！"
    else:
        title = f"{current_month}月目標！"
    
    # モーダルコンテンツの設定
    if is_week:
        modal_content = f"""
        <div class="modal-content">
            <h2>週間目標値の変更</h2>
            <div class="input-group">
                <label>A目標数</label>
                <input type="number" id="week_a_target" value="{a_target}">
            </div>
            <div class="input-group">
                <label>K目標数</label>
                <input type="number" id="week_k_target" value="{k_target}">
            </div>
            <button onclick="saveWeekTargets()">変更</button>
        </div>"""
    elif is_last:
        target1, target3, target5 = read_last_target_values()
        modal_content = f"""
        <div class="modal-content">
            <h2>最終目標値の変更</h2>
            <div class="input-group">
                <label>最低目標</label>
                <input type="number" id="last_target1" value="{target1}">
            </div>
            <div class="input-group">
                <label>第二目標</label>
                <input type="number" id="last_target3" value="{target3}">
            </div>
            <div class="input-group">
                <label>最高目標</label>
                <input type="number" id="last_target5" value="{target5}">
            </div>
            <button onclick="saveLastTargets()">変更</button>
        </div>"""
    else:
        modal_content = f"""
        <div class="modal-content">
            <h2>月間目標値の変更</h2>
            <div class="input-group">
                <label>A目標数</label>
                <input type="number" id="month_a_target" value="{a_target}">
            </div>
            <div class="input-group">
                <label>K目標数</label>
                <input type="number" id="month_k_target" value="{k_target}">
            </div>
            <button onclick="saveMonthTargets()">変更</button>
        </div>"""
    
    # 右下セクションの内容を条件分岐
    if is_last:
        # 最終目標用の表示（①〇〇件まで：➁〇〇件, ③〇〇件まで：④〇〇件, ⑤〇〇件まで：⑥〇〇件）
        target1, target3, target5 = read_last_target_values()  # 3つの値を正しく受け取る
        target2 = target1 - ak_total
        target4 = target3 - ak_total
        target6 = target5 - ak_total
        
        # 目標達成時の★と画像表示
        if target2 <= 0:
            display_target2 = f'<span style="white-space: nowrap;">★{abs(target2)}件<img src="{achievement_img_path}" style="width:7vh; height:7vh; vertical-align: middle;"></span>'
        else:
            display_target2 = f"{target2}件"
            
        if target4 <= 0:
            display_target4 = f'<span style="white-space: nowrap;">★{abs(target4)}件<img src="{achievement_img_path}" style="width:7vh; height:7vh; vertical-align: middle;"></span>'
        else:
            display_target4 = f"{target4}件"
            
        if target6 <= 0:
            display_target6 = f'<span style="white-space: nowrap;">★{abs(target6)}件<img src="{achievement_img_path}" style="width:7vh; height:7vh; vertical-align: middle;"></span>'
        else:
            display_target6 = f"{target6}件"
        
        target_section = f"""
        <div style="height: 100%; display: flex; flex-direction: column; justify-content: center; padding: 1vh; box-sizing: border-box;">
            <div style="text-align: center; margin: 0.5vh 0;">
                <div style="font-size: 7vh; white-space: nowrap;">{target1}件まで: <span class="important" style="font-size: 8vh;">{display_target2}</span></div>
            </div>
            <div style="text-align: center; margin: 0.5vh 0;">
                <div style="font-size: 7vh; white-space: nowrap;">{target3}件まで: <span class="important" style="font-size: 8vh;">{display_target4}</span></div>
            </div>
            <div style="text-align: center; margin: 0.5vh 0;">
                <div style="font-size: 7vh; white-space: nowrap;">{target5}件まで: <span class="important" style="font-size: 8vh;">{display_target6}</span></div>
            </div>
        </div>"""
    else:
        if remainder_a <= 0:
            display_a = f'<span style="white-space: nowrap;">★{abs(remainder_a)}件<img src="{achievement_img_path}" style="width:7vh; height:7vh; vertical-align: middle;"></span>'
        else:
            display_a = f"{remainder_a}件"
            
        if remainder_k <= 0:
            display_k = f'<span style="white-space: nowrap;">★{abs(remainder_k)}件<img src="{achievement_img_path}" style="width:7vh; height:7vh; vertical-align: middle;"></span>'
        else:
            display_k = f"{remainder_k}件"
        
        target_section = f"""
        <div style="height: 100%; display: flex; flex-direction: column; justify-content: center; padding: 1vh; box-sizing: border-box;">
            <div style="text-align: center; margin: 0.5vh 0;">
                <div style="font-size: 7vh; white-space: nowrap;">{a_target}件まで: <span class="important" style="font-size: 8vh;">{display_a}</span></div>
            </div>
            <div style="text-align: center; margin: 0.5vh 0;">
                <div style="font-size: 7vh; white-space: nowrap;">{k_target}件まで: <span class="important" style="font-size: 8vh;">{display_k}</span></div>
            </div>
        </div>"""

    # タイムスタンプを追加（キャッシュ回避用）
    timestamp = int(time.time())
    
    return f"""<html>
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="60">
  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="0">
  <title>{title}</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      width: 100vw;
      height: 100vh;
      font-family: 'Arial', sans-serif;
      background: linear-gradient(to bottom right, #f8f9fa, #e9ecef);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
      overflow: hidden;
      box-sizing: border-box;
    }}
    h1 {{
      font-size: 12vh;
      margin: 1vh 0;
      text-shadow: 2px 2px 4px #bdc3c7;
      white-space: nowrap;
    }}
    .container {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      grid-template-rows: 1fr 1fr;
      gap: 1vh;
      width: 98%;
      height: 80%;
      margin-bottom: 1vh;
      box-sizing: border-box;
    }}
    .section {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      border: 0.4vh solid #3498db;
      border-radius: 2vh;
      background-color: #ffffff;
      box-shadow: 0 0.5vh 1vh rgba(0,0,0,0.1);
      padding: 1vh;
      box-sizing: border-box;
      overflow: hidden;
    }}
    .section p {{
      font-size: 10vh;
      margin: 0.5vh 0;
      text-align: center;
      line-height: 1.2;
      white-space: nowrap;
    }}
    .important {{
      color: #e74c3c;
      font-weight: bold;
      font-size: 12vh;
      white-space: nowrap;
    }}
    .small {{
      padding: 0.5vh;
    }}
    img {{
      width: 8vh;
      height: 8vh;
      margin-left: 0.5vh;
      vertical-align: middle;
    }}
    .timestamp {{
      display: none;
    }}
    
    /* モーダル関連のスタイル */
    .target-change-btn {{
      position: fixed;
      top: 20px;
      right: 20px;
      padding: 10px 20px;
      background-color: #3498db;
      color: white;
      border: none;
      border-radius: 5px;
      cursor: pointer;
      font-size: 16px;
      z-index: 1000;
    }}
    .target-change-btn:hover {{
      background-color: #2980b9;
    }}
    .modal {{
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background-color: rgba(0,0,0,0.5);
      z-index: 1001;
    }}
    .modal-content {{
      position: relative;
      background-color: white;
      margin: 15% auto;
      padding: 20px;
      width: 80%;
      max-width: 500px;
      border-radius: 5px;
    }}
    .input-group {{
      margin: 15px 0;
    }}
    .input-group label {{
      display: block;
      margin-bottom: 5px;
    }}
    .input-group input {{
      width: 100%;
      padding: 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }}
    .modal-content button {{
      width: 100%;
      padding: 10px;
      background-color: #3498db;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }}
    .modal-content button:hover {{
      background-color: #2980b9;
    }}
  </style>
</head>
<body>
  <button class="target-change-btn" onclick="showModal()">目標数変更</button>
  <div id="targetModal" class="modal">
    {modal_content}
  </div>
  <h1>{title}</h1>
  <div class="container">
    <div class="section">
      <p>A残込　<span class="important">{a_total}</span> 件</p>
    </div>
    <div class="section">
      <p>K残込　<span class="important">{k_total}</span> 件</p>
    </div>
    <div class="section">
      <p>AK残込　<span class="important">{ak_total}</span> 件</p>
    </div>
    <div class="section small">
      {target_section}
    </div>
  </div>
  <div class="timestamp">{timestamp}</div>
  <script>
    function showModal() {{
      document.getElementById('targetModal').style.display = 'block';
    }}
    
    // Seleniumからのアクセスを容易にするためにカスタム属性を設定する関数
    function setTargetAttributes(type, values) {{
      document.body.setAttribute('data-target-type', type);
      for (let key in values) {{
        document.body.setAttribute('data-target-' + key, values[key]);
      }}
      document.body.setAttribute('data-target-updated', 'true');
    }}
    
    function saveWeekTargets() {{
      const aTarget = document.getElementById('week_a_target').value;
      const kTarget = document.getElementById('week_k_target').value;
      
      // LocalStorageに保存
      localStorage.setItem('week_a_target', aTarget);
      localStorage.setItem('week_k_target', kTarget);
      
      // テキストファイルへの即時更新のためにカスタム属性を設定
      setTargetAttributes('week', {{
        a_target: aTarget,
        k_target: kTarget
      }});
      
      // モーダルを閉じる
      document.getElementById('targetModal').style.display = 'none';
    }}
    
    function saveMonthTargets() {{
      const aTarget = document.getElementById('month_a_target').value;
      const kTarget = document.getElementById('month_k_target').value;
      
      // LocalStorageに保存
      localStorage.setItem('month_a_target', aTarget);
      localStorage.setItem('month_k_target', kTarget);
      
      // テキストファイルへの即時更新のためにカスタム属性を設定
      setTargetAttributes('month', {{
        a_target: aTarget,
        k_target: kTarget
      }});
      
      // モーダルを閉じる
      document.getElementById('targetModal').style.display = 'none';
    }}
    
    function saveLastTargets() {{
      const target1 = document.getElementById('last_target1').value;
      const target3 = document.getElementById('last_target3').value;
      const target5 = document.getElementById('last_target5').value;
      
      // LocalStorageに保存
      localStorage.setItem('last_target1', target1);
      localStorage.setItem('last_target3', target3);
      localStorage.setItem('last_target5', target5);
      
      // テキストファイルへの即時更新のためにカスタム属性を設定
      setTargetAttributes('last', {{
        target1: target1,
        target3: target3,
        target5: target5
      }});
      
      // モーダルを閉じる
      document.getElementById('targetModal').style.display = 'none';
    }}
    
    window.onclick = function(event) {{
      const modal = document.getElementById('targetModal');
      if (event.target == modal) {{
        modal.style.display = 'none';
      }}
    }}
  </script>
</body>
</html>"""

def auto_login_apollo(driver):
    """
    Apolloサイトへの自動ログイン
    """
    try:
        driver.get("https://apollo-scedure77.net/LOGIN/")
        time.sleep(2)

        # ログインフォーム要素の取得と入力
        account_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "account"))
        )
        pass_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "pass"))
        )

        account_input.send_keys("imurayap33")
        pass_input.send_keys("imimp0633")

        # ログインボタンをクリック
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'p.login > input[name="Submit"]'))
        )
        login_button.click()
        time.sleep(5)
        return True
    except TimeoutException:
        print("Apolloログインでエラーが発生しました。")
        return False

def auto_login_tomato(driver):
    """
    Tomatoサイトへの自動ログイン
    """
    try:
        driver.get("https://tomato-systemprograms1455.com/LOGIN/")
        time.sleep(2)

        # ログインフォーム要素の取得と入力
        account_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "account"))
        )
        pass_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "pass"))
        )

        account_input.send_keys("imurayap33")
        pass_input.send_keys("imimp0633")

        # ログインボタンをクリック
        login_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'p.login > input[type="image"]'))
        )
        login_button.click()
        time.sleep(5)
        return True
    except TimeoutException:
        print("Tomatoログインでエラーが発生しました。")
        return False

def create_headless_options():
    """
    ヘッドレスモード用のChromeオプションを作成
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument('--remote-debugging-port=9222')
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    
    # Chrome User Dataディレクトリの設定
    user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    return options

def create_display_options():
    """
    表示用ブラウザのオプションを設定
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")  # 最大化して起動
    options.add_argument("--disable-infobars")  # 情報バーを非表示
    options.add_argument("--disable-extensions")  # 拡張機能を無効化
    options.add_argument("--disable-popup-blocking")  # ポップアップブロックを無効化
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # キャッシュを無効化
    options.add_argument("--disable-application-cache")
    options.add_argument("--disable-cache")
    options.add_argument("--disable-offline-load-stale-cache")
    options.add_argument("--disk-cache-size=0")
    
    # Chrome User Dataディレクトリの設定
    user_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_user_data_display")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    
    # プリファレンス設定
    prefs = {
        "profile.default_content_setting_values.notifications": 2,  # 通知を無効化
        "profile.default_content_settings.popups": 0,  # ポップアップを許可
        "download.default_directory": os.getcwd(),  # ダウンロードディレクトリを現在のディレクトリに設定
        "browser.cache.disk.enable": False,  # ディスクキャッシュを無効化
        "browser.cache.memory.enable": False,  # メモリキャッシュを無効化
        "browser.cache.offline.enable": False,  # オフラインキャッシュを無効化
        "network.http.use-cache": False,  # HTTPキャッシュを無効化
        "profile.default_content_setting_values.media_stream": 2,  # メディアストリームを無効化
        "profile.default_content_setting_values.geolocation": 2,  # 位置情報を無効化
        "profile.default_content_setting_values.automatic_downloads": 1,  # 自動ダウンロードを許可
        "profile.default_content_setting_values.mixed_script": 1,  # 混合スクリプトを許可
        "profile.default_content_setting_values.media_stream_mic": 2,  # マイクを無効化
        "profile.default_content_setting_values.media_stream_camera": 2  # カメラを無効化
    }
    options.add_experimental_option("prefs", prefs)
    
    return options

def get_current_week_range():
    """
    現在の日付から、その週の月曜日から日曜日の範囲を返します。
    """
    current_date = datetime.now()
    # 現在の曜日を取得（0=月曜日, 6=日曜日）
    current_weekday = current_date.weekday()
    # その週の月曜日を計算
    monday = current_date - timedelta(days=current_weekday)
    # その週の日曜日を計算
    sunday = monday + timedelta(days=6)
    return monday, sunday

def get_week_url():
    """
    現在の週のApolloカレンダーURLを生成します。
    """
    monday, _ = get_current_week_range()
    return f"https://apollo-scedure77.net/CAL/week.php?w={monday.strftime('%Y-%m-%d')}"

def get_month_url():
    """
    現在の月の1日のApolloカレンダーURLを生成します。
    """
    current_date = datetime.now()
    first_day = current_date.replace(day=1)
    return f"https://apollo-scedure77.net/CAL/week.php?w={first_day.strftime('%Y-%m-%d')}"

def get_chrome_driver_path():
    """
    実行環境に応じて適切なChromeDriverのパスを取得します。
    """
    try:
        # 一時ディレクトリを作成
        temp_dir = tempfile.mkdtemp()
        logging.info(f"一時ディレクトリを作成: {temp_dir}")
        
        # ChromeDriverをダウンロード
        chromedriver_path = ChromeDriverManager().install()
        logging.info(f"ChromeDriverをダウンロード: {chromedriver_path}")
        
        if getattr(sys, 'frozen', False):
            # 実行ファイルとして実行されている場合
            # ダウンロードしたドライバーを一時ディレクトリにコピー
            target_path = os.path.join(temp_dir, "chromedriver.exe")
            shutil.copy2(chromedriver_path, target_path)
            logging.info(f"ChromeDriverを一時ディレクトリにコピー: {target_path}")
            return target_path
        else:
            # 通常のPythonスクリプトとして実行されている場合
            return chromedriver_path
    except Exception as e:
        logging.error(f"ChromeDriverの設定中にエラーが発生: {str(e)}")
        raise

def refresh_browser_session(driver):
    """
    ブラウザセッションを更新します。
    セッションが無効な場合は新しいセッションを作成します。
    """
    try:
        # 現在のURLを取得してセッションの状態を確認
        current_url = driver.current_url
        return True
    except Exception as e:
        logging.error(f"ブラウザセッションの更新に失敗: {e}")
        return False

# -------------------------------
# メイン処理開始
# -------------------------------

if __name__ == "__main__":
    try:
        logging.info("アプリケーション開始")
        
        # ChromeDriverのパスを取得
        chromedriver_path = get_chrome_driver_path()
        logging.info(f"ChromeDriverパス: {chromedriver_path}")
        
        # apolloサイト用ドライバの起動＆ログイン
        headless_options = create_headless_options()
        
        logging.info("ChromeDriverサービス作成開始")
        extraction_service = ChromeService(chromedriver_path)
        extraction_driver = webdriver.Chrome(service=extraction_service, options=headless_options)
        logging.info("ChromeDriver初期化完了")

        if not auto_login_apollo(extraction_driver):
            logging.error("Apolloログイン失敗")
            extraction_driver.quit()
            exit(1)
        logging.info("Apolloログイン成功")

        # 現在の週のURLを取得
        target_url = get_week_url()
        extraction_driver.get(target_url)
        logging.info(f"Apollo URL取得成功: {target_url}")

        # tomatoサイト用ドライバの起動＆ログイン
        tomato_service = Service(chromedriver_path)
        tomato_driver = webdriver.Chrome(service=tomato_service, options=headless_options)
        logging.info("Tomato ChromeDriver初期化完了")

        if not auto_login_tomato(tomato_driver):
            logging.error("Tomatoログイン失敗")
            extraction_driver.quit()
            tomato_driver.quit()
            exit(1)
        logging.info("Tomatoログイン成功")

        # ④ HTML表示用ドライバの設定（非ヘッドレスモード）
        week_filename = "week_result.html"
        month_filename = "month_result.html"
        last_filename = "last_result.html"
        week_path = "file:///" + os.path.abspath(week_filename)
        month_path = "file:///" + os.path.abspath(month_filename)
        last_path = "file:///" + os.path.abspath(last_filename)

        display_service = Service(chromedriver_path)
        display_options = create_display_options()
        display_driver = webdriver.Chrome(service=display_service, options=display_options)
        logging.info("表示用ChromeDriver初期化完了")

        # 初期データの取得
        try:
            # 週間データ取得用URL
            target_url = get_week_url()
            extraction_driver.get(target_url)
            time.sleep(3)
            week_a_total = recalc_total_week()
            time.sleep(2)

            # 月間データ取得用URL
            month_url = get_month_url()
            extraction_driver.get(month_url)
            time.sleep(3)
            month_a_total = recalc_total_month()
            time.sleep(2)

            # tomatoサイトからデータ取得（週間用）
            current_date = datetime.now()
            url_ueno = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=13&y={current_date.year}&m={current_date.month:02d}#cal"
            url_kyoto = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=10&y={current_date.year}&m={current_date.month:02d}#cal"
            
            # 週間表示用のK残込を取得
            _, _, UenoNet = get_oguchi_value(tomato_driver, url_ueno)
            time.sleep(2)
            _, _, KyotoNet = get_oguchi_value(tomato_driver, url_kyoto)
            time.sleep(2)
            week_k_total = UenoNet + KyotoNet

            # 月間表示用のK残込を取得
            _, _, UenoNetMonth = get_oguchi_value_month(tomato_driver, url_ueno)
            time.sleep(2)
            _, _, KyotoNetMonth = get_oguchi_value_month(tomato_driver, url_kyoto)
            time.sleep(2)
            month_k_total = UenoNetMonth + KyotoNetMonth

            # 最終目標表示用のデータ（月間と同じデータを使用）
            last_a_total = month_a_total
            last_k_total = month_k_total

            # 3つのHTMLファイルを作成
            with open(week_filename, "w", encoding="utf-8") as file:
                html_content = create_html_content(week_a_total, week_k_total, is_week=True)
                file.write(html_content)
            with open(month_filename, "w", encoding="utf-8") as file:
                html_content = create_html_content(month_a_total, month_k_total, is_week=False)
                file.write(html_content)
            with open(last_filename, "w", encoding="utf-8") as file:
                html_content = create_html_content(last_a_total, last_k_total, is_week=False, is_last=True)
                file.write(html_content)
            
            logging.info("初期データ取得・HTML作成完了")
        except Exception as e:
            logging.error(f"初期データ取得でエラー発生: {str(e)}")
            # エラー時は初期表示用HTMLを作成
            initial_html = "<html><body><h1>データ取得中にエラーが発生しました...</h1></body></html>"
            for filename in [week_filename, month_filename, last_filename]:
                with open(filename, "w", encoding="utf-8") as file:
                    file.write(initial_html)

        # 最初のタブで週間目標を表示
        display_driver.get(week_path)
        time.sleep(2)

        # 新しいタブを開く（月間目標用）
        display_driver.execute_script("window.open('', '_blank');")
        time.sleep(2)

        # 新しいタブに切り替えて月間目標を表示
        display_driver.switch_to.window(display_driver.window_handles[-1])
        display_driver.get(month_path)
        time.sleep(2)

        # 新しいタブを開く（最終目標用）
        display_driver.execute_script("window.open('', '_blank');")
        time.sleep(2)

        # 新しいタブに切り替えて最終目標を表示
        display_driver.switch_to.window(display_driver.window_handles[-1])
        display_driver.get(last_path)
        time.sleep(2)

        logging.info("初期HTML表示完了")
        logging.info(f"タブの数: {len(display_driver.window_handles)}")

        # ⑤ 定期更新ループ（1分毎）
        while True:
            try:
                # ブラウザセッションの状態を確認
                if not refresh_browser_session(display_driver):
                    logging.warning("ブラウザセッションが無効です。新しいセッションを作成します。")
                    display_driver.quit()
                    display_driver = webdriver.Chrome(service=display_service, options=display_options)
                    # 各タブを再作成
                    display_driver.get(week_path)
                    time.sleep(2)
                    display_driver.execute_script("window.open('', '_blank');")
                    time.sleep(2)
                    display_driver.switch_to.window(display_driver.window_handles[-1])
                    display_driver.get(month_path)
                    time.sleep(2)
                    display_driver.execute_script("window.open('', '_blank');")
                    time.sleep(2)
                    display_driver.switch_to.window(display_driver.window_handles[-1])
                    display_driver.get(last_path)
                    time.sleep(2)

                # 週間データ取得用URL
                target_url = get_week_url()
                extraction_driver.get(target_url)
                time.sleep(3)
                extraction_driver.refresh()
                time.sleep(3)
                
                # 週間表示用のデータ取得
                week_a_total = recalc_total_week()
                time.sleep(2)

                # 月間データ取得用URL
                month_url = get_month_url()
                extraction_driver.get(month_url)
                time.sleep(3)
                extraction_driver.refresh()
                time.sleep(3)
                
                # 月間表示用のデータ取得
                month_a_total = recalc_total_month()
                time.sleep(2)

                # 現在の年月を更新
                current_date = datetime.now()
                target_year = current_date.year
                target_month = current_date.month

                # tomatoサイトのURLを更新
                url_ueno = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=13&y={target_year}&m={target_month:02d}#cal"
                url_kyoto = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=10&y={target_year}&m={target_month:02d}#cal"

                # 週間表示用のK残込を取得
                _, _, UenoNet = get_oguchi_value(tomato_driver, url_ueno)
                time.sleep(2)
                _, _, KyotoNet = get_oguchi_value(tomato_driver, url_kyoto)
                time.sleep(2)
                week_k_total = UenoNet + KyotoNet

                # 月間表示用のK残込を取得
                _, _, UenoNetMonth = get_oguchi_value_month(tomato_driver, url_ueno)
                time.sleep(2)
                _, _, KyotoNetMonth = get_oguchi_value_month(tomato_driver, url_kyoto)
                time.sleep(2)
                month_k_total = UenoNetMonth + KyotoNetMonth

                # 最終目標表示用のデータ（月間と同じデータを使用）
                last_a_total = month_a_total
                last_k_total = month_k_total

                # 現在のタブのURLとハンドルを記録
                current_handle = display_driver.current_window_handle
                current_url = display_driver.current_url
                
                # 即時更新をチェックする（モーダルから変更された場合の処理）
                if check_for_target_updates(display_driver):
                    # 目標値をすぐに再読み込み
                    week_a_target, week_k_target = read_target_values()
                    month_a_target, month_k_target = read_month_target_values()
                    last_target1, last_target3, last_target5 = read_last_target_values()
                    
                    # HTMLファイルを即時更新
                    with open(week_filename, "w", encoding="utf-8") as file:
                        html_content = create_html_content(week_a_total, week_k_total, is_week=True)
                        file.write(html_content)
                    with open(month_filename, "w", encoding="utf-8") as file:
                        html_content = create_html_content(month_a_total, month_k_total, is_week=False)
                        file.write(html_content)
                    with open(last_filename, "w", encoding="utf-8") as file:
                        html_content = create_html_content(last_a_total, last_k_total, is_week=False, is_last=True)
                        file.write(html_content)
                    
                    # タイムスタンプを追加してキャッシュを回避
                    timestamp = int(time.time())
                    
                    # 現在のタブのURLに基づいて、適切なHTMLファイルを更新
                    if "week_result.html" in current_url:
                        display_driver.get(f"{week_path}?t={timestamp}")
                    elif "month_result.html" in current_url:
                        display_driver.get(f"{month_path}?t={timestamp}")
                    elif "last_result.html" in current_url:
                        display_driver.get(f"{last_path}?t={timestamp}")
                    else:
                        # URLが不明な場合は、現在のタブのURLを維持
                        display_driver.get(current_url)
                
                # 通常の定期更新（1分ごと）
                # 現在のタブから目標値を更新（localStorageを使用）
                get_target_values_from_storage(display_driver)
                
                # 目標値を再読み込み（全てのタブでの更新を反映）
                week_a_target, week_k_target = read_target_values()
                month_a_target, month_k_target = read_month_target_values()
                last_target1, last_target3, last_target5 = read_last_target_values()
                
                # HTML更新前の待機
                time.sleep(1)
                
                # タイムスタンプを追加してキャッシュを回避
                timestamp = int(time.time())
                
                # HTMLファイルを更新（最新の目標値を反映）
                with open(week_filename, "w", encoding="utf-8") as file:
                    html_content = create_html_content(week_a_total, week_k_total, is_week=True)
                    file.write(html_content)
                with open(month_filename, "w", encoding="utf-8") as file:
                    html_content = create_html_content(month_a_total, month_k_total, is_week=False)
                    file.write(html_content)
                with open(last_filename, "w", encoding="utf-8") as file:
                    html_content = create_html_content(last_a_total, last_k_total, is_week=False, is_last=True)
                    file.write(html_content)
                time.sleep(2)
                
                # 現在のタブのURLに基づいて、適切なHTMLファイルを更新
                if "week_result.html" in current_url:
                    display_driver.get(f"{week_path}?t={timestamp}")
                elif "month_result.html" in current_url:
                    display_driver.get(f"{month_path}?t={timestamp}")
                elif "last_result.html" in current_url:
                    display_driver.get(f"{last_path}?t={timestamp}")
                else:
                    # URLが不明な場合は、現在のタブのURLを維持
                    display_driver.get(current_url)
                
                # 週間・月間・最終目標のデータを常に表示
                week_ak_total = week_a_total + week_k_total
                week_remainder = 890 - week_ak_total
                month_ak_total = month_a_total + month_k_total
                month_remainder = 890 - month_ak_total
                last_ak_total = last_a_total + last_k_total
                
                # 目標値を読み込み
                week_a_target, week_k_target = read_target_values()
                month_a_target, month_k_target = read_month_target_values()
                last_target1, last_target3, last_target5 = read_last_target_values()
                
                # 目標までの残り件数
                week_a_remainder = week_a_target - week_a_total
                week_k_remainder = week_k_target - week_k_total
                month_a_remainder = month_a_target - month_a_total
                month_k_remainder = month_k_target - month_k_total
                last_target2 = last_target1 - last_ak_total
                last_target4 = last_target3 - last_ak_total
                
                print(f"週間データ: A残込 = {week_a_total}, K残込 = {week_k_total}, AK残込 = {week_ak_total}")
                print(f"週間目標: A目標 = {week_a_target}件まで残り{week_a_remainder}件, K目標 = {week_k_target}件まで残り{week_k_remainder}件")
                print(f"月間データ: A残込 = {month_a_total}, K残込 = {month_k_total}, AK残込 = {month_ak_total}")
                print(f"月間目標: A目標 = {month_a_target}件まで残り{month_a_remainder}件, K目標 = {month_k_target}件まで残り{month_k_remainder}件")
                print(f"最終目標: 目標1 = {last_target1}件まで残り{last_target2}件, 目標2 = {last_target3}件まで残り{last_target4}件, 目標3 = {last_target5}件まで残り{last_target5 - last_ak_total}件")
                print("-" * 80)
                
                # 次の更新までの待機（55秒に短縮し、処理時間の余裕を持たせる）
                time.sleep(1)
            except Exception as e:
                logging.error(f"メインループでエラー発生: {str(e)}")
                time.sleep(60)  # エラー時も1分待機してから再試行
    except Exception as e:
        logging.error(f"重大なエラーが発生: {str(e)}")
        messagebox.showerror("エラー", f"アプリケーションの起動に失敗しました。\nエラー内容: {str(e)}\n\nログファイル {log_file} を確認してください。")