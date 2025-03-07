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


def recalc_total():
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

def get_oguchi_value(driver, page_url):
    """
    tomatoサイトの指定ページにアクセスし、
    ① 3/3-9の範囲で実質大口をカウント
    ② <a>タグをすべて調べ、outerHTMLに "fa fa-wifi" が含まれる場合は除外
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
                    
                    # 除外カウント（wifi）
                    if "fa fa-wifi" in outer_html:
                        exclude_count += 1
                        
            except NoSuchElementException:
                continue

    except Exception as e:
        print(f"データ取得エラー (tomato): {e}")

    net_value = total_oguchi - exclude_count
    return total_oguchi, exclude_count, net_value

def read_target_values():
    """
    目標値をテキストファイルから読み取ります。
    ファイルが存在しない場合はデフォルト値を使用します。
    """
    try:
        with open("target_values.txt", "r", encoding="utf-8") as file:
            lines = file.readlines()
            a_target = int(lines[0].strip())
            k_target = int(lines[1].strip())
            return a_target, k_target
    except Exception as e:
        print(f"目標値ファイルの読み込みエラー: {e}")
        # デフォルト値を返す
        return 216, 22

def update_html(a_total, k_total):
    """
    取得した数値をHTMLに埋め込み、ローカルファイルとして保存します。
    表示項目は
      左上：A残込
      右上：K残込
      左下：AK残込 (A残込 + K残込)
      右下：目標値までの残件数（target_values.txtから読み取り）
    なお、目標を超えた場合はマイナス表記ではなく「★」表記に変更し、
    残件数が0以下の場合は、テキストの横に achievement.png の画像を表示します。
    """
    # 目標値を読み取り
    a_target, k_target = read_target_values()
    
    ak_total = a_total + k_total
    remainder_a = a_target - a_total
    remainder_k = k_target - k_total
    week_number = calculate_week_number()

    # 画像の相対パスを使用
    achievement_img_path = "achievement.png"

    if remainder_a <= 0:
        display_a = f'<span style="white-space: nowrap;">★{abs(remainder_a)}件<img src="{achievement_img_path}" style="width:6vw; height:6vw; vertical-align: middle;"></span>'
    else:
        display_a = f"{remainder_a}件"
        
    if remainder_k <= 0:
        display_k = f'<span style="white-space: nowrap;">★{abs(remainder_k)}件<img src="{achievement_img_path}" style="width:6vw; height:6vw; vertical-align: middle;"></span>'
    else:
        display_k = f"{remainder_k}件"

    html_content = f"""<html>
<head>
  <meta charset="UTF-8">
  <title>{week_number}w目標！</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      width: 100%;
      height: 100%;
      font-family: 'Arial', sans-serif;
      background: linear-gradient(to bottom right, #f8f9fa, #e9ecef);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
    }}
    h1 {{
      font-size: 7vw;
      margin: 2vh 0;
      text-shadow: 3px 3px 5px #bdc3c7;
    }}
    .container {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      grid-template-rows: 1fr 1fr;
      gap: 3vw;
      width: 90%;
      height: 70%;
      box-sizing: border-box;
    }}
    .section {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      border: 0.5vw solid #3498db;
      border-radius: 2vw;
      background-color: #ffffff;
      box-shadow: 0px 1vw 2vw rgba(0,0,0,0.1);
    }}
    p {{
      font-size: 5vw;
      margin: 0.5vh 0;
    }}
    .important {{
      font-size: 7vw;
      color: #e74c3c;
      font-weight: bold;
    }}
    /* 右下ボックス用の小さい文字サイズ */
    .small {{
      white-space: nowrap;
    }}
    .small p {{
      font-size: 4vw;
      margin: 0.5vh 0;
      white-space: nowrap;
    }}
    .small .important {{
      font-size: 6vw;
      display: inline-block;
      white-space: nowrap;
    }}
    /* 追加：右下セクションの特別なスタイル */
    .section.small {{
      padding: 1vw;
      overflow: visible;
      white-space: nowrap;
    }}
  </style>
</head>
<body>
  <h1>{week_number}w目標！</h1>
  <div class="container">
    <!-- 左上：A残込 -->
    <div class="section">
      <p>A残込　<span class="important">{a_total}</span> 件</p>
    </div>
    <!-- 右上：K残込 -->
    <div class="section">
      <p>K残込　<span class="important">{k_total}</span> 件</p>
    </div>
    <!-- 左下：AK残込 -->
    <div class="section">
      <p>AK残込　<span class="important">{ak_total}</span> 件</p>
    </div>
    <!-- 右下：目標までの残件数 -->
    <div class="section small">
      <div style="white-space: nowrap;">
        <p>{a_target}件まで: <span class="important">{display_a}</span></p>
        <p>{k_target}件まで: <span class="important">{display_k}</span></p>
      </div>
    </div>
  </div>
</body>
</html>"""
    with open(html_filename, "w", encoding="utf-8") as file:
        file.write(html_content)

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

def create_display_options():
    """
    表示用のChromeオプションを作成（非ヘッドレスモード）
    """
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    return options

def create_headless_options(user_data_suffix=''):
    """
    ヘッドレスモード用のChromeオプションを作成
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
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

def get_current_month_goal():
    """
    現在の月の目標値をテキストファイルから読み取ります。
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
        # デフォルト値を返す（2月の場合）
        return 890, 900

def update_month_goal_html(a_total, k_total):
    """
    月間目標用のHTMLを生成します。
    """
    current_date = datetime.now()
    current_month = current_date.month
    
    # 月間目標値を読み取り
    a_target, k_target = get_current_month_goal()
    
    ak_total = a_total + k_total
    remainder_a = a_target - a_total
    remainder_k = k_target - k_total

    # 画像の相対パスを使用
    achievement_img_path = "achievement.png"

    if remainder_a <= 0:
        display_a = f'<span style="white-space: nowrap;">★{abs(remainder_a)}件<img src="{achievement_img_path}" style="width:6vw; height:6vw; vertical-align: middle;"></span>'
    else:
        display_a = f"{remainder_a}件"
        
    if remainder_k <= 0:
        display_k = f'<span style="white-space: nowrap;">★{abs(remainder_k)}件<img src="{achievement_img_path}" style="width:6vw; height:6vw; vertical-align: middle;"></span>'
    else:
        display_k = f"{remainder_k}件"

    html_content = f"""<html>
<head>
  <meta charset="UTF-8">
  <title>{current_month}月目標！</title>
  <style>
    html, body {{
      margin: 0;
      padding: 0;
      width: 100%;
      height: 100%;
      font-family: 'Arial', sans-serif;
      background: linear-gradient(to bottom right, #f8f9fa, #e9ecef);
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: flex-start;
    }}
    h1 {{
      font-size: 7vw;
      margin: 2vh 0;
      text-shadow: 3px 3px 5px #bdc3c7;
    }}
    .container {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      grid-template-rows: 1fr 1fr;
      gap: 3vw;
      width: 90%;
      height: 70%;
      box-sizing: border-box;
    }}
    .section {{
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      border: 0.5vw solid #3498db;
      border-radius: 2vw;
      background-color: #ffffff;
      box-shadow: 0px 1vw 2vw rgba(0,0,0,0.1);
    }}
    p {{
      font-size: 5vw;
      margin: 0.5vh 0;
    }}
    .important {{
      font-size: 7vw;
      color: #e74c3c;
      font-weight: bold;
    }}
    .small {{
      white-space: nowrap;
    }}
    .small p {{
      font-size: 4vw;
      margin: 0.5vh 0;
      white-space: nowrap;
    }}
    .small .important {{
      font-size: 6vw;
      display: inline-block;
      white-space: nowrap;
    }}
    .section.small {{
      padding: 1vw;
      overflow: visible;
      white-space: nowrap;
    }}
  </style>
</head>
<body>
  <h1>{current_month}月目標！</h1>
  <div class="container">
    <!-- 左上：A残込 -->
    <div class="section">
      <p>A残込　<span class="important">{a_total}</span> 件</p>
    </div>
    <!-- 右上：K残込 -->
    <div class="section">
      <p>K残込　<span class="important">{k_total}</span> 件</p>
    </div>
    <!-- 左下：AK残込 -->
    <div class="section">
      <p>AK残込　<span class="important">{ak_total}</span> 件</p>
    </div>
    <!-- 右下：目標までの残件数 -->
    <div class="section small">
      <div style="white-space: nowrap;">
        <p>{a_target}件まで: <span class="important">{display_a}</span></p>
        <p>{k_target}件まで: <span class="important">{display_k}</span></p>
      </div>
    </div>
  </div>
</body>
</html>"""
    month_html_filename = "month_goal.html"
    with open(month_html_filename, "w", encoding="utf-8") as file:
        file.write(html_content)
    return month_html_filename

def cleanup_chrome_processes():
    """
    実行中のChromeプロセスをクリーンアップします。
    """
    try:
        if os.name == 'nt':  # Windows
            os.system('taskkill /f /im chrome.exe >nul 2>&1')
            os.system('taskkill /f /im chromedriver.exe >nul 2>&1')
        else:  # Linux/Mac
            os.system('pkill -f chrome')
            os.system('pkill -f chromedriver')
        time.sleep(2)  # プロセスが完全に終了するのを待つ
    except Exception as e:
        logging.error(f"Chromeプロセスのクリーンアップ中にエラーが発生: {str(e)}")

# -------------------------------
# メイン処理開始
# -------------------------------

if __name__ == "__main__":
    try:
        logging.info("アプリケーション開始")
        
        # ChromeDriverのパスを取得
        chromedriver_path = get_chrome_driver_path()
        logging.info(f"ChromeDriverパス: {chromedriver_path}")
        
        # 既存のChromeプロセスをクリーンアップ
        cleanup_chrome_processes()
        time.sleep(2)
        
        # 各種ディレクトリの作成とクリーンアップ
        for dir_name in ["chrome_user_data_apollo", "chrome_user_data_tomato", "chrome_user_data_display"]:
            dir_path = os.path.join(os.getcwd(), dir_name)
            if os.path.exists(dir_path):
                try:
                    shutil.rmtree(dir_path)
                    time.sleep(1)
                except Exception as e:
                    logging.warning(f"ディレクトリの削除に失敗: {dir_path}, エラー: {str(e)}")
            os.makedirs(dir_path)
            logging.info(f"ディレクトリを作成: {dir_path}")
        
        time.sleep(2)  # ディレクトリ操作後の待機
        
        # apolloサイト用ドライバの起動＆ログイン
        headless_options_apollo = create_headless_options()
        
        logging.info("ChromeDriverサービス作成開始")
        extraction_service = ChromeService(chromedriver_path)
        extraction_driver = webdriver.Chrome(service=extraction_service, options=headless_options_apollo)
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
        headless_options_tomato = create_headless_options()
        tomato_service = Service(chromedriver_path)
        tomato_driver = webdriver.Chrome(service=tomato_service, options=headless_options_tomato)
        logging.info("Tomato ChromeDriver初期化完了")

        if not auto_login_tomato(tomato_driver):
            logging.error("Tomatoログイン失敗")
            extraction_driver.quit()
            tomato_driver.quit()
            exit(1)
        logging.info("Tomatoログイン成功")

        # HTML表示用ドライバの設定（非ヘッドレスモード）
        html_filename = "clinic_result.html"
        abs_html_path = "file:///" + os.path.abspath(html_filename)
        display_service = Service(chromedriver_path)
        display_options = create_display_options()
        display_driver = webdriver.Chrome(service=display_service, options=display_options)
        
        month_display_service = Service(chromedriver_path)
        month_display_options = create_display_options()
        month_display_driver = webdriver.Chrome(service=month_display_service, options=month_display_options)
        
        logging.info("表示用ChromeDriver初期化完了")

        initial_html = "<html><body><p>初期表示中...</p></body></html>"
        with open(html_filename, "w", encoding="utf-8") as file:
            file.write(initial_html)
        display_driver.get(abs_html_path)
        logging.info("初期HTML表示完了")

        # 定期更新ループ（1分毎）
        while True:
            try:
                # 現在の週のURLを再取得（週が変わった場合に対応）
                target_url = get_week_url()
                extraction_driver.get(target_url)
                time.sleep(2)
                extraction_driver.refresh()
                time.sleep(2)
                
                # apolloサイトからデータ取得
                a_total = recalc_total()
                time.sleep(2)

                # 現在の年月を更新
                current_date = datetime.now()
                target_year = current_date.year
                target_month = current_date.month

                # tomatoサイトのURLを更新
                url_ueno = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=13&y={target_year}&m={target_month:02d}#cal"
                url_kyoto = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=10&y={target_year}&m={target_month:02d}#cal"

                # tomatoサイト：上野・京都のネット値取得
                _, _, UenoNet = get_oguchi_value(tomato_driver, url_ueno)
                time.sleep(2)
                _, _, KyotoNet = get_oguchi_value(tomato_driver, url_kyoto)
                time.sleep(2)
                k_total = UenoNet + KyotoNet

                # 週間目標のHTML更新
                update_html(a_total, k_total)
                time.sleep(1)
                
                # 月間目標のHTML更新
                month_html_filename = update_month_goal_html(a_total, k_total)
                month_abs_html_path = "file:///" + os.path.abspath(month_html_filename)
                
                # 両方の表示を更新
                display_driver.get(abs_html_path)
                month_display_driver.get(month_abs_html_path)
                time.sleep(1)

                ak_total = a_total + k_total
                print(f"更新完了: A残込 = {a_total}, K残込 = {k_total}, AK残込 = {ak_total}（{time.strftime('%Y-%m-%d %H:%M:%S')}）")
                
                time.sleep(55)
            except Exception as e:
                logging.error(f"メインループでエラー発生: {str(e)}")
                time.sleep(60)
    except Exception as e:
        logging.error(f"重大なエラーが発生: {str(e)}")
        messagebox.showerror("エラー", f"アプリケーションの起動に失敗しました。\nエラー内容: {str(e)}\n\nログファイル {log_file} を確認してください。")
