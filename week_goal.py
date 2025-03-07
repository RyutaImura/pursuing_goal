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
    表示用のChromeオプションを作成（非ヘッドレスモード）
    """
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--start-maximized")  # ウィンドウを最大化
    
    # Chrome User Dataディレクトリの設定
    user_data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "chrome_user_data_display")
    options.add_argument(f"--user-data-dir={user_data_dir}")
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

        # ③ tomatoサイト用ドライバの起動＆ログイン
        tomato_service = Service(chromedriver_path)
        tomato_driver = webdriver.Chrome(service=tomato_service, options=headless_options)
        logging.info("Tomato ChromeDriver初期化完了")

        if not auto_login_tomato(tomato_driver):
            logging.error("Tomatoログイン失敗")
            extraction_driver.quit()
            tomato_driver.quit()
            exit(1)
        logging.info("Tomatoログイン成功")

        # 現在の年月を取得
        current_date = datetime.now()
        target_year = current_date.year
        target_month = current_date.month

        # スクレイピング対象の月別ページURL（上野用、京都用）
        url_ueno = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=13&y={target_year}&m={target_month:02d}#cal"
        url_kyoto = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=10&y={target_year}&m={target_month:02d}#cal"

        # ④ HTML表示用ドライバの設定（非ヘッドレスモード）
        html_filename = "clinic_result.html"
        abs_html_path = "file:///" + os.path.abspath(html_filename)
        display_service = Service(chromedriver_path)
        display_options = create_display_options()
        display_driver = webdriver.Chrome(service=display_service, options=display_options)
        logging.info("表示用ChromeDriver初期化完了")

        initial_html = "<html><body><p>初期表示中...</p></body></html>"
        with open(html_filename, "w", encoding="utf-8") as file:
            file.write(initial_html)
        display_driver.get(abs_html_path)
        logging.info("初期HTML表示完了")

        # ⑤ 定期更新ループ（1分毎）
        while True:
            try:
                # 現在の週のURLを再取得（週が変わった場合に対応）
                target_url = get_week_url()
                extraction_driver.get(target_url)
                time.sleep(3)  # ページ読み込み待機を3秒に延長
                extraction_driver.refresh()
                time.sleep(3)  # リフレッシュ後の待機も3秒に延長
                
                # apolloサイトからデータ取得
                a_total = recalc_total()
                time.sleep(2)  # データ取得後の待機

                # 現在の年月を更新
                current_date = datetime.now()
                target_year = current_date.year
                target_month = current_date.month

                # tomatoサイトのURLを更新
                url_ueno = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=13&y={target_year}&m={target_month:02d}#cal"
                url_kyoto = f"https://tomato-systemprograms1455.com/CAL/monthly.php?c=10&y={target_year}&m={target_month:02d}#cal"

                # tomatoサイト：上野・京都のネット値取得し、K残込とする
                _, _, UenoNet = get_oguchi_value(tomato_driver, url_ueno)
                time.sleep(2)  # 上野データ取得後の待機
                _, _, KyotoNet = get_oguchi_value(tomato_driver, url_kyoto)
                time.sleep(2)  # 京都データ取得後の待機
                k_total = UenoNet + KyotoNet

                # HTML更新前の待機
                time.sleep(1)
                
                # HTML更新
                update_html(a_total, k_total)
                time.sleep(2)  # HTML更新後の待機を2秒に延長
                
                # 表示の更新
                display_driver.get(abs_html_path)
                time.sleep(2)  # 表示更新後の待機を2秒に延長

                ak_total = a_total + k_total
                remainder = 890 - ak_total
                print(f"更新完了: A残込 = {a_total}, K残込 = {k_total}, AK残込 = {ak_total}, 890まで {remainder}件（{time.strftime('%Y-%m-%d %H:%M:%S')}）")
                
                # 次の更新までの待機（55秒に短縮し、処理時間の余裕を持たせる）
                time.sleep(55)
            except Exception as e:
                logging.error(f"メインループでエラー発生: {str(e)}")
                time.sleep(60)  # エラー時も1分待機してから再試行
    except Exception as e:
        logging.error(f"重大なエラーが発生: {str(e)}")
        messagebox.showerror("エラー", f"アプリケーションの起動に失敗しました。\nエラー内容: {str(e)}\n\nログファイル {log_file} を確認してください。")
