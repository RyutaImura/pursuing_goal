from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

def create_display_options():
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--start-maximized")
    return options

def main():
    # ChromeDriverの設定
    service = Service(ChromeDriverManager().install())
    options = create_display_options()
    driver = webdriver.Chrome(service=service, options=options)
    
    # テスト用のHTMLファイルを作成
    html_content = """
    <html>
    <head><title>テストページ</title></head>
    <body><h1>テストページ</h1></body>
    </html>
    """
    
    # 2つのHTMLファイルを作成
    for i, filename in enumerate(["test1.html", "test2.html"]):
        with open(filename, "w", encoding="utf-8") as f:
            f.write(html_content.replace("テストページ", f"テストページ {i+1}"))
    
    # 最初のタブでファイルを開く
    driver.get("file:///" + os.path.abspath("test1.html"))
    time.sleep(1)
    
    # 新しいタブを開く
    driver.execute_script("window.open('', '_blank');")
    time.sleep(1)
    
    # 新しいタブに切り替え
    driver.switch_to.window(driver.window_handles[-1])
    time.sleep(1)
    
    # 2つ目のファイルを表示
    driver.get("file:///" + os.path.abspath("test2.html"))
    
    print("タブを開きました。確認後、このウィンドウを閉じてください。")
    print("タブの数:", len(driver.window_handles))
    
    # 1分ごとに各タブを更新
    try:
        while True:
            for handle in driver.window_handles:
                driver.switch_to.window(handle)
                driver.refresh()
                print(f"タブを更新: {driver.title}")
                time.sleep(1)
            time.sleep(58)  # 残りの時間を待機
    except KeyboardInterrupt:
        driver.quit()

if __name__ == "__main__":
    main() 