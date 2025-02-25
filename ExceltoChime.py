import pandas as pd
import requests
import schedule
import time
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import base64
import japanize_matplotlib
import os
from openpyxl import load_workbook
import xlwings as xw
import pyfiglet
import random
import sys
# 仮想環境の構築手順
# 仮想環境を作成: python -m venv .venv
# 仮想環境をアクティベート: .venv\Scripts\activate

# 仮想環境のアクティベート確認
# ターミナルに (.venv) と表示されれば成功
#pip freeze > requirements.txt
#pip install -r requirements.txt
#python -m pip install setuptools

# Excelファイルのパスとシート名
excel_file_path = 'book1.xlsx'
sheet_name = 'Sheet1'
excel_file_path2 = 'book2.xlsx'
sheet_name2 = 'Sheet1'
excel_file_path3 = r'Book3.xlsx'
sheet_name3 = 'Sheet1'

# Webhook URL
webhook_url = 'https://webhook-test.com/98b2ff73ac2524023b6f209fc5cb7c7e'
print("システム起動 終了はCtrl+C")

# ローディングアニメーションを表示する関数
def loading_animation():
    loading_text = "Loading"
    for i in range(4):  # 4回繰り返す
        print(loading_text + '.' * i, end='\r')  # カーソルを行の先頭に戻す
        time.sleep(0.5)  # 0.5秒待つ
    print(loading_text + '....')  # 最後に4つのドットを表示

# かっこいい文字を表示
ascii_art = pyfiglet.figlet_format("SYSTEM START", font="slant")
print(ascii_art)
# ローディングアニメーションを表示
loading_animation()

def send_message(message):
    try:
        response = requests.post(webhook_url, json=message)
        if response.status_code == 200:
            print("メッセージが正常に投稿されました")
        else:
            print("メッセージの投稿に失敗しました")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def get_excel_data():
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        
        place_filter = '神奈川'
        df = df[df['部署'] == place_filter]
        
        sum_all = df['数値'].sum()
        
        days_limit = 20
        sum_over_limit = df[df['経過日数'] > days_limit]['数値'].sum()
        print(sum_all, sum_over_limit)

        return sum_all, sum_over_limit
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        return 0, 0

def send_simple_message(content):
    max_size = 3900  # 3.9KB
    # サイズをチェックして調整
    while len(content.encode('utf-8')) > max_size:
        # 行数を減らして調整
        lines = content.split('\n')
        content = '\n'.join(lines[:-1])
    
    message = {
        "Content": "/md\n" + content
    }
    # ここで実際の送信処理を行う
    # print(message)

def send_total_message(sum_all, sum_over_limit):
    content = f"合計は: {sum_all}\n20日以上経過の数値の合計は: {sum_over_limit}"
    send_simple_message(content)

def send_graph_message():
    try:
        image_path = os.path.abspath('histogram.png')
        send_simple_message(f"今後の予想のグラフです。[画像のパス]({image_path})")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def make_graph(df):
    try:
        # 日数の設定値
        day_settings = {
            'low': 5,
            'min': 10,
            'max': 20
        }
        
        df.set_index('経過日数', inplace=True)
        # グラフを作成
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # データを期間で分類
        data_low = df[(df.index >= day_settings['low']) & 
                     (df.index < day_settings['min'])]  # 5日以上10日未満
        data_middle = df[(df.index >= day_settings['min']) & 
                        (df.index < day_settings['max'])]  # 10日以上20日未満
        data_pending_break = df[df.index >= day_settings['max']]  # 20日以上
        
        # ヒストグラム - 期間ごとに色分け
        max_days = df.index.max()
        bins = list(range(day_settings['low'], int(max_days) + 2, 1))  # 1日ごとのビン
        
        # 5-10日のデータ
        ax.hist(data_low.index, bins=bins, weights=data_low['数値'], 
               color='lightgreen', edgecolor='green', alpha=0.5,
               label=f'{day_settings["low"]}～{day_settings["min"]}日')
        # 10-20日のデータ
        ax.hist(data_middle.index, bins=bins, weights=data_middle['数値'], 
               color='lightblue', edgecolor='blue', alpha=0.5,
               label=f'{day_settings["min"]}～{day_settings["max"]}日')
        # 20日以上のデータ
        ax.hist(data_pending_break.index, bins=bins, weights=data_pending_break['数値'], 
               color='pink', edgecolor='red', alpha=0.5,
               label=f'{day_settings["max"]}日以上')
        
        ax.set_xlabel('経過日数')
        ax.set_ylabel('数値')
        ax.grid(True, alpha=0.3)

        plt.title('経過日数別の数値分布')
        plt.legend(loc='upper right')

        plt.tight_layout()
        plt.savefig('histogram.png')
        plt.close()
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        
def send_daily_data():
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
        make_graph(df)
        send_graph_message()
    except Exception as e:
        print(f"エラーが発生しました: {e}")

def update_excel(file_path):
    # Excelアプリケーションを起動し、ブックを開く
    app = xw.App(visible=True)  # Excelを非表示で実行
    wb = app.books.open(file_path)

    # 外部参照を更新
    wb.api.RefreshAll()
    app.api.CalculateUntilAsyncQueriesDone()

    # シートの内容を確認（例として最初のシートのA1:A11を表示）
    print(wb.sheets[0].range("A1:A11").value)

    # 保存して閉じる
    wb.save()
    wb.close()
    app.quit()

def process_and_send_data(file_path, min_time=8, max_time=12):
    # Excelファイルを更新
    update_excel(file_path)
    
    # pandasでデータを取得
    df = pd.read_excel(file_path)
    
    # 出勤時間で並べ替え
    df_sorted = df.sort_values(by='出勤時間')
    
    # 出勤時間を数値に変換
    df_sorted['出勤時間'] = pd.to_numeric(df_sorted['出勤時間'], errors='coerce')

    # 引数で指定された範囲でフィルタリング
    df_filtered = df_sorted[(df_sorted['出勤時間'] >= min_time) & (df_sorted['出勤時間'] <= max_time)]
    
    # 必要な列を選択
    df_filtered_columns = df_filtered[['出勤時間', 'あ']]
    markdown_content = df_filtered_columns.to_markdown(index=False)
    send_simple_message(markdown_content)
    print(markdown_content)

# 実行
process_and_send_data(excel_file_path3)

send_daily_data()
send_simple_message("今日は備品補充の日です。必要な備品を確認してください。")
#post_supply_reminder()
# スケジュール設定
#schedule.every().day.at("11:00").do(post_pending_data)
#schedule.every().tuesday.at("10:00").do(post_supply_reminder)
#schedule.every().friday.at("10:00").do(post_supply_reminder)

#while True:
#    try:
#        schedule.run_pending()
#        time.sleep(1)
#    except Exception as e:
#        print(f"エラーが発生しました: {e}")