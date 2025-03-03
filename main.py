import os
import time
import pyfiglet
from services.excel_service import ExcelService
from services.chime_service import ChimeService
from services.graph_service import GraphService
from config.settings import Settings

# 仮想環境の構築手順
# 仮想環境を作成: python -m venv .venv
# 仮想環境をアクティベート: .venv\Scripts\activate

# 仮想環境のアクティベート確認
# ターミナルに (.venv) と表示されれば成功
#pip freeze > requirements.txt
#pip install -r requirements.txt
#python -m pip install setuptools

# ローディングアニメーションを表示する関数
def show_loading():
    """ローディング表示"""
    loading_text = "Loading"
    for i in range(4):  # 4回繰り返す
        print(loading_text + '.' * i, end='\r')  # カーソルを行の先頭に戻す
        time.sleep(0.5)  # 0.5秒待つ
    print(loading_text + '....')  # 最後に4つのドットを表示

# サービスのインスタンス化
excel_service = ExcelService()
chime_service = ChimeService()
graph_service = GraphService()

def make_and_send_data(book_key):
    """データ作成と送信"""
    try:
        excel_service.update_book(Settings.EXCEL_FILES[book_key]['path'])
        
        df = excel_service.get_data(book_key)
        
        graph_service.make_graph(df)
        
        msg_content = df[['出勤時間', 'あ']].to_markdown(index=False)
        chime_service.send_msg(msg_content)
        
        img_path = os.path.abspath(Settings.GRAPH_SETTINGS['output_path'])
        chime_service.send_msg(f"今後の予想のグラフです。[画像のパス]({img_path})")
        
    except Exception as e:
        print(f"エラー発生: {e}")

def main():
    show_title = pyfiglet.figlet_format("SYSTEM START", font="slant")
    print(show_title)
    show_loading()
    make_and_send_data('book3')
    chime_service.send_msg("今日は備品補充の日です。必要な備品を確認してください。")

if __name__ == "__main__":
    main()



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