# 仮想環境の構築手順
# 仮想環境を作成: python -m venv .venv
# 仮想環境をアクティベート: .venv\Scripts\activate

# 仮想環境のアクティベート確認
# ターミナルに (.venv) と表示されれば成功
#pip freeze > requirements.txt
#pip install -r requirements.txt
#python -m pip install setuptools
import os
import time
import pandas as pd
import xlwings as xw
import requests
import matplotlib.pyplot as plt
import japanize_matplotlib
import pyfiglet
import schedule
from datetime import datetime

# 設定値
class Settings:
    # Excelファイル設定
    excel_files = {
        'book3': {
            'path': 'data.xlsx',
            'sheet': 'Sheet1'
        }
    }
    
    # グラフ設定
    graph_settings = {
        'output_path': 'output_graph.png',
        'day_limits': {
            'low': 0,
            'min': 30,
            'max': 60
        }
    }
    
    # Chime設定
    hook_url = "https://your-webhook-url"
    max_msg_size = 4096

# ローディングアニメーションを表示する関数
def showLoad():
    """ローディング表示"""
    loading_text = "Loading"
    for i in range(4):
        print(loading_text + '.' * i, end='\r')
        time.sleep(0.5)
    print(loading_text + '....')

# Excelデータ取得クラス
class ExcelSvc:
    def __init__(self):
        self.settings = Settings.excel_files

    def getData(self, book):
        """Excelデータ取得"""
        file_info = self.settings.get(book)
        if not file_info:
            raise ValueError(f"Unknown book key: {book}")
            
        return pd.read_excel(
            file_info['path'],
            sheet_name=file_info['sheet']
        )

    def updateBook(self, filePath):
        """Excelブック更新"""
        try:
            app = xw.App(visible=False)
            wb = app.books.open(filePath)
            wb.api.RefreshAll()
            app.api.CalculateUntilAsyncQueriesDone()
            wb.save()
            wb.close()
            app.quit()
        except Exception as e:
            print(f"エラーが発生しました: {e}")
            os.system("taskkill /f /im EXCEL.EXE")

# Chimeメッセージ送信クラス
class ChimeSvc:
    def __init__(self):
        self.hook_url = Settings.hook_url
        self.max_size = Settings.max_msg_size

    def sendMsg(self, content):
        """メッセージ送信"""
        content = self.fixMsgSize(content)
        message = {"Content": f"/md\n{content}"}
        
        try:
            response = requests.post(self.hook_url, json=message)
            if response.status_code == 200:
                print("メッセージが正常に投稿されました")
            else:
                print("メッセージの投稿に失敗しました")
        except Exception as e:
            print(f"エラーが発生しました: {e}")

    def fixMsgSize(self, content):
        """メッセージサイズ調整"""
        while len(content.encode('utf-8')) > self.max_size:
            lines = content.split('\n')
            content = '\n'.join(lines[:-1])
        return content

# グラフ作成クラス
class GraphSvc:
    def __init__(self):
        self.settings = Settings.graph_settings

    def plotGraph(self, data):
        """グラフ作成のメイン処理"""
        try:
            data.set_index('経過日数', inplace=True)
            fig, ax = plt.subplots(figsize=(12, 8))
            
            self.plotData(ax, data)
            self.setGraphStyle(ax)
            
            plt.savefig(self.settings['output_path'])
            plt.close()
        except Exception as e:
            print(f"グラフ作成でエラー: {e}")

    def plotData(self, ax, data):
        """データごとのグラフ作成"""
        day_limits = self.settings['day_limits']
        bins = self.makeBins(data)
        
        ranges = [
            {'data': data[(data.index >= day_limits['low']) & (data.index < day_limits['min'])],
             'color': 'lightgreen', 'edge': 'green', 'label': f"{day_limits['low']}～{day_limits['min']}日"},
            {'data': data[(data.index >= day_limits['min']) & (data.index < day_limits['max'])],
             'color': 'lightblue', 'edge': 'blue', 'label': f"{day_limits['min']}～{day_limits['max']}日"},
            {'data': data[data.index >= day_limits['max']],
             'color': 'pink', 'edge': 'red', 'label': f"{day_limits['max']}日以上"}
        ]
        
        for irange in ranges:
            ax.hist(irange['data'].index, bins=bins, weights=irange['data']['数値'],
                   color=irange['color'], edgecolor=irange['edge'],
                   alpha=0.5, label=irange['label'])

    def makeBins(self, data):
        """グラフの区切り作成"""
        day_limits = self.settings['day_limits']
        max_days = data.index.max()
        return list(range(day_limits['low'], int(max_days) + 2, 1))

    def setGraphStyle(self, ax):
        """グラフの見た目設定"""
        ax.set_xlabel('経過日数')
        ax.set_ylabel('数値')
        ax.grid(True, alpha=0.3)
        plt.title('経過日数別の数値分布')
        plt.legend(loc='upper right')
        plt.tight_layout()

# サービスのインスタンス化
excelSvc = ExcelSvc()
chimeSvc = ChimeSvc()
graphSvc = GraphSvc()

def runData(book):
    """データ作成と送信"""
    try:
        print(f"データ処理を開始します: {book}")
        excelSvc.updateBook(Settings.excel_files[book]['path'])
        
        data = excelSvc.getData(book)
        print("データを取得しました")
        
        graphSvc.plotGraph(data)
        print("グラフを作成しました")
        
        # データの一部を送信
        msg_content = data[['出勤時間', 'あ']].to_markdown(index=False)
        chimeSvc.sendMsg(msg_content)
        
        # グラフのパスを送信
        img_path = os.path.abspath(Settings.graph_settings['output_path'])
        chimeSvc.sendMsg(f"今後の予想のグラフです。[画像のパス]({img_path})")
        
        print("処理が完了しました")
    except Exception as e:
        print(f"エラー発生: {e}")

def remindSupply():
    """備品補充のリマインダー送信"""
    chimeSvc.sendMsg("今日は備品補充の日です。必要な備品を確認してください。")

def setSched():
    """スケジュール設定"""
    # 平日の11時にデータを送信
    schedule.every().monday.at("11:00").do(lambda: runData('book3'))
    schedule.every().tuesday.at("11:00").do(lambda: runData('book3'))
    schedule.every().wednesday.at("11:00").do(lambda: runData('book3'))
    schedule.every().thursday.at("11:00").do(lambda: runData('book3'))
    schedule.every().friday.at("11:00").do(lambda: runData('book3'))
    
    # 火曜と金曜に備品リマインダー
    schedule.every().tuesday.at("10:00").do(remindSupply)
    schedule.every().friday.at("10:00").do(remindSupply)
    
    print("スケジュールを設定しました")

def main():
    """メイン関数"""
    # タイトル表示
    titleStr = pyfiglet.figlet_format("KPI 自動通知システム", font="slant")
    print(titleStr)
    
    # ローディング表示
    showLoad()
    
    # 起動時に一度実行
    runData('book3')
    
    # スケジュール設定
    setSched()
    
    print(f"システムが起動しました - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("このウィンドウを閉じないでください。閉じると通知が停止します。")
    print("Ctrl+Cで終了できます。")
    
    # スケジュールを継続的に実行
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nシステムを終了します...")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()