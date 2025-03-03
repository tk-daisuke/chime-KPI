import requests
from config.settings import Settings

class ChimeService:
    def __init__(self):
        self.webhook_url = Settings.WEBHOOK_URL
        self.max_size = Settings.MAX_MESSAGE_SIZE

    def send_msg(self, content):
        """メッセージ送信"""
        content = self._fix_msg_size(content)
        message = {"Content": f"/md\n{content}"}
        
        try:
            response = requests.post(self.webhook_url, json=message)
            if response.status_code == 200:
                print("メッセージが正常に投稿されました")
            else:
                print("メッセージの投稿に失敗しました")
        except Exception as e:
            print(f"エラーが発生しました: {e}")

    def _fix_msg_size(self, content):
        """メッセージサイズ調整"""
        while len(content.encode('utf-8')) > self.max_size:
            lines = content.split('\n')
            content = '\n'.join(lines[:-1])
        return content
