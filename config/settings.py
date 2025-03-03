import os

class Settings:
    # Excel設定
    EXCEL_FILES = {
        'book1': {'path': 'book1.xlsx', 'sheet': 'Sheet1'},
        'book2': {'path': 'book2.xlsx', 'sheet': 'Sheet1'},
        'book3': {'path': 'Book3.xlsx', 'sheet': 'Sheet1'},
    }
    
    # Webhook設定
    WEBHOOK_URL = 'https://webhook-test.com/98b2ff73ac2524023b6f209fc5cb7c7e'
    
    # グラフ設定
    GRAPH_SETTINGS = {
        'day_limits': {
            'low': 5,
            'min': 10,
            'max': 20
        },
        'output_path': 'histogram.png'
    }
    
    # メッセージ設定
    MAX_MESSAGE_SIZE = 3900  # 3.9KB
