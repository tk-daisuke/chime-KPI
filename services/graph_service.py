import matplotlib.pyplot as plt
import japanize_matplotlib
from config.settings import Settings

class GraphService:
    def __init__(self):
        self.settings = Settings.GRAPH_SETTINGS

    def make_graph(self, df):
        """グラフ作成のメイン処理"""
        try:
            df.set_index('経過日数', inplace=True)
            fig, ax = plt.subplots(figsize=(12, 8))
            
            self._make_data_graph(ax, df)
            self._set_graph_look(ax)
            
            plt.savefig(self.settings['output_path'])
            plt.close()
        except Exception as e:
            print(f"グラフ作成でエラー: {e}")

    def _make_data_graph(self, ax, df):
        """データごとのグラフ作成"""
        day_limits = self.settings['day_limits']
        bins = self._make_bins(df)
        
        ranges = [
            {'data': df[(df.index >= day_limits['low']) & (df.index < day_limits['min'])],
             'color': 'lightgreen', 'edge': 'green', 'label': f"{day_limits['low']}～{day_limits['min']}日"},
            {'data': df[(df.index >= day_limits['min']) & (df.index < day_limits['max'])],
             'color': 'lightblue', 'edge': 'blue', 'label': f"{day_limits['min']}～{day_limits['max']}日"},
            {'data': df[df.index >= day_limits['max']],
             'color': 'pink', 'edge': 'red', 'label': f"{day_limits['max']}日以上"}
        ]
        
        for range_info in ranges:
            ax.hist(range_info['data'].index, bins=bins, weights=range_info['data']['数値'],
                   color=range_info['color'], edgecolor=range_info['edge'],
                   alpha=0.5, label=range_info['label'])

    def _make_bins(self, df):
        """グラフの区切り作成"""
        day_limits = self.settings['day_limits']
        max_days = df.index.max()
        return list(range(day_limits['low'], int(max_days) + 2, 1))

    def _set_graph_look(self, ax):
        """グラフの見た目設定"""
        ax.set_xlabel('経過日数')
        ax.set_ylabel('数値')
        ax.grid(True, alpha=0.3)
        plt.title('経過日数別の数値分布')
        plt.legend(loc='upper right')
        plt.tight_layout()
