import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.widgets import Slider
from datetime import datetime

class StockVisualizer:
    def __init__(self, file_name):
        self.file_name = file_name
        self.application_path = self._determine_application_path()
        self.file_path = os.path.join(self.application_path, file_name)
        self.df_stocks = self._load_data()

    @staticmethod
    def _determine_application_path():
        if getattr(sys, 'frozen', False):
            return sys._MEIPASS
        else:
            return os.path.dirname(__file__)

    def _load_data(self):
        df = pd.read_excel(self.file_path, sheet_name='4.2_Timestamps', usecols=[0, 1, 2])
        df['ActionValue'] = df['Action'].str.extract('(\d+)', expand=False).astype(float)
        df['Operation'] = df['Action'].str.extract('(Add|Subtract)', expand=False)
        df['Operation'] = df['Operation'].map({'Add': 1, 'Subtract': -1})
        df['AdjustedValue'] = df['Operation'] * df['ActionValue']
        return df

    def _calculate_stocks(self, date):
        filtered_df = self.df_stocks[self.df_stocks['Timestamp'] <= date]
        stock_levels = filtered_df.groupby('Item').agg({'AdjustedValue': 'sum'}).reset_index()
        return stock_levels

    def plot_stocks(self, initial_date_index=0):
        fig, ax = plt.subplots()
        plt.subplots_adjust(left=0.1, bottom=0.25)
        dates = pd.to_datetime(self.df_stocks['Timestamp'].unique()).sort_values()
        stock_levels = self._calculate_stocks(dates[initial_date_index])

        bars = plt.barh(stock_levels['Item'], stock_levels['AdjustedValue'], color='skyblue')
        axcolor = 'lightgoldenrodyellow'
        ax_date = plt.axes([0.1, 0.1, 0.65, 0.03], facecolor=axcolor)
        date_slider = Slider(ax=ax_date, label='Date', valmin=0, valmax=len(dates)-1, valinit=initial_date_index, valfmt='%0.0f')

        def update(val):
            ax.clear()
            current_date_index = int(date_slider.val)
            stock_levels = self._calculate_stocks(dates[current_date_index])
            plt.barh(stock_levels['Item'], stock_levels['AdjustedValue'], color='skyblue')
            ax.figure.canvas.draw_idle()

        date_slider.on_changed(update)

        plt.ylabel('Item')
        plt.xlabel('Stock Level')
        plt.title('Stock Levels Over Time')
        plt.show()

if __name__ == "__main__":
    visualizer = StockVisualizer('EUC_Perth_Assets.xlsx')
    visualizer.plot_stocks()
