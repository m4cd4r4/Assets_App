import os
import sys
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.widgets import Slider
from datetime import datetime

class StockLevelVisualizer:
    def __init__(self, file_name):
        self.application_path = self._determine_application_path()
        self.file_path = os.path.join(self.application_path, file_name)
        self.df_items = self._load_data('4.2_Items')
        self.df_timestamps = self._load_data('4.2_Timestamps')
        if not self.df_timestamps.empty:
            self.initial_date = pd.to_datetime(self.df_timestamps['Timestamp']).min()
            self.final_date = pd.to_datetime(self.df_timestamps['Timestamp']).max()
            self.current_date = self.initial_date
        else:
            self.initial_date = self.final_date = self.current_date = pd.to_datetime('today')

    @staticmethod
    def _determine_application_path():
        if getattr(sys, 'frozen', False):
            return sys._MEIPASS
        else:
            return os.path.dirname(__file__)

    def _load_data(self, sheet_name):
        return pd.read_excel(self.file_path, sheet_name=sheet_name)

    def _calculate_stock_levels(self, date):
        df_filtered = self.df_timestamps[self.df_timestamps['Timestamp'] <= date].copy()
        df_filtered['ActionValue'] = df_filtered['Action'].str.extract(r'(\d+)', expand=False).astype(float)
        df_filtered['Operation'] = df_filtered['Action'].str.contains('Add').map({True: 1, False: -1})
        df_filtered['AdjustedValue'] = df_filtered['ActionValue'] * df_filtered['Operation']
        stock_levels = df_filtered.groupby('Item').agg({'AdjustedValue': 'sum'}).reset_index()
        stock_levels = pd.merge(self.df_items, stock_levels, on='Item', how='left').fillna(0)
        return stock_levels

    def plot(self):
        fig, ax = plt.subplots()
        plt.subplots_adjust(bottom=0.25)
        
        stock_levels = self._calculate_stock_levels(self.current_date)
        plt.barh(stock_levels['Item'], stock_levels['AdjustedValue'], color='skyblue')
        plt.xlim(0, 120)  # Set the maximum range as specified

        axdate = plt.axes([0.25, 0.1, 0.65, 0.03])
        date_slider = Slider(axdate, 'Date', valmin=mdates.date2num(self.initial_date), valmax=mdates.date2num(self.final_date), valinit=mdates.date2num(self.current_date), valfmt='%0.0f')

        def update(val):
            date = mdates.num2date(date_slider.val).date()
            stock_levels = self._calculate_stock_levels(date)
            ax.clear()
            plt.barh(stock_levels['Item'], stock_levels['AdjustedValue'], color='skyblue')
            plt.xlim(0, 120)  # Ensure the x-axis is consistent
            fig.canvas.draw_idle()

        date_slider.on_changed(update)
        plt.show()

if __name__ == "__main__":
    visualizer = StockLevelVisualizer('EUC_Perth_Assets.xlsx')
    visualizer.plot()
