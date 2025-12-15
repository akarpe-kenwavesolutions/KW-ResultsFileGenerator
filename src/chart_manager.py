from openpyxl.chart import Reference
# CHANGE: Import XYSeries specifically for Scatter Charts
from openpyxl.chart.series import XYSeries, SeriesLabel
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource
from config import Config


class ChartManager:
    @staticmethod
    def update_chart_range(worksheet, num_rows):
        if not worksheet._charts:
            return

        chart = worksheet._charts[0]
        min_row = Config.DATA_START_ROW
        max_row = min_row + num_rows - 1

        if num_rows == 0:
            return

        # 1. Define References
        x_ref = Reference(worksheet, min_col=4, min_row=min_row, max_row=max_row)
        y1_ref = Reference(worksheet, min_col=Config.COL_DRI_THICKNESS, min_row=min_row, max_row=max_row)
        y2_ref = Reference(worksheet, min_col=Config.COL_NOM_THICKNESS, min_row=min_row, max_row=max_row)

        # 2. Wrap Data
        x_data = AxDataSource(NumRef(x_ref))
        y1_data = NumDataSource(NumRef(y1_ref))
        y2_data = NumDataSource(NumRef(y2_ref))

        # 3. Create New Series (Using XYSeries)
        # Note: XYSeries usually does not take 'cat' in __init__, it takes xVal and yVal.
        # But looking at source code, generic Series works if cast? No, the error was explicit.
        # Let's try standard initialization.

        series1 = XYSeries()
        series1.yVal = y1_data  # Y-Values (NumDataSource)
        series1.xVal = x_data  # X-Values (AxDataSource) - Scatter calls it xVal, not cat
        series1.title = SeriesLabel(v="DRI Thickness")

        series2 = XYSeries()
        series2.yVal = y2_data
        series2.xVal = x_data
        series2.title = SeriesLabel(v="Nominal Thickness")

        # 4. Assign to Chart
        chart.series = [series1, series2]
