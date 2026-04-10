from openpyxl.chart import Reference
from openpyxl.chart.series import XYSeries, SeriesLabel
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
from openpyxl.drawing.text import Font as DrawingFont
from openpyxl.chart.text import RichText
from config import Config


class ChartManager:
    @staticmethod
    def update_chart_range(worksheet, num_rows, chart_title=None, data_start_row=None,
                           row_offset=0, x_axis_max=None, shift_chart_col=0):
        """
        Updates the chart data range, title, position, and x-axis max.

        Parameters:
        -----------
        worksheet       : openpyxl worksheet object
        num_rows        : int - number of data rows
        chart_title     : str, optional - title to set for the chart
        data_start_row  : int, optional - dynamic starting row for data
        row_offset      : int - how many rows were inserted (shifts chart down)
        x_axis_max      : float, optional - max value for x-axis (max midpoint)
        shift_chart_col : int - how many columns to shift chart right (0 or 1)
        """
        if not worksheet._charts:
            return

        chart = worksheet._charts[0]

        # --- SHIFT CHART POSITION (down for inserted rows, right if pipe type col needs space) ---
        if hasattr(chart, 'anchor'):
            if hasattr(chart.anchor, '_from'):
                chart.anchor._from.row += row_offset
                chart.anchor._from.col += shift_chart_col
            if hasattr(chart.anchor, 'to'):
                chart.anchor.to.row += row_offset
                chart.anchor.to.col += shift_chart_col

        # --- SET CHART TITLE ---
        if chart_title:
            chart.title = chart_title

            if hasattr(chart.title, 'tx') and chart.title.tx is not None:
                rich = RichText()
                paragraph = Paragraph()

                char_props = CharacterProperties()
                char_props.latin = DrawingFont(typeface="Aptos Narrow")
                char_props.sz = 1400
                char_props.b = False

                para_props = ParagraphProperties()
                para_props.defRPr = char_props
                paragraph.pPr = para_props

                text_run = RegularTextRun()
                text_run.t = chart_title
                text_run.rPr = char_props
                paragraph.r = [text_run]

                rich.p = [paragraph]
                chart.title.tx.rich = rich

                if hasattr(chart.title, 'layout') and chart.title.layout is None:
                    from openpyxl.chart.layout import Layout, ManualLayout
                    chart.title.layout = Layout()
                    chart.title.layout.manualLayout = ManualLayout()
                    chart.title.layout.manualLayout.yMode = "edge"
                    chart.title.layout.manualLayout.xMode = "edge"
                    chart.title.layout.manualLayout.x = 0.1
                    chart.title.layout.manualLayout.y = 0.05

        if num_rows == 0:
            return

        # --- SET DATA RANGES ---
        min_row = data_start_row if data_start_row is not None else (Config.DATA_START_ROW + row_offset)
        max_row = min_row + num_rows - 1

        x_ref  = Reference(worksheet, min_col=4,                        min_row=min_row, max_row=max_row)
        y1_ref = Reference(worksheet, min_col=Config.COL_DRI_THICKNESS,  min_row=min_row, max_row=max_row)
        y2_ref = Reference(worksheet, min_col=Config.COL_NOM_THICKNESS,  min_row=min_row, max_row=max_row)

        x_data  = AxDataSource(NumRef(x_ref))
        y1_data = NumDataSource(NumRef(y1_ref))
        y2_data = NumDataSource(NumRef(y2_ref))

        series1 = XYSeries()
        series1.yVal  = y1_data
        series1.xVal  = x_data
        series1.title = SeriesLabel(v="DRI Thickness")

        series2 = XYSeries()
        series2.yVal  = y2_data
        series2.xVal  = x_data
        series2.title = SeriesLabel(v="Nominal Thickness")

        chart.series = [series1, series2]

        # --- SET X-AXIS MAX ---
        if x_axis_max and x_axis_max > 0:
            if hasattr(chart, 'x_axis'):
                chart.x_axis.scaling.max = float(x_axis_max)
                chart.x_axis.scaling.min = 0.0
