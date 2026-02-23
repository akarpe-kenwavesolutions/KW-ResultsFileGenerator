from openpyxl.chart import Reference
from openpyxl.chart.series import XYSeries, SeriesLabel
from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
from openpyxl.drawing.text import Font as DrawingFont
from openpyxl.chart.text import RichText
from config import Config


class ChartManager:
    @staticmethod
    def update_chart_range(worksheet, num_rows, chart_title=None, data_start_row=None, row_offset=0, x_axis_max=None):
        if not worksheet._charts:
            return

        chart = worksheet._charts[0]

        if row_offset > 0 and hasattr(chart, 'anchor'):
            if hasattr(chart.anchor, '_from'):
                chart.anchor._from.row += row_offset
            if hasattr(chart.anchor, 'to'):
                chart.anchor.to.row += row_offset

        if chart_title:
            chart.title = chart_title
            if hasattr(chart.title, 'tx') and chart.title.tx is not None:
                from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RegularTextRun
                from openpyxl.drawing.text import Font as DrawingFont
                from openpyxl.chart.text import RichText
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

        min_row = data_start_row if data_start_row is not None else (Config.DATA_START_ROW + row_offset)
        max_row = min_row + num_rows - 1

        if num_rows == 0:
            return

        from openpyxl.chart import Reference
        from openpyxl.chart.series import XYSeries, SeriesLabel
        from openpyxl.chart.data_source import NumDataSource, NumRef, AxDataSource

        x_ref = Reference(worksheet, min_col=4, min_row=min_row, max_row=max_row)
        y1_ref = Reference(worksheet, min_col=Config.COL_DRI_THICKNESS, min_row=min_row, max_row=max_row)
        y2_ref = Reference(worksheet, min_col=Config.COL_NOM_THICKNESS, min_row=min_row, max_row=max_row)

        x_data = AxDataSource(NumRef(x_ref))
        y1_data = NumDataSource(NumRef(y1_ref))
        y2_data = NumDataSource(NumRef(y2_ref))

        series1 = XYSeries()
        series1.yVal = y1_data
        series1.xVal = x_data
        series1.title = SeriesLabel(v="DRI Thickness")

        series2 = XYSeries()
        series2.yVal = y2_data
        series2.xVal = x_data
        series2.title = SeriesLabel(v="Nominal Thickness")

        chart.series = [series1, series2]

        # --- Set x-axis max to max midpoint value ---
        if x_axis_max and x_axis_max > 0:
            if hasattr(chart, 'x_axis'):
                chart.x_axis.scaling.max = float(x_axis_max)
                chart.x_axis.scaling.min = 0.0

