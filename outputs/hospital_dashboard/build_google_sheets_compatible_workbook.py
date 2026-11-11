from __future__ import annotations
import csv
import math
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parents[2]
SOURCE_XLSX = ROOT / 'hospital_patient_analytics_dash_1.xlsx'
SOURCE_CSV = ROOT / 'clean_patients_cleaned.csv'
OUTPUT_XLSX = ROOT / 'outputs' / 'hospital_dashboard' / 'Hospital_Patient_Analytics_Dashboard_GoogleSheets_Verified.xlsx'
NAVY = '12324D'
BLUE = '2F78B7'
TEAL = '23988E'
AMBER = 'D99A31'
PURPLE = '7A5AA6'
GREEN = '4E8A68'
SLATE = '687A89'
PALE = 'EAF2F8'
WHITE = 'FFFFFF'


def normalize_csv_value(header, value):
    if value == '':
        return None
    if header in {'admission_date', 'discharge_date'}:
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y'):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass
        return value.strip()
    if header in {'age', 'treatment_cost', 'opd_waiting_time_min', 'room_allocation_time_min', 'patient_rating'}:
        try:
            return float(value)
        except ValueError:
            return value.strip()
    return value.strip()


def values_equal(header, expected, actual):
    expected = normalize_csv_value(header, expected)
    if isinstance(actual, datetime):
        actual = actual.date()
    if isinstance(actual, date) and not isinstance(actual, datetime):
        pass
    if header in {'age', 'treatment_cost', 'opd_waiting_time_min', 'room_allocation_time_min', 'patient_rating'} and actual is not None:
        try:
            actual = float(actual)
        except (TypeError, ValueError):
            pass
    if isinstance(expected, float) and isinstance(actual, float):
        return math.isclose(expected, actual, rel_tol=1e-12, abs_tol=1e-9)
    return expected == actual


def validate_patient_sheet(ws, csv_path):
    with csv_path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        headers = next(reader)
        source_rows = list(reader)
    if ws.max_row != len(source_rows) + 1 or ws.max_column != len(headers):
        raise ValueError(f'Patient Data dimensions differ from CSV: sheet={ws.max_row-1}x{ws.max_column}, CSV={len(source_rows)}x{len(headers)}')
    sheet_headers = [c.value for c in ws[1]]
    if sheet_headers != headers:
        raise ValueError('Patient Data column headers do not exactly match clean_patients_cleaned.csv')
    for row_num, (source_row, sheet_row) in enumerate(zip(source_rows, ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True)), start=2):
        for header, expected, actual in zip(headers, source_row, sheet_row):
            if not values_equal(header, expected, actual):
                raise ValueError(f'Data mismatch at Patient Data row {row_num}, column {header}: CSV={expected!r}, XLSX={actual!r}')
    return headers, source_rows


def read_chart_table(chart_ws, col_start, col_end, row_start=2):
    headers = [chart_ws.cell(1, c).value for c in range(col_start, col_end + 1)]
    rows = []
    for row in range(row_start, chart_ws.max_row + 1):
        vals = [chart_ws.cell(row, c).value for c in range(col_start, col_end + 1)]
        if vals[0] is None:
            break
        rows.append(vals)
    return headers, rows


def style_table(ws, ref):
    table = Table(displayName=ref[0], ref=ref[1])
    table.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)


def add_data_block(ws, col, title, headers, rows, start_row=1):
    ws.cell(start_row, col, title).font = Font(bold=True, color=WHITE, size=11)
    ws.cell(start_row, col).fill = PatternFill('solid', fgColor=NAVY)
    ws.merge_cells(start_row=start_row, start_column=col, end_row=start_row, end_column=col + len(headers) - 1)
    for j, header in enumerate(headers, start=col):
        cell = ws.cell(start_row + 1, j, header)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill('solid', fgColor=BLUE)
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
    for i, row in enumerate(rows, start=start_row + 2):
        for j, val in enumerate(row, start=col):
            ws.cell(i, j, val)
    last_row = max(start_row + 2, start_row + 1 + len(rows))
    for j, header in enumerate(headers, start=col):
        ws.column_dimensions[ws.cell(start_row + 1, j).column_letter].width = max(15, min(30, len(str(header)) + 4))
    return (col, start_row + 1, col + len(headers) - 1, last_row)


def add_bar_chart(ws, source_ws, block, title, anchor, *, horizontal=False, stacked=None, width=14.0, height=8.2, colors=None, x_title=None, y_title=None):
    min_col, min_row, max_col, max_row = block
    chart = BarChart()
    chart.type = 'bar' if horizontal else 'col'
    chart.style = 10
    chart.title = title
    chart.y_axis.title = x_title if horizontal else y_title
    chart.x_axis.title = y_title if horizontal else x_title
    chart.height = height
    chart.width = width
    chart.grouping = stacked or 'clustered'
    if stacked:
        chart.overlap = 100
    data = Reference(source_ws, min_col=min_col + 1, max_col=max_col, min_row=min_row, max_row=max_row)
    cats = Reference(source_ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend.position = 'b' if max_col - min_col > 1 else 'r'
    chart.varyColors = False
    palette = colors or [BLUE, TEAL, AMBER, PURPLE, GREEN, SLATE]
    for i, series in enumerate(chart.series):
        series.graphicalProperties.solidFill = palette[i % len(palette)]
        series.graphicalProperties.line.solidFill = palette[i % len(palette)]
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = False
    ws.add_chart(chart, anchor)
    return chart


def add_line_chart(ws, source_ws, block, title, anchor, *, width=14.0, height=8.2, y_title=None, number_format=None):
    min_col, min_row, max_col, max_row = block
    chart = LineChart()
    chart.style = 13
    chart.title = title
    chart.y_axis.title = y_title
    chart.height = height
    chart.width = width
    data = Reference(source_ws, min_col=min_col + 1, max_col=max_col, min_row=min_row, max_row=max_row)
    cats = Reference(source_ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None if max_col == min_col + 1 else chart.legend
    for i, series in enumerate(chart.series):
        series.graphicalProperties.line.solidFill = [BLUE, TEAL, AMBER, PURPLE][i % 4]
        series.graphicalProperties.line.width = 26000
        series.marker.symbol = 'circle'
        series.marker.size = 6
    if number_format:
        chart.y_axis.numFmt = number_format
    ws.add_chart(chart, anchor)
    return chart


def add_pie_chart(ws, source_ws, block, title, anchor, *, doughnut=True, width=13.2, height=8.2):
    min_col, min_row, max_col, max_row = block
    chart = PieChart()
    chart.style = 10
    chart.title = title
    chart.height = height
    chart.width = width
    labels = Reference(source_ws, min_col=min_col, min_row=min_row + 1, max_row=max_row)
    data = Reference(source_ws, min_col=min_col + 1, min_row=min_row, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.varyColors = True
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showLeaderLines = True
    if doughnut:
        chart.holeSize = 58
    colors = [BLUE, TEAL, AMBER, PURPLE, GREEN, SLATE, '5BA7C8', 'B95C56']
    chart.series[0].data_points = [DataPoint(idx=i, spPr=GraphicalProperties(solidFill=colors[i % len(colors)])) for i in range(max_row - min_row)]
    ws.add_chart(chart, anchor)
    return chart


def set_caption(ws, cell, text):
    c = ws[cell]
    c.value = text
    c.font = Font(color=SLATE, italic=True, size=10)
    c.alignment = Alignment(vertical='center', wrap_text=True)


def cramers_v(records, a, b):
    cross = defaultdict(Counter)
    row_counts, col_counts = Counter(), Counter()
    n = 0
    for r in records:
        x, y = str(r.get(a, '')), str(r.get(b, ''))
        if not x or not y:
            continue
        cross[x][y] += 1
        row_counts[x] += 1
        col_counts[y] += 1
        n += 1
    if n == 0 or min(len(row_counts) - 1, len(col_counts) - 1) <= 0:
        return 0.0
    chi2 = 0.0
    for x, nx in row_counts.items():
        for y, ny in col_counts.items():
            expected = nx * ny / n
            observed = cross[x][y]
            if expected:
                chi2 += (observed - expected) ** 2 / expected
    return math.sqrt(chi2 / (n * min(len(row_counts) - 1, len(col_counts) - 1)))


def eta(records, group_field):
    grouped = defaultdict(list)
    ages = []
    for r in records:
        age = r.get('age')
        group = str(r.get(group_field, '')).strip()
        if age is None or not group:
            continue
        ages.append(float(age))
        grouped[group].append(float(age))
    if not ages:
        return 0.0
    mean = sum(ages) / len(ages)
    total = sum((x - mean) ** 2 for x in ages)
    if total == 0:
        return 0.0
    between = 0.0
    for group, xs in grouped.items():
        avg = sum(xs) / len(xs)
        between += len(xs) * (avg - mean) ** 2
    return math.sqrt(max(0.0, between / total))


def main():
    wb = load_workbook(SOURCE_XLSX)
    patient_ws = wb['Patient Data']
    headers, source_rows = validate_patient_sheet(patient_ws, SOURCE_CSV)
    chart_ws = wb['Chart Data']
    records = []
    for row in patient_ws.iter_rows(min_row=2, values_only=True):
        rec = dict(zip(headers, row))
        dt = rec.get('admission_date')
        if isinstance(dt, datetime):
            dt = dt.date()
        rec['admission_date'] = dt
        records.append(rec)
    june_rows = [r for r in records if r['admission_date'] and date(2026, 6, 1) <= r['admission_date'] <= date(2026, 6, 30)]
    if len(june_rows) != len(records):
        raise ValueError(f'Expected all dataset rows to be within June 2026, found {len(june_rows)} of {len(records)}')
    required_for_new_charts = ['age', 'gender', 'patient_type', 'severity_level', 'blood_group', 'disease']
    missing_by_field = {field: sum(r.get(field) in (None, '') for r in june_rows) for field in required_for_new_charts}
    if any(missing_by_field.values()):
        raise ValueError(f'Missing values would omit records from new charts: {missing_by_field}')

    # Rebuild supplemental chart source data from all 40,596 patient-level rows.
    if 'Plot Data' in wb.sheetnames:
        del wb['Plot Data']
    plot_data = wb.create_sheet('Plot Data')
    plot_data.sheet_view.showGridLines = False
    plot_data.freeze_panes = 'A3'

    # A helper block starts at column A; datasets for all generated charts occupy stable columns.
    blocks = {}
    blocks['daily'] = (1, 1, 2, chart_ws.max_row)
    # Existing Chart Data A:B; table header row is row 1 and values through row 31.
    # Existing summary tables are copied as values to the analysis source tab to simplify chart imports.
    def copy_chart_data(name, start_col, end_col, plot_col, title):
        hdr, rows = read_chart_table(chart_ws, start_col, end_col)
        return add_data_block(plot_data, plot_col, title, hdr, rows)

    blocks['daily'] = copy_chart_data('daily', 1, 2, 1, 'Daily patient records')
    blocks['patient_type'] = copy_chart_data('patient_type', 3, 4, 4, 'Patient type')
    blocks['diseases'] = copy_chart_data('diseases', 5, 6, 7, 'Top 10 diseases')
    blocks['cost_day'] = copy_chart_data('cost_day', 10, 11, 10, 'Average treatment cost per patient-day')
    blocks['service'] = copy_chart_data('service', 12, 13, 13, 'Average service time')
    blocks['outcomes'] = copy_chart_data('outcomes', 14, 15, 16, 'Patient outcomes')
    blocks['severity'] = copy_chart_data('severity', 16, 17, 19, 'Severity')
    blocks['ratings'] = copy_chart_data('ratings', 19, 20, 22, 'Rating distribution')
    blocks['payments'] = copy_chart_data('payments', 22, 23, 25, 'Payment method records')
    blocks['insurance'] = copy_chart_data('insurance', 25, 26, 28, 'Insurance records')
    blocks['cost_rating'] = copy_chart_data('cost_rating', 28, 29, 31, 'Average cost by rating')
    blocks['emergency'] = copy_chart_data('emergency', 31, 32, 34, 'Emergency patients by insurance')
    blocks['payment_cost'] = copy_chart_data('payment_cost', 34, 35, 37, 'Average treatment cost by payment method')
    blocks['ratings_by_type'] = copy_chart_data('ratings_by_type', 36, 38, 40, 'Rating counts by patient type')
    blocks['rating_summary'] = copy_chart_data('rating_summary', 40, 46, 44, 'Rating five-number summary by patient type')

    # Overlapping age histogram by gender; ten-year bins, all eligible records retained.
    ages = [int(float(r['age'])) for r in june_rows]
    lo = (min(ages) // 10) * 10
    hi = (max(ages) // 10) * 10
    age_labels = [f'{lower}–{lower+9}' for lower in range(lo, hi + 1, 10)]
    genders = sorted({str(r['gender']).strip() for r in june_rows})
    age_counts = defaultdict(Counter)
    for r in june_rows:
        age_counts[f'{(int(float(r["age"])) // 10) * 10}–{(int(float(r["age"])) // 10) * 10 + 9}'][str(r['gender']).strip()] += 1
    age_rows = [[band] + [age_counts[band][g] for g in genders] for band in age_labels]
    blocks['age'] = add_data_block(plot_data, 1, 'Age histogram by gender', ['Age band'] + genders, age_rows, start_row=40)

    # June total snapshot mirrors the daily animation's categories; animation itself is in the Plotly app.
    patient_types = sorted({str(r['patient_type']).strip() for r in june_rows})
    severity_order = ['Mild', 'Moderate', 'Severe', 'Critical']
    severity_order = [s for s in severity_order if any(str(r['severity_level']).strip() == s for r in june_rows)]
    type_severity = defaultdict(Counter)
    for r in june_rows:
        type_severity[str(r['patient_type']).strip()][str(r['severity_level']).strip()] += 1
    type_rows = [[patient_type] + [type_severity[patient_type][s] for s in severity_order] for patient_type in patient_types]
    blocks['type_severity'] = add_data_block(plot_data, 6, 'Patient type by severity · June total snapshot', ['Patient type'] + severity_order, type_rows, start_row=40)

    # Blood group and severity record counts as a Sheets-compatible stacked column chart.
    blood_groups = sorted({str(r['blood_group']).strip() for r in june_rows})
    blood_severity = defaultdict(Counter)
    for r in june_rows:
        blood_severity[str(r['blood_group']).strip()][str(r['severity_level']).strip()] += 1
    blood_rows = [[bg] + [blood_severity[bg][s] for s in severity_order] for bg in blood_groups]
    blocks['blood_severity'] = add_data_block(plot_data, 11, 'Blood group by severity', ['Blood group'] + severity_order, blood_rows, start_row=40)

    # Associations used by the Plotly heatmap, repeated here as a readable color-scaled matrix.
    vars_ = ['patient_type', 'gender', 'age', 'blood_group', 'disease']
    labels = ['Patient type', 'Gender', 'Age', 'Blood group', 'Disease']
    matrix = []
    for i, a in enumerate(vars_):
        row = [labels[i]]
        for j, b in enumerate(vars_):
            if i == j:
                value = 1.0
            elif a == 'age' or b == 'age':
                field = b if a == 'age' else a
                value = eta(june_rows, field)
            else:
                value = cramers_v(june_rows, a, b)
            row.append(value)
        matrix.append(row)
    if 'Association Map' in wb.sheetnames:
        del wb['Association Map']
    assoc = wb.create_sheet('Association Map')
    assoc.sheet_view.showGridLines = False
    assoc.merge_cells('A1:F1')
    assoc['A1'] = 'Association map · June 2026'
    assoc['A1'].font = Font(bold=True, color=WHITE, size=16)
    assoc['A1'].fill = PatternFill('solid', fgColor=NAVY)
    assoc['A2'] = 'Cramér’s V for categorical pairs; correlation ratio (η) for age-to-category pairs. Scores describe association, not causation.'
    assoc.merge_cells('A2:F2')
    assoc['A2'].font = Font(italic=True, color=SLATE, size=10)
    for j, label in enumerate(['Variable'] + labels, 1):
        c = assoc.cell(4, j, label)
        c.font = Font(bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=BLUE)
        c.alignment = Alignment(horizontal='center', wrap_text=True)
    for i, row in enumerate(matrix, 5):
        for j, v in enumerate(row, 1):
            c = assoc.cell(i, j, v)
            c.number_format = '0.000'
            c.alignment = Alignment(horizontal='center')
            if j == 1:
                c.font = Font(bold=True, color=NAVY)
    assoc.conditional_formatting.add('B5:F9', ColorScaleRule(start_type='num', start_value=0, start_color='F3F7FA', mid_type='num', mid_value=0.5, mid_color='77B7C5', end_type='num', end_value=1, end_color=NAVY))
    assoc.column_dimensions['A'].width = 22
    for col in 'BCDEF': assoc.column_dimensions[col].width = 19
    assoc.freeze_panes = 'B5'

    for ws in (plot_data,):
        ws.freeze_panes = 'A3'
        for row in ws.iter_rows(min_row=3):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.0' if isinstance(cell.value, float) and not cell.value.is_integer() else '#,##0'

    # Fill the prepared blank plot zones on the summary Dashboard with real Excel chart objects.
    dashboard = wb['Dashboard']
    for row in range(72, 87):
        for col in range(7, 15):
            cell = dashboard.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None
    dashboard['G71'] = 'Patient rating distribution by service type'
    dashboard['K71'] = 'Patient rating five-number summary'
    set_caption(dashboard, 'A11', 'Daily record volume during June. The dip on weekend dates is visible in the source record counts.')
    set_caption(dashboard, 'A26', 'Ten most frequent recorded disease labels. These are dataset categories, not independently verified diagnoses.')
    set_caption(dashboard, 'A41', 'Average treatment cost per patient-day as defined in the workbook; compare with the same denominator.')
    dashboard.row_dimensions[11].height = 27
    dashboard.row_dimensions[26].height = 27
    dashboard.row_dimensions[41].height = 27
    add_line_chart(dashboard, plot_data, blocks['daily'], 'Daily patient records · 1–30 June 2026', 'A12', width=24, height=7.3, y_title='Records')
    add_bar_chart(dashboard, plot_data, blocks['diseases'], 'Top 10 diseases by records', 'A27', horizontal=True, width=24, height=7.3, x_title='Records')
    add_bar_chart(dashboard, plot_data, blocks['cost_day'], 'Average treatment cost per patient-day', 'A42', horizontal=True, width=24, height=7.3, x_title='INR per patient-day', colors=[BLUE, TEAL])
    add_bar_chart(dashboard, plot_data, blocks['ratings_by_type'], 'Rating distribution by patient type', 'G72', width=12.5, height=8.0, y_title='Records', colors=[BLUE, TEAL])
    add_bar_chart(dashboard, plot_data, blocks['rating_summary'], 'Rating five-number summary by type', 'K72', width=12.5, height=8.0, y_title='Rating', colors=[BLUE, TEAL, AMBER, PURPLE, GREEN, SLATE])
    set_caption(dashboard, 'G87', 'Grouped columns show record counts at each rating for OPD and IPD.')
    set_caption(dashboard, 'K87', 'This compact summary shows min, quartiles, median, and max. The Plotly tab includes the box plot.')
    dashboard.row_dimensions[87].height = 27

    # A dedicated tab holds the remaining classic charts; the richer Plotly versions remain in the web dashboard.
    if 'Visual Analysis' in wb.sheetnames:
        del wb['Visual Analysis']
    visual = wb.create_sheet('Visual Analysis')
    visual.sheet_view.showGridLines = False
    visual.merge_cells('A1:W1')
    visual['A1'] = 'Hospital Patient Analytics · supplementary charts'
    visual['A1'].font = Font(bold=True, color=WHITE, size=17)
    visual['A1'].fill = PatternFill('solid', fgColor=NAVY)
    visual['A2'] = 'Static workbook charts for Google Sheets compatibility. Use the Plotly dashboard for date animation, interactive scatter, violin, and animated views.'
    visual.merge_cells('A2:W2')
    visual['A2'].font = Font(italic=True, color=SLATE)
    visual['A2'].alignment = Alignment(wrap_text=True)
    visual.column_dimensions['A'].width = 18
    for c in range(2, 25):
        visual.column_dimensions[get_column_letter(c)].width = 13
    chart_blocks = [
        ('Emergency patients by insurance', blocks['emergency'], 'June average of emergency records per day by recorded insurance status.', 'bar', True, False, [TEAL, AMBER], 'Patients per day', None),
        ('Average treatment cost by rating', blocks['cost_rating'], 'Mean treatment cost across rating groups; differences are descriptive and not causal.', 'line', False, False, [TEAL], 'Average cost (INR)', None),
        ('Average service time', blocks['service'], 'OPD waiting time and IPD room-allocation time are distinct processes.', 'bar', False, False, [BLUE, AMBER], 'Minutes', None),
        ('OPD and IPD record distribution', blocks['patient_type'], 'Record volume by service type. Counts refer to records, not unique people.', 'pie', False, True, None, None, None),
        ('Recorded patient outcomes', blocks['outcomes'], 'Counts by recorded patient status; interpret labels using the source definitions.', 'pie', False, True, None, None, None),
        ('Records by severity', blocks['severity'], 'Count of records in each severity category; not a quality or risk-adjusted measure.', 'bar', True, False, [TEAL], 'Records', None),
        ('Average treatment cost by payment method', blocks['payment_cost'], 'Average recorded cost for each payment category; case mix may differ across groups.', 'bar', True, False, [BLUE], 'Average cost (INR)', None),
        ('Records by payment method', blocks['payments'], 'Distribution of records across the recorded payment categories.', 'bar', True, False, [BLUE], 'Records', None),
        ('Insurance coverage records', blocks['insurance'], 'Records marked insured or uninsured; this does not show claim approval.', 'pie', False, True, None, None, None),
        ('Age distribution by gender · overlapping histogram', blocks['age'], 'June age bands in ten-year bins. Transparent bars overlap to compare gender distributions.', 'bar', False, False, [BLUE, AMBER, TEAL, PURPLE], 'Age band', 'Records'),
        ('Patient type and severity · June snapshot', blocks['type_severity'], 'Static June totals for the categories that animate by date in the Plotly app.', 'bar', False, False, [GREEN, BLUE, AMBER, 'B95C56'], 'Patient type', 'Records'),
        ('Blood group by severity', blocks['blood_severity'], 'Severity counts by blood group. Categorical counts are shown as stacked bars in this workbook.', 'bar', False, False, [GREEN, BLUE, AMBER, 'B95C56'], 'Blood group', 'Records'),
    ]
    placements = []
    for idx, (title, block, caption, kind, horizontal, donut, colors, xt, yt) in enumerate(chart_blocks):
        group_row = 4 + (idx // 2) * 19
        group_col = 1 if idx % 2 == 0 else 12
        capcell = visual.cell(group_row, group_col, caption)
        capcell.font = Font(color=SLATE, italic=True, size=10)
        capcell.alignment = Alignment(wrap_text=True, vertical='center')
        visual.merge_cells(start_row=group_row, start_column=group_col, end_row=group_row, end_column=group_col+9)
        chart_anchor = f'{visual.cell(group_row+1,group_col).coordinate}'
        if kind == 'pie':
            chart = add_pie_chart(visual, plot_data, block, title, chart_anchor, doughnut=donut, width=15.5, height=8.5)
        elif kind == 'line':
            chart = add_line_chart(visual, plot_data, block, title, chart_anchor, width=15.5, height=8.5, y_title=xt, number_format='₹#,##0')
        else:
            chart = add_bar_chart(visual, plot_data, block, title, chart_anchor, horizontal=horizontal, stacked='stacked' if idx == 11 else None, width=15.5, height=8.5, colors=colors, x_title=xt, y_title=yt)
        placements.append((title, chart_anchor, type(chart).__name__))
        visual.row_dimensions[group_row].height = 32

    # The association sheet uses color scale cells rather than a chart object: this format is supported by Sheets.
    visual['A61'] = 'Correlation/association map is on the Association Map tab.'
    visual['A61'].font = Font(bold=True, color=NAVY)
    visual['A62'] = 'Cramér’s V and correlation ratio (η) measure association, not causality; see notes in that tab.'
    visual['A62'].font = Font(color=SLATE, italic=True)

    # Add/update provenance note without altering any patient data.
    if 'Lists' in wb.sheetnames:
        wb['Lists'].sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = 'auto'
    wb.active = wb.sheetnames.index('Dashboard')
    OUTPUT_XLSX.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_XLSX)

    # Reopen the saved artifact and recheck data integrity and chart package parts.
    check = load_workbook(OUTPUT_XLSX, read_only=False, data_only=False)
    validate_patient_sheet(check['Patient Data'], SOURCE_CSV)
    chart_counts = {ws.title: len(ws._charts) for ws in check.worksheets}
    total_charts = sum(chart_counts.values())
    if chart_counts.get('Dashboard') != 5 or chart_counts.get('Visual Analysis') != 12:
        raise ValueError(f'Unexpected chart count after saving: {chart_counts}')
    if total_charts != 17:
        raise ValueError(f'Expected 17 chart objects, found {total_charts}')
    print('OUTPUT', OUTPUT_XLSX)
    print('SOURCE_PATIENT_ROWS', len(source_rows), 'COLUMNS', len(headers))
    print('JUNE_ROWS', len(june_rows), 'MISSING_REQUIRED_CHART_FIELDS', missing_by_field)
    print('CHART_COUNTS', chart_counts, 'TOTAL', total_charts)
    print('ASSOCIATION_ROWS', len(matrix), 'AGE_BINS', len(age_labels), 'GENDERS', genders)
    print('CHART_TITLES')
    for title, anchor, chart_type in placements:
        print(f'  {chart_type}: {title} at {anchor}')


if __name__ == '__main__':
    main()
