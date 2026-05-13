from flask import Flask, render_template, request, jsonify, send_file, redirect
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pipeline import *

app = Flask(__name__)

revenue_years     = []
market_size_years = []
df_all_data       = None

# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.route('/')
def index():
    global revenue_years, market_size_years, df_all_data
    if not revenue_years:    revenue_years    = load_revenue_data()
    if not market_size_years: market_size_years = load_market_size_years()
    if df_all_data is None:  df_all_data = load_all_data()

    svp_options = [
        '1.SVP - Employee Experience',
        '2.SVP - Digital Infrastructure',
        '3.SVP - Operational Experience',
        '4.SVP - Customer Experience',
    ]
    regions = ['France', 'International']

    default_display_start = revenue_years[0]    if revenue_years    else market_size_years[0]
    default_display_end   = market_size_years[-1] if market_size_years else revenue_years[-1]
    default_cagr_start    = revenue_years[-1]   if revenue_years    else market_size_years[0]
    default_cagr_end      = market_size_years[-1] if market_size_years else revenue_years[-1]
    # Default pred start = last revenue year (so overlap starts there)
    default_pred_start    = revenue_years[-1]   if revenue_years    else market_size_years[0]

    return render_template('index.html',
        revenue_years=revenue_years,
        market_size_years=market_size_years,
        svp_options=svp_options,
        regions=regions,
        default_display_start=default_display_start,
        default_display_end=default_display_end,
        default_cagr_start=default_cagr_start,
        default_cagr_end=default_cagr_end,
        default_pred_start=default_pred_start,
    )


@app.route('/results_data', methods=['POST'])
def results_data():
    cagr_start_year    = int(request.form.get('cagr_start_year'))
    cagr_end_year      = int(request.form.get('cagr_end_year'))
    display_start_year = int(request.form.get('display_start_year'))
    display_end_year   = int(request.form.get('display_end_year'))
    pred_start_year    = int(request.form.get('pred_start_year'))
    selected_svps      = request.form.getlist('svp')
    region             = request.form.get('region')
    result = resultat_data(df_all_data=df_all_data,
                           display_start_year=display_start_year, 
                           display_end_year=display_end_year, 
                           cagr_start_year=cagr_start_year, 
                           cagr_end_year=cagr_end_year, 
                           pred_start_year=pred_start_year, 
                           selected_svps=selected_svps, 
                           region=region)
    return jsonify(result)

@app.route('/download_excel', methods=['POST'])
def download_excel():
    try:
        data        = request.json
        table_data  = data.get('data', [])
        columns     = data.get('columns', [])
        disp_start  = data.get('display_start_year', '')
        disp_end    = data.get('display_end_year', '')

        if not table_data or not columns:
            return jsonify({'error': 'No data to export'}), 400

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Analysis {disp_start}-{disp_end}"

        hdr_fill  = PatternFill(start_color="667EEA", end_color="667EEA", fill_type="solid")
        hdr_font  = Font(bold=True, color="FFFFFF", size=11)
        pred_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        real_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        tot_fill  = PatternFill(start_color="E7E7FF", end_color="E7E7FF", fill_type="solid")
        tot_font  = Font(bold=True, color="667EEA", size=10)
        border    = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC'))

        for ci, col in enumerate(columns, 1):
            cell = ws.cell(row=1, column=ci, value=col)
            cell.fill = hdr_fill; cell.font = hdr_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        for ri, row_data in enumerate(table_data, 2):
            is_total = row_data.get('Offer') == 'TOTAL'
            for ci, col in enumerate(columns, 1):
                value = row_data.get(col, '')
                cell  = ws.cell(row=ri, column=ci, value=value)
                cell.border = border
                if is_total:
                    cell.fill = tot_fill; cell.font = tot_font
                elif col.startswith('Predicted_'):
                    cell.fill = pred_fill
                elif col.startswith('Real_'):
                    cell.fill = real_fill
                if col.startswith(('Revenue_','Real_','Predicted_')):
                    if isinstance(value, (int, float)):
                        cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal='right')
                elif col.startswith('CAGR_'):
                    cell.alignment = Alignment(horizontal='center')
                else:
                    cell.alignment = Alignment(horizontal='left')

        for ci, col in enumerate(columns, 1):
            mx = max((len(str(ws.cell(row=r, column=ci).value or ''))
                      for r in range(1, len(table_data)+2)), default=10)
            ws.column_dimensions[get_column_letter(ci)].width = min(mx + 2, 50)

        ws.freeze_panes = 'A2'
        out = BytesIO(); wb.save(out); out.seek(0)
        return send_file(out,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'SVP_Analysis_{disp_start}_{disp_end}.xlsx')

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/back')
def back():
    return redirect('/')


if __name__ == '__main__':
    app.run(debug=True)
