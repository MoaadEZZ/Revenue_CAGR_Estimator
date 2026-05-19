from flask import Flask, render_template, request, jsonify, send_file, redirect
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pipeline import *
import os
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd

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
    

@app.route('/files')
def files():
    """File management page"""
    return render_template('files.html')


@app.route('/settings')
def settings():
    """Settings page (placeholder for future use)"""
    return render_template('settings.html')


@app.route('/file_status')
def file_status():
    """Get current file status and metadata"""
    try:
        files_info = {}
        warnings = []

        # Check CAGR Mapper
        if os.path.exists(DATA_CAGR_MAPPER):
            stat = os.stat(DATA_CAGR_MAPPER)
            mod_time = datetime.fromtimestamp(stat.st_mtime)
            days_old = (datetime.now() - mod_time).days
            
            try:
                df = pd.read_excel(DATA_CAGR_MAPPER)
                files_info['cagr'] = {
                    'modified': mod_time.strftime('%Y-%m-%d %H:%M'),
                    'rows': len(df),
                    'size': f"{stat.st_size / 1024:.1f} KB",
                    'days_old': days_old
                }
                
                if days_old > 30:
                    warnings.append(f"⚠️ CAGR Mapper is {days_old} days old. Consider updating it.")
                elif days_old > 14:
                    warnings.append(f"⚡ CAGR Mapper is {days_old} days old. Update recommended soon.")
            except Exception as e:
                print(f"Error reading CAGR file: {e}")

        # Check Market Size
        if os.path.exists(DATA_MARKET_SIZE):
            stat = os.stat(DATA_MARKET_SIZE)
            mod_time = datetime.fromtimestamp(stat.st_mtime)
            days_old = (datetime.now() - mod_time).days
            
            try:
                df = pd.read_excel(DATA_MARKET_SIZE)
                files_info['market'] = {
                    'modified': mod_time.strftime('%Y-%m-%d %H:%M'),
                    'rows': len(df),
                    'size': f"{stat.st_size / 1024:.1f} KB",
                    'days_old': days_old
                }
                
                if days_old > 30:
                    warnings.append(f"⚠️ Market Size file is {days_old} days old. Consider updating it.")
                elif days_old > 14:
                    warnings.append(f"⚡ Market Size file is {days_old} days old. Update recommended soon.")
            except Exception as e:
                print(f"Error reading Market Size file: {e}")

        # ✅ ALWAYS return valid structure
        print(files_info)
        return jsonify({
            'files': files_info if files_info else {},
            'warnings': warnings
        })

    except Exception as e:
        print(f"Error in file_status: {e}")
        return jsonify({
            'files': {},
            'warnings': [f'Error loading file status: {str(e)}']
        }), 500



@app.route('/file_history')
def file_history():
    """Get version history of uploaded files"""
    try:
        history = []
        
        # Check backup directories
        backup_dirs = {
            'cagr': os.path.join(DATA_DIR, 'backups', 'cagr'),
            'market': os.path.join(DATA_DIR, 'backups', 'market')
        }
        
        for file_type, backup_dir in backup_dirs.items():
            if os.path.exists(backup_dir):
                files = sorted(os.listdir(backup_dir), reverse=True)
                for filename in files[:5]:  # Last 5 versions
                    filepath = os.path.join(backup_dir, filename)
                    if os.path.isfile(filepath):
                        try:
                            df = pd.read_excel(filepath)
                            stat = os.stat(filepath)
                            mod_time = datetime.fromtimestamp(stat.st_mtime)
                            
                            history.append({
                                'file_type': file_type,
                                'timestamp': mod_time.strftime('%Y-%m-%d %H:%M:%S'),
                                'rows': len(df),
                                'filename': filename
                            })
                        except:
                            pass
        
        # Sort by timestamp descending
        history = sorted(history, key=lambda x: x['timestamp'], reverse=True)
        
        return jsonify({'history': history})

    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/upload_file', methods=['POST'])
def upload_file():
    """Handle file uploads with validation"""
    try:
        file_type = request.form.get('file_type')
        uploaded_file = request.files.get('file')
        
        if not uploaded_file:
            return jsonify({'error': 'No file provided'}), 400
        
        if not file_type:
            return jsonify({'error': 'File type not specified'}), 400
        
        # ✅ Read file into memory FIRST
        try:
            file_bytes = uploaded_file.read()
            uploaded_file.seek(0)  # Reset stream
            df = pd.read_excel(uploaded_file)
        except Exception as e:
            return jsonify({'error': f'Invalid Excel file: {str(e)}'}), 400
        
        # Validate required columns based on file type
        if file_type == 'cagr':
            required_cols = ['catégorie', 'Sub-segment 2', 'Segment', 'pr_offer']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                return jsonify({
                    'error': f'Missing required columns: {", ".join(missing)}'
                }), 400
            target_path = DATA_CAGR_MAPPER
            
        elif file_type == 'market':
            required_cols = ['Year', 'Million EUR']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                return jsonify({
                    'error': f'Missing required columns: {", ".join(missing)}'
                }), 400
            target_path = DATA_MARKET_SIZE
            
        else:
            return jsonify({'error': 'Invalid file type'}), 400
        
        # Create backup of old file
        if os.path.exists(target_path):
            backup_old_file(file_type, target_path)
        
        # ✅ Save file properly using BytesIO
        try:
            with open(target_path, 'wb') as f:
                f.write(file_bytes)
            
            # Verify file was written correctly
            if not os.path.exists(target_path) or os.path.getsize(target_path) == 0:
                raise Exception("File was not saved properly")
                
        except Exception as e:
            return jsonify({'error': f'Failed to save file: {str(e)}'}), 500
        
        # Reload global data
        global revenue_years, market_size_years, df_all_data
        revenue_years = load_revenue_data()
        market_size_years = load_market_size_years()
        df_all_data = load_all_data()
        
        return jsonify({
            'success': True,
            'message': f'{file_type} file uploaded successfully',
            'rows': len(df),
            'timestamp': datetime.now().isoformat()
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500



def backup_old_file(file_type, source_path):
    """Keep previous version as backup"""
    try:
        if file_type == 'cagr':
            backup_dir = os.path.join(DATA_DIR, 'backups', 'cagr')
        else:
            backup_dir = os.path.join(DATA_DIR, 'backups', 'market')
        
        os.makedirs(backup_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = os.path.join(backup_dir, f'backup_{timestamp}.xlsx')
        
        if os.path.exists(source_path):
            shutil.copy2(source_path, backup_path)
            print(f"✓ Backup created: {backup_path}")
    except Exception as e:
        print(f"⚠ Warning: Could not create backup: {e}")


@app.route('/back')
def back():
    return redirect('/')


if __name__ == '__main__':
    app.run(debug=True)
