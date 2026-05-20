import pandas as pd
import numpy as np
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

DATA_CAGR_MAPPER  = os.path.join(DATA_DIR, "CAGR Mapper.xlsx")
DATA_MARKET_SIZE  = os.path.join(DATA_DIR, "Market Sizing - 3Q 2025.xlsx")
DATA_ALL_PRODUCTS = os.path.join(DATA_DIR, "All products requested.xlsx")
DATA_REVENUE      = os.path.join(DATA_DIR, "data.xlsx") \
                    if os.path.exists(os.path.join(DATA_DIR, "data.xlsx")) else None

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def calculate_cagr(start_value, end_value, years_diff):
    if start_value > 0 and years_diff > 0:
        return (end_value / start_value) ** (1 / years_diff) - 1
    return 0


# ─────────────────────────────────────────────
# Data loaders
# ─────────────────────────────────────────────

def load_revenue_data():
    try:
        df = pd.read_excel(DATA_REVENUE)
        if not all(c in df.columns for c in ["Offer", "Period", "_S_Revenue_actual"]):
            return []
        df['Period'] = pd.to_datetime(df['Period'], errors='coerce')
        df = df.dropna(subset=['Period'])
        years = sorted(df['Period'].dt.year.unique().tolist())
        print(f"✓ Revenue years: {years}")
        return years
    except Exception as e:
        print(f"Error loading revenue data: {e}")
        return []

def load_custom_cagr():
    """Load custom CAGR file (offer based)
    
    Returns:
        dict: {offer_name: cagr_value} or None if file doesn't exist
        
    Behavior:
        - Only offers with valid CAGR values are included
        - Offers with missing/null CAGR are skipped (won't overwrite)
        - Column names are case-insensitive
    """
    try:
        filepath = os.path.join(DATA_DIR, 'custom_cagr.xlsx')
        
        if not os.path.exists(filepath):
            print("ℹ No custom CAGR file found")
            return None
        
        df = pd.read_excel(filepath)
        
        # Normalize column names: strip whitespace and convert to title case
        df.columns = df.columns.str.strip().str.title()
        
        # Find columns (case-insensitive)
        offer_col = None
        cagr_col = None
        
        for col in df.columns:
            if col.lower() == 'offer':
                offer_col = col
            elif col.lower() == 'cagr':
                cagr_col = col
        
        if not offer_col or not cagr_col:
            print(f"⚠ Custom CAGR: Missing required columns. Found: {list(df.columns)}")
            return None
        
        # Create result dictionary
        custom_cagr_dict = {}
        skipped_count = 0
        
        for idx, row in df.iterrows():
            offer = row[offer_col]
            cagr_value = row[cagr_col]
            
            # Skip if offer is missing
            if pd.isna(offer) or offer == '':
                skipped_count += 1
                continue
            
            # Skip if CAGR is missing or null
            if pd.isna(cagr_value) or cagr_value == '':
                print(f"  ⚠ Skipping {offer}: CAGR value is missing")
                skipped_count += 1
                continue
            
            # Convert to float
            try:
                cagr_float = float(cagr_value)
                offer_str = str(offer).strip()
                custom_cagr_dict[offer_str] = cagr_float
            except ValueError:
                print(f"  ⚠ Skipping {offer}: CAGR '{cagr_value}' is not numeric")
                skipped_count += 1
                continue
        
        if custom_cagr_dict:
            print(f"✓ Custom CAGR loaded: {len(custom_cagr_dict)} offers")
            if skipped_count > 0:
                print(f"  ({skipped_count} rows skipped due to missing/invalid values)")
            return custom_cagr_dict
        else:
            print("⚠ Custom CAGR file has no valid entries")
            return None
            
    except Exception as e:
        print(f"✗ Error loading custom CAGR: {e}")
        return None

def load_market_size_years():
    try:
        df = pd.read_excel(DATA_MARKET_SIZE,
                           sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
        years = sorted(df['Year'].unique().tolist())
        print(f"✓ Market size years: {years}")
        return years
    except Exception as e:
        print(f"Error loading market size years: {e}")
        return []


def load_cagr_mapper():
    try:
        df = pd.read_excel(DATA_CAGR_MAPPER)
        df.columns = df.columns.str.strip()
        required = ['catégorie', 'Sub-segment 2', 'Segment', 'pr_offer']
        if not all(c in df.columns for c in required):
            return None
        df = df[required].copy()
        df = df.dropna(subset=['pr_offer'])
        df = df[df['pr_offer'].apply(lambda x: isinstance(x, str))]
        for col in required:
            df[col] = df[col].astype(str).str.rstrip()
        if "." in df["pr_offer"].iloc[0]:
            df["pr_offer"] = df["pr_offer"].str.split(".", n=1, expand=True)[1]
        df = df.rename(columns={'pr_offer': 'Offer', 'catégorie': 'Market_Category'})
        df = df.drop_duplicates(subset=['Offer'], keep='first')
        print(f"✓ CAGR Mapper: {len(df)} offers")
        return df
    except Exception as e:
        print(f"Error loading CAGR mapper: {e}")
        return None


def load_all_data():
    try:
        # ── SVP mapping ──
        df_prod = pd.read_excel(DATA_ALL_PRODUCTS)
        svp_col = next((c for c in ['Strategic','SVP','Strategic Value Proposition','SVP Category']
                        if c in df_prod.columns), None)
        if svp_col:
            df_prod['Offer'] = df_prod['Offer'].astype(str).str.strip()
            if df_prod['Offer'].str.contains('.', regex=False).any():
                df_prod['Offer_Clean'] = df_prod['Offer'].str.split('.', n=1, expand=True)[1].str.strip()
            else:
                df_prod['Offer_Clean'] = df_prod['Offer']
            offer_to_svp = dict(zip(df_prod['Offer_Clean'], df_prod[svp_col]))
        else:
            offer_to_svp = {}

        # ── CAGR mapper ──
        df_mapper = load_cagr_mapper()
        if df_mapper is None:
            return None

        # ── Market size ──
        df_ms = pd.read_excel(DATA_MARKET_SIZE,
                              sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
        df_ms = df_ms[["Segment","Sub-segment 2","Year","Million EUR"]].copy()
        df_ms["Sub-segment 2"] = df_ms["Sub-segment 2"].astype(str).str.rstrip()
        df_ms = df_ms.dropna(subset=["Sub-segment 2","Year","Million EUR"])
        df_ms = df_ms.groupby(["Sub-segment 2","Year"], as_index=False).agg(
            {"Million EUR":"sum","Segment":"first"})
        df_ms['Year'] = df_ms['Year'].astype(int)
        df_ms_pivot = df_ms.pivot_table(
            index=["Sub-segment 2","Segment"], columns="Year",
            values="Million EUR", fill_value=0).reset_index()
        df_ms_pivot.columns = [
            f"Million EUR_{c}" if isinstance(c, int) else c
            for c in df_ms_pivot.columns]

        # ── Revenue ──
        df_rev = pd.read_excel(DATA_REVENUE)
        if 'Region' not in df_rev.columns:
            df_rev['Region'] = 'FR'
        df_rev = df_rev[["Offer","Period","_S_Revenue_actual","Region"]]
        df_rev['Period'] = pd.to_datetime(df_rev['Period'], errors='coerce')
        df_rev = df_rev.dropna(subset=['Period'])
        df_rev['Year'] = df_rev['Period'].dt.year
        if df_rev["_S_Revenue_actual"].dtype == object:
            df_rev["_S_Revenue_actual"] = (df_rev["_S_Revenue_actual"]
                .astype(str).str.replace(r"[,$]","",regex=True))
        df_rev["_S_Revenue_actual"] = pd.to_numeric(
            df_rev["_S_Revenue_actual"], errors='coerce').fillna(0)

        df_rev_grp = df_rev.groupby(['Offer','Year','Region'], as_index=False)[
            '_S_Revenue_actual'].sum()
        df_rev_piv = df_rev_grp.pivot_table(
            index=['Offer','Region'], columns='Year',
            values='_S_Revenue_actual', fill_value=0).reset_index()
        df_rev_piv.columns = [
            f"Revenue_{c}" if isinstance(c, int) else c
            for c in df_rev_piv.columns]

        # ── Merge ──
        df = df_rev_piv.merge(df_mapper, on='Offer', how='inner')
        df['SVP'] = df['Offer'].map(offer_to_svp).fillna('Unknown')
        df = df.merge(df_ms_pivot, on='Sub-segment 2', how='left')
        print(f"✓ Final dataset: {len(df)} rows")
        return df

    except Exception as e:
        print(f"Error in load_all_data: {e}")
        import traceback; traceback.print_exc()
        return None

def resultat_data(df_all_data, display_start_year, display_end_year, cagr_start_year, cagr_end_year,
                  pred_start_year, selected_svps, region, use_custom_cagr='no'):

    try:
        print(f"\n{'='*60}")
        print(f"Display : {display_start_year} → {display_end_year}")
        print(f"CAGR    : {cagr_start_year} → {cagr_end_year}")
        print(f"Pred from: {pred_start_year}")
        print(f"SVPs    : {selected_svps}  Region: {region}")
        print(f"Custom CAGR: {use_custom_cagr}")
        print(f"{'='*60}")

        # Load custom CAGR if requested
        custom_cagr_dict = None    
        if use_custom_cagr == 'yes':
            custom_cagr_dict = load_custom_cagr()
            if custom_cagr_dict:
                print(f"\n📊 Applying custom CAGR to {len(custom_cagr_dict)} offers\n")

        if not selected_svps:
            return {'error': 'Please select at least one SVP category'}
        if df_all_data is None:
            return {'error': 'Data not loaded.'}

        df = df_all_data[df_all_data['SVP'].isin(selected_svps)].copy()
        if region == 'France':
            df = df[df['Region'] == 'FR']
        elif region == 'International':
            df = df[df['Region'] == 'INT']
        if len(df) == 0:
            return {'error': 'No data found for selected criteria.'}

        # Revenue years present in the data
        rev_cols  = sorted([c for c in df.columns if c.startswith('Revenue_')])
        rev_years = [int(c.split('_')[1]) for c in rev_cols]
        last_rev_year = rev_years[-1] if rev_years else display_start_year

        all_display_years = list(range(display_start_year, display_end_year + 1))

        # ── Per-offer computation ──────────────────────────────────────────

        results_list = []
        cagr_values  = []

        # Totals for chart
        tot_actual    = {y: 0.0 for y in all_display_years}
        tot_real_post = {y: 0.0 for y in all_display_years}
        tot_predicted = {y: 0.0 for y in all_display_years}
        has_real_post = {y: False for y in all_display_years}

        custom_cagr_applied_count = 0

        for _, row in df.iterrows():

            offer           = row['Offer']
            market_category = row.get('Market_Category', '---')
            svp             = row.get('SVP', '---')
            sub2            = row['Sub-segment 2']

            # ✅ Check for custom CAGR first
            cagr = None
            cagr_source = 'market'
            
            if use_custom_cagr == 'yes' and custom_cagr_dict:
                # Exact match lookup
                if offer in custom_cagr_dict:
                    cagr = custom_cagr_dict[offer]
                    cagr_source = 'custom'
                    custom_cagr_applied_count += 1
                    print(f"  ✓ {offer}: Custom CAGR = {cagr*100:.2f}%")

            if sub2 is None or (isinstance(sub2, float) and np.isnan(sub2)):
                sub2 = '---'

            # If no custom CAGR, calculate from market size
            if cagr is None:
                sc = f"Million EUR_{cagr_start_year}"
                ec = f"Million EUR_{cagr_end_year}"

                if sc in df.columns and ec in df.columns:
                    ms_s = row[sc] if not pd.isna(row[sc]) else 0
                    ms_e = row[ec] if not pd.isna(row[ec]) else 0
                    cagr = calculate_cagr(ms_s, ms_e, cagr_end_year - cagr_start_year)
                else:
                    cagr = 0
                    print(f"⚠ Warning: Market size data missing for {cagr_start_year}-{cagr_end_year}")

            cagr_values.append(cagr)

            # Actual revenues from data
            actual = {}
            for y in rev_years:
                col = f"Revenue_{y}"
                if col in row.index:
                    v = row[col]
                    actual[y] = float(v) if not pd.isna(v) else 0.0

            # Seed for projection = revenue at (pred_start_year - 1)
            seed_y = pred_start_year
            seed_v = actual.get(seed_y, None)

            # If pred_start_year has no data, use last available revenue year
            if seed_v is None:
                seed_v = actual.get(last_rev_year, 0.0)
                seed_y = last_rev_year

            result_row = OrderedDict()
            result_row['Offer']          = offer
            result_row['SVP']            = svp
            result_row['Sub-segment 2']  = sub2
            result_row['Market_Category']= market_category
            result_row[f'CAGR_{cagr_start_year}/{cagr_end_year}'] = f"{cagr*100:.2f}%"

            anchor_val = tot_predicted.get(pred_start_year, 0)

            for y in all_display_years:

                act_val  = actual.get(y, None)
                pred_val = None

                if y < pred_start_year:
                    if act_val is not None:
                        tot_predicted[y] += act_val
                    else:
                        tot_predicted[y] = 0

                if y >= pred_start_year:
                    if y == pred_start_year:
                        pred_val = act_val if act_val is not None else seed_v * ((1 + cagr) ** (pred_start_year - seed_y))
                    else:
                        steps    = y - seed_y
                        pred_val = seed_v * ((1 + cagr) ** steps)

                if act_val is not None and pred_val is not None:
                    result_row[f'Real_{y}']      = f"{act_val:,.2f}"
                    result_row[f'Predicted_{y}'] = f"{pred_val:,.2f}"
                    tot_actual[y]    += act_val
                    tot_real_post[y] += act_val
                    tot_predicted[y] += pred_val
                    has_real_post[y]  = True
                elif pred_val is not None:
                    result_row[f'Predicted_{y}'] = f"{pred_val:,.2f}"
                    tot_actual[y]    += pred_val
                    tot_predicted[y] += pred_val
                else:
                    v = act_val if act_val is not None else 0.0
                    result_row[f'Revenue_{y}'] = f"{v:,.2f}"
                    tot_actual[y] += v

            results_list.append(result_row)

        print(f"\n✓ Custom CAGR applied to {custom_cagr_applied_count} offers")
        print(f"{'='*60}\n")

        # Sort
        results_list = sorted(results_list, key=lambda x: x['Offer'])

        # ── Column order ───────────────────────────────────────────────────

        fixed_cols = ['Offer','SVP','Sub-segment 2','Market_Category',
                      f'CAGR_{cagr_start_year}/{cagr_end_year}']

        year_cols  = []

        for y in all_display_years:
            if any(f'Revenue_{y}' in r for r in results_list):
                year_cols.append(f'Revenue_{y}')
            if any(f'Real_{y}' in r for r in results_list):
                year_cols.append(f'Real_{y}')
            if any(f'Predicted_{y}' in r for r in results_list):
                year_cols.append(f'Predicted_{y}')
        column_order = fixed_cols + year_cols

        # ── TOTAL row ──────────────────────────────────────────────────────

        total_row = OrderedDict()
        total_row['Offer']           = 'TOTAL'
        total_row['SVP']             = ''
        total_row['Sub-segment 2']   = ''
        total_row['Market_Category'] = ''
        total_row[f'CAGR_{cagr_start_year}/{cagr_end_year}'] = ''

        for col in year_cols:
            y = int(col.split('_')[-1])
            if col.startswith('Revenue_'):
                total_row[col] = f"{tot_actual[y]:,.2f}"
            elif col.startswith('Real_'):
                total_row[col] = f"{tot_real_post[y]:,.2f}"
            elif col.startswith('Predicted_'):
                total_row[col] = f"{tot_predicted[y]:,.2f}"
        results_list.append(total_row)

        # ── KPIs ───────────────────────────────────────────────────────────

        avg_cagr   = np.mean(cagr_values) if cagr_values else 0

        total_row_data = results_list[-1]

        rev_start_col = f'Revenue_{display_start_year}'
        rev_end_col   = f'Real_{pred_start_year}'

        Actual_cagr = 0

        if rev_start_col in total_row_data and rev_end_col in total_row_data:
            rev_start_str = total_row_data[rev_start_col]
            rev_end_str   = total_row_data[rev_end_col]
            
            if rev_start_str and rev_start_str != '':
                rev_start = float(rev_start_str.replace(',', ''))
                rev_end   = float(rev_end_str.replace(',', ''))
                Actual_cagr = calculate_cagr(rev_start, rev_end, pred_start_year - display_start_year)

        # ── Chart arrays ───────────────────────────────────────────────────────────

        chart_actual     = []
        chart_real_post  = []
        chart_predicted  = []

        for y in all_display_years:
            if y < pred_start_year:
                chart_actual.append(tot_actual[y])
                chart_real_post.append(None)
                chart_predicted.append(None)
            elif y == pred_start_year:
                seed_value = tot_actual[y]
                chart_actual.append(seed_value)
                chart_real_post.append(seed_value)
                chart_predicted.append(seed_value)
            else:
                chart_actual.append(None)
                if has_real_post[y]:
                    chart_real_post.append(tot_real_post[y])
                else:
                    chart_real_post.append(None)
                chart_predicted.append(tot_predicted[y] if tot_predicted[y] > 0 else None)

        return {
            'cagr_start_year'    : cagr_start_year,
            'cagr_end_year'      : cagr_end_year,
            'display_start_year' : display_start_year,
            'display_end_year'   : display_end_year,
            'pred_start_year'    : pred_start_year,
            'avg_cagr'           : f"{avg_cagr*100:.2f}%",
            'tot_predicted'      : tot_predicted,
            'actual_cagr'        : f"{Actual_cagr*100:.2f}%",
            'result_count'       : len(results_list) - 1,
            'chart_labels'       : [str(y) for y in all_display_years],
            'chart_actual'       : chart_actual,
            'chart_real_post'    : chart_real_post,
            'chart_predicted'    : chart_predicted,
            'table_data'         : results_list,
            'column_order'       : column_order,
        }

    except Exception as e:
        import traceback; traceback.print_exc()
        return {'error': str(e)}
