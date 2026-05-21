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
        print("\n" + "="*60)
        print("LOADING ALL DATA")
        print("="*60)
        
        # ── SVP mapping ──
        print(f"\n📊 Loading SVP mapping from All Products...")
        df_prod = pd.read_excel(DATA_ALL_PRODUCTS)
        print(f"  ✓ Loaded {len(df_prod)} rows")
        print(f"  Available columns: {list(df_prod.columns)}")
        
        # ✅ Find SVP column (case-insensitive)
        svp_col = None
        for col in df_prod.columns:
            if col.lower() in ['strategic', 'svp', 'strategic value proposition', 'svp category']:
                svp_col = col
                break
        
        print(f"  SVP column found: {svp_col}")
        
        # ✅ Prepare additional columns
        df_prod['Offer'] = df_prod['Offer'].astype(str).str.strip()
        if df_prod['Offer'].str.contains('.', regex=False).any():
            df_prod['Offer_Clean'] = df_prod['Offer'].str.split('.', n=1, expand=True)[1].str.strip()
        else:
            df_prod['Offer_Clean'] = df_prod['Offer']
        
        # Map SVP
        if svp_col:
            offer_to_svp = dict(zip(df_prod['Offer_Clean'], df_prod[svp_col]))
            print(f"  ✓ Mapped {len(offer_to_svp)} offers to SVP")
        else:
            offer_to_svp = {}
            print(f"  ⚠ No SVP column found")
        
        # ✅ Map Business Line (check multiple possible column names)
        offer_to_business_line = {}
        business_line_col = None
        for col in df_prod.columns:
            if col.lower() in ['business line', 'business_line', 'businessline']:
                business_line_col = col
                break
        
        if business_line_col:
            offer_to_business_line = dict(zip(df_prod['Offer_Clean'], df_prod[business_line_col]))
            print(f"  ✓ Mapped {len(offer_to_business_line)} offers to Business_Line")
        else:
            print(f"  ⚠ No 'Business_Line' column found")
        
        # ✅ Map Domain (check multiple possible column names)
        offer_to_domain = {}
        domain_col = None
        for col in df_prod.columns:
            if col.lower() in ['domain', 'sub_domain', 'subdomain']:
                domain_col = col
                break
        
        if domain_col:
            offer_to_domain = dict(zip(df_prod['Offer_Clean'], df_prod[domain_col]))
            print(f"  ✓ Mapped {len(offer_to_domain)} offers to Domain")
        else:
            print(f"  ⚠ No 'Domain' column found")

        # ── CAGR mapper ──
        print(f"\n📋 Loading CAGR Mapper...")
        df_mapper = load_cagr_mapper()
        if df_mapper is None:
            print(f"  ✗ CAGR Mapper failed to load")
            return None
        print(f"  ✓ Loaded {len(df_mapper)} offers")
        
        # ✅ ADD SVP TO MAPPER BEFORE MERGING
        print(f"\n🔗 Adding SVP to CAGR Mapper...")
        df_mapper['SVP'] = df_mapper['Offer'].map(offer_to_svp).fillna('Unknown')
        df_mapper['Business_Line'] = df_mapper['Offer'].map(offer_to_business_line).fillna('---')
        df_mapper['Domain'] = df_mapper['Offer'].map(offer_to_domain).fillna('---')
        print(f"  ✓ Added SVP, Business_Line, Domain to mapper")

        # ── Market size ──
        print(f"\n📈 Loading Market Size...")
        df_ms = pd.read_excel(DATA_MARKET_SIZE,
                              sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
        print(f"  ✓ Loaded {len(df_ms)} rows")
        
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
        print(f"  ✓ Pivoted to {len(df_ms_pivot)} sub-segments")

        # ── Revenue ──
        print(f"\n💰 Loading Revenue data...")
        if not DATA_REVENUE or not os.path.exists(DATA_REVENUE):
            print(f"  ✗ Revenue file not found: {DATA_REVENUE}")
            return None
            
        df_rev = pd.read_excel(DATA_REVENUE)
        print(f"  ✓ Loaded {len(df_rev)} rows")
        print(f"  Columns: {list(df_rev.columns)}")
        
        if 'Region' not in df_rev.columns:
            df_rev['Region'] = 'FR'
            print(f"  ⚠ No Region column, defaulting to 'FR'")
        
        # ✅ Check for Product and Rep_Code columns (case-insensitive)
        required_cols = ["Offer","Period","_S_Revenue_actual","Region"]
        optional_cols = ["Product", "Rep_Code"]
        
        available_cols = required_cols.copy()
        for col in optional_cols:
            # Check if column exists (case-insensitive)
            matching_col = next((c for c in df_rev.columns if c.lower() == col.lower()), None)
            if matching_col:
                available_cols.append(matching_col)
                print(f"  ✓ Found optional column: {matching_col}")
            else:
                print(f"  ⚠ Missing optional column: {col}")
        
        df_rev = df_rev[available_cols]
        df_rev['Period'] = pd.to_datetime(df_rev['Period'], errors='coerce')
        df_rev = df_rev.dropna(subset=['Period'])
        print(f"  ✓ Cleaned {len(df_rev)} revenue rows")
        
        df_rev['Year'] = df_rev['Period'].dt.year
        
        if df_rev["_S_Revenue_actual"].dtype == object:
            df_rev["_S_Revenue_actual"] = (df_rev["_S_Revenue_actual"]
                .astype(str).str.replace(r"[,$]","",regex=True))
        df_rev["_S_Revenue_actual"] = pd.to_numeric(
            df_rev["_S_Revenue_actual"], errors='coerce').fillna(0)

        # ✅ Store original revenue data with Product and Rep_Code
        df_rev_detailed = df_rev.copy()
        
        # Default aggregation by Offer, Year, Region (for backward compatibility)
        df_rev_grp = df_rev.groupby(['Offer','Year','Region'], as_index=False)[
            '_S_Revenue_actual'].sum()
        df_rev_piv = df_rev_grp.pivot_table(
            index=['Offer','Region'], columns='Year',
            values='_S_Revenue_actual', fill_value=0).reset_index()
        df_rev_piv.columns = [
            f"Revenue_{c}" if isinstance(c, int) else c
            for c in df_rev_piv.columns]
        print(f"  ✓ Pivoted to {len(df_rev_piv)} offer-region combinations")

        # ── Merge (LEFT - keep all revenue data) ──
        print(f"\n🔗 Merging data...")
        df = df_rev_piv.merge(df_mapper, on='Offer', how='left')
        print(f"  ✓ After LEFT merge with CAGR Mapper: {len(df)} rows")
        
        # ✅ Track offers with missing CAGR mapper data
        missing_offers = df[df['Sub-segment 2'].isna()]['Offer'].tolist()
        print(f"  ⚠ {len(missing_offers)} offers missing CAGR mapper data")
        
        # ✅ Replace NaN values with defaults
        print(f"\n🧹 Cleaning NaN values...")
        
        # Categorical columns → '---'
        categorical_cols = ['SVP', 'Sub-segment 2', 'Market_Category', 'Segment', 
                           'Business_Line', 'Domain', 'Product', 'Rep_Code']
        for col in categorical_cols:
            if col in df.columns:
                before = df[col].isna().sum()
                df[col] = df[col].fillna('---')
                df[col] = df[col].replace('nan', '---')
                df[col] = df[col].astype(str).str.replace('nan', '---')
                after = df[col].isna().sum()
                if before > 0:
                    print(f"  ✓ {col}: cleaned {before} NaN values")
        
        # Numerical columns (Million EUR_*) → 0
        million_eur_cols = [c for c in df.columns if c.startswith('Million EUR_')]
        for col in million_eur_cols:
            before = df[col].isna().sum()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            if before > 0:
                print(f"  ✓ {col}: cleaned {before} NaN values")
        
        # Revenue columns → 0
        revenue_cols = [c for c in df.columns if c.startswith('Revenue_')]
        for col in revenue_cols:
            before = df[col].isna().sum()
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            if before > 0:
                print(f"  ✓ {col}: cleaned {before} NaN values")
        
        # ✅ Merge with market size (LEFT to keep all offers)
        print(f"\n🔗 Merging with Market Size...")
        df = df.merge(df_ms_pivot, on='Sub-segment 2', how='left')
        print(f"  ✓ After LEFT merge with Market Size: {len(df)} rows")
        
        # ✅ Fill remaining NaN in Million EUR columns
        for col in million_eur_cols:
            if col in df.columns:
                df[col] = df[col].fillna(0)
        
        # ✅ Store detailed revenue data and warnings
        df.attrs['detailed_revenue'] = df_rev_detailed
        df.attrs['missing_offers'] = missing_offers
        
        print(f"\n{'='*60}")
        print(f"✓ FINAL DATASET: {len(df)} rows, {len(df.columns)} columns")
        print(f"{'='*60}\n")
        
        if missing_offers:
            print(f"⚠ {len(missing_offers)} offers with missing CAGR mapper data:")
            for offer in missing_offers[:10]:
                print(f"    - {offer}")
            if len(missing_offers) > 10:
                print(f"    ... and {len(missing_offers) - 10} more")
        
        return df

    except Exception as e:
        print(f"\n{'='*60}")
        print(f"✗ ERROR in load_all_data:")
        print(f"{'='*60}")
        print(f"{str(e)}\n")
        import traceback
        traceback.print_exc()
        print(f"{'='*60}\n")
        return None


def resultat_data(df_all_data, display_start_year, display_end_year, cagr_start_year, cagr_end_year,
                  pred_start_year, selected_svps, region, use_custom_cagr='no', selected_columns=None,
                  subseg_filter='all', aggregation_level='offer'):
    try:
        print(f"\n{'='*60}")
        print(f"Display : {display_start_year} → {display_end_year}")
        print(f"CAGR    : {cagr_start_year} → {cagr_end_year}")
        print(f"Pred from: {pred_start_year}")
        print(f"SVPs    : {selected_svps}  Region: {region}")
        print(f"Custom CAGR: {use_custom_cagr}")
        print(f"Sub-segment filter: {subseg_filter}")
        print(f"Aggregation: {aggregation_level}")
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
        
        # ✅ Apply sub-segment filter
        if subseg_filter == 'with_subseg':
            df = df[df['Sub-segment 2'].notna() & (df['Sub-segment 2'] != '---') & (df['Sub-segment 2'] != 'nan')]
        elif subseg_filter == 'no_subseg':
            df = df[df['Sub-segment 2'].isna() | (df['Sub-segment 2'] == '---') | (df['Sub-segment 2'] == 'nan')]
        
        if region == 'France':
            df = df[df['Region'] == 'FR']
        elif region == 'International':
            df = df[df['Region'] == 'INT']
        if len(df) == 0:
            return {'error': 'No data found for selected criteria.'}

        # ✅ Re-aggregate revenue based on aggregation level
        if aggregation_level in ['product', 'rep_code'] and 'detailed_revenue' in df_all_data.attrs:
            df_rev_detailed = df_all_data.attrs['detailed_revenue']
            
            # Filter by selected offers
            selected_offers = df['Offer'].unique()
            df_rev_filtered = df_rev_detailed[df_rev_detailed['Offer'].isin(selected_offers)].copy()
            
            # Apply region filter
            if region == 'France':
                df_rev_filtered = df_rev_filtered[df_rev_filtered['Region'] == 'FR']
            elif region == 'International':
                df_rev_filtered = df_rev_filtered[df_rev_filtered['Region'] == 'INT']
            
            # Group by aggregation level
            group_cols = ['Offer', 'Year', 'Region']
            if aggregation_level == 'product' and 'Product' in df_rev_filtered.columns:
                group_cols.append('Product')
            elif aggregation_level == 'rep_code' and 'Rep_code' in df_rev_filtered.columns:
                group_cols.append('Rep_code')
            
            df_rev_grp = df_rev_filtered.groupby(group_cols, as_index=False)['_S_Revenue_actual'].sum()
            
            # Pivot
            index_cols = ['Offer', 'Region']
            if aggregation_level == 'product' and 'Product' in df_rev_grp.columns:
                index_cols.append('Product')
            elif aggregation_level == 'rep_code' and 'Rep_code' in df_rev_grp.columns:
                index_cols.append('Rep_code')
            
            df_rev_piv = df_rev_grp.pivot_table(
                index=index_cols, columns='Year',
                values='_S_Revenue_actual', fill_value=0).reset_index()
            df_rev_piv.columns = [
                f"Revenue_{c}" if isinstance(c, int) else c
                for c in df_rev_piv.columns]
            
            # Merge back with df (keep other columns)
            merge_on = ['Offer', 'Region']
            df = df.drop(columns=[c for c in df.columns if c.startswith('Revenue_')], errors='ignore')
            df = df.merge(df_rev_piv, on=merge_on, how='left')

        # Revenue years present in the data
        rev_cols  = sorted([c for c in df.columns if c.startswith('Revenue_')])
        rev_years = [int(c.split('_')[1]) for c in rev_cols]
        last_rev_year = rev_years[-1] if rev_years else display_start_year

        all_display_years = list(range(display_start_year, display_end_year + 1))

        # ── CAGR column names ──────────────────────────────────────────────
        mif_cagr_col_name = f'MIF_CAGR_{cagr_start_year}/{cagr_end_year}'
        cagr_col_name = None
        if use_custom_cagr == 'yes':
            cagr_col_name = f'CAGR'

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
            sub2            = row.get('Sub-segment 2', '---')
            business_line   = row.get('Business_Line', '---')
            domain          = row.get('Domain', '---')
            
            # ✅ Get Product and Rep_code if aggregation level requires it
            product  = row.get('Product', '---') if aggregation_level == 'product' else None
            rep_code = row.get('Rep_code', '---') if aggregation_level == 'rep_code' else None

            # ── Calculate both CAGRs ───────────────────────────────────────
            cagr        = None
            mif_cagr    = None
            cagr_source = 'market'

            if sub2 is None or (isinstance(sub2, float) and np.isnan(sub2)):
                sub2 = '---'

            sc = f"Million EUR_{cagr_start_year}"
            ec = f"Million EUR_{cagr_end_year}"

            if sc in df.columns and ec in df.columns:
                ms_s     = row[sc] if not pd.isna(row[sc]) else 0
                ms_e     = row[ec] if not pd.isna(row[ec]) else 0
                mif_cagr = calculate_cagr(ms_s, ms_e, cagr_end_year - cagr_start_year)
            else:
                mif_cagr = 0

            # Check for custom CAGR
            if use_custom_cagr == 'yes' and custom_cagr_dict:
                if offer in custom_cagr_dict:
                    cagr        = custom_cagr_dict[offer]
                    cagr_source = 'custom'
                    custom_cagr_applied_count += 1

            # If no custom CAGR, fall back to MIF CAGR
            if cagr is None:
                cagr        = mif_cagr
                cagr_source = 'market'

            cagr_values.append(cagr)

            # ── Actual revenues from data ──────────────────────────────────
            actual = {}
            for y in rev_years:
                col = f"Revenue_{y}"
                if col in row.index:
                    v = row[col]
                    actual[y] = float(v) if not pd.isna(v) else 0.0

            # Seed for projection
            seed_y = pred_start_year
            seed_v = actual.get(seed_y, None)

            if seed_v is None:
                seed_v = actual.get(last_rev_year, 0.0)
                seed_y = last_rev_year

            # ── Build result row ───────────────────────────────────────────
            result_row = OrderedDict()
            result_row['Offer']           = offer
            result_row['SVP']             = svp
            result_row['Sub-segment 2']   = sub2
            result_row['Market_Category'] = market_category
            result_row['Business_Line']   = business_line
            result_row['Domain']          = domain
            
            # ✅ Add Product and Rep_code if applicable
            if product is not None:
                result_row['Product'] = product
            if rep_code is not None:
                result_row['Rep_code'] = rep_code

            # Both CAGR columns
            if use_custom_cagr == 'yes' and cagr_source == 'custom':
                result_row[cagr_col_name] = f"{cagr*100:.2f}% (custom)"

            result_row[mif_cagr_col_name] = f"{mif_cagr*100:.2f}%"

            # ── Year columns ───────────────────────────────────────────────
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

        year_cols = []
        for y in all_display_years:
            if any(f'Revenue_{y}' in r for r in results_list):
                year_cols.append(f'Revenue_{y}')
            if any(f'Real_{y}' in r for r in results_list):
                year_cols.append(f'Real_{y}')
            if any(f'Predicted_{y}' in r for r in results_list):
                year_cols.append(f'Predicted_{y}')

        # ── Filter columns based on user selection ─────────────────────────

        if selected_columns is None:
            selected_columns = ['Offer', 'SVP', 'Sub-segment 2', 'Market_Category', 'CAGR', 'MIF_CAGR']

        filtered_column_order = []

        # Fixed columns
        fixed_cols_to_check = ['Offer', 'SVP', 'Sub-segment 2', 'Market_Category', 
                               'Business_Line', 'Domain', 'Product', 'Rep_code']
        for col in fixed_cols_to_check:
            if col == 'Offer' or col in selected_columns:
                filtered_column_order.append(col)

        # CAGR columns
        if 'CAGR' in selected_columns and cagr_col_name != None:
            filtered_column_order.append(cagr_col_name)
        if 'MIF_CAGR' in selected_columns:
            filtered_column_order.append(mif_cagr_col_name)

        # Year columns always shown
        filtered_column_order.extend(year_cols)

        column_order = filtered_column_order

        # ✅ Filter table rows to only include selected columns
        results_list_filtered = []
        for r in results_list:
            filtered_row = OrderedDict()
            for col in column_order:
                if col in r:
                    filtered_row[col] = r[col]
            results_list_filtered.append(filtered_row)
        results_list = results_list_filtered

        # ── TOTAL row ──────────────────────────────────────────────────────

        total_row = OrderedDict()
        total_row['Offer']           = 'TOTAL'
        total_row['SVP']             = ''
        total_row['Sub-segment 2']   = ''
        total_row['Market_Category'] = ''
        total_row['Business_Line']   = ''
        total_row['Domain']          = ''
        if 'Product' in column_order:
            total_row['Product'] = ''
        if 'Rep_code' in column_order:
            total_row['Rep_code'] = ''
        if cagr_col_name != None:
            total_row[cagr_col_name] = ''
        total_row[mif_cagr_col_name] = ''

        for col in year_cols:
            y = int(col.split('_')[-1])
            if col.startswith('Revenue_'):
                total_row[col] = f"{tot_actual[y]:,.2f}"
            elif col.startswith('Real_'):
                total_row[col] = f"{tot_real_post[y]:,.2f}"
            elif col.startswith('Predicted_'):
                total_row[col] = f"{tot_predicted[y]:,.2f}"

        # Filter total row too
        filtered_total = OrderedDict()
        for col in column_order:
            if col in total_row:
                filtered_total[col] = total_row[col]
        results_list.append(filtered_total)

        # ── KPIs ───────────────────────────────────────────────────────────

        avg_cagr      = np.mean(cagr_values) if cagr_values else 0
        total_row_data = results_list[-1]

        rev_start_col = f'Revenue_{display_start_year}'
        rev_end_col   = f'Real_{pred_start_year}'
        Actual_cagr   = 0

        if rev_start_col in total_row_data and rev_end_col in total_row_data:
            rev_start_str = total_row_data[rev_start_col]
            rev_end_str   = total_row_data[rev_end_col]
            if rev_start_str and rev_start_str != '':
                rev_start   = float(rev_start_str.replace(',', ''))
                rev_end     = float(rev_end_str.replace(',', ''))
                Actual_cagr = calculate_cagr(rev_start, rev_end, pred_start_year - display_start_year)

        # ── Chart arrays ───────────────────────────────────────────────────

        chart_actual    = []
        chart_real_post = []
        chart_predicted = []

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

                # ✅ Ensure no NaN values in response
        def clean_for_json(obj):
            """Recursively replace NaN with '---' or 0"""
            if isinstance(obj, dict):
                return {k: clean_for_json(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [clean_for_json(item) for item in obj]
            elif isinstance(obj, float):
                if np.isnan(obj):
                    return 0
                return obj
            elif obj is None:
                return '---'
            return obj

        result_dict = {
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
        
        return clean_for_json(result_dict)


    except Exception as e:
        import traceback; traceback.print_exc()
        return {'error': str(e)}
