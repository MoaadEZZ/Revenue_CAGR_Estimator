import pandas as pd
import numpy as np
import os
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict

from mapping import (
    DATA_CAGR_SVP,
    DATA_MARKET_SIZE,
    DATA_ALL_PRODUCTS,
    build_category_mapper,
)


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

# Revenue is now uploaded as two regional files (France / International)
# and combined into a single data.xlsx with a Region column — the format
# the rest of the pipeline expects.
DATA_REVENUE_FR   = os.path.join(DATA_DIR, "data_france.xlsx")
DATA_REVENUE_INTL = os.path.join(DATA_DIR, "data_international.xlsx")
DATA_REVENUE      = os.path.join(DATA_DIR, "data.xlsx")


def combine_revenue_files():
    """Combine the France + International revenue uploads into the single
    data.xlsx (with a Region column) that the rest of the pipeline expects.
    Safe to call any time either regional file changes."""
    frames = []
    if os.path.exists(DATA_REVENUE_FR):
        df_fr = pd.read_excel(DATA_REVENUE_FR)
        df_fr['Region'] = 'FR'
        frames.append(df_fr)
    if os.path.exists(DATA_REVENUE_INTL):
        df_intl = pd.read_excel(DATA_REVENUE_INTL)
        df_intl['Region'] = 'INT'
        frames.append(df_intl)
    if not frames:
        print("⚠ No regional revenue files found — data.xlsx not (re)generated")
        return None
    df_combined = pd.concat(frames, ignore_index=True)
    df_combined.to_excel(DATA_REVENUE, index=False)
    print(f"✓ Combined revenue files → {DATA_REVENUE} ({len(df_combined)} rows)")
    return df_combined


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
    try:
        filepath = os.path.join(DATA_DIR, 'custom_cagr.xlsx')
        if not os.path.exists(filepath):
            print("No custom CAGR file found")
            return None
        df = pd.read_excel(filepath)
        df.columns = df.columns.str.strip().str.title()
        offer_col = None
        cagr_col  = None
        for col in df.columns:
            if col.lower() == 'offer': offer_col = col
            elif col.lower() == 'cagr': cagr_col = col
        if not offer_col or not cagr_col:
            print(f"⚠ Custom CAGR: Missing required columns. Found: {list(df.columns)}")
            return None
        custom_cagr_dict = {}
        skipped_count = 0
        for idx, row in df.iterrows():
            offer      = row[offer_col]
            cagr_value = row[cagr_col]
            if pd.isna(offer) or offer == '':
                skipped_count += 1; continue
            if pd.isna(cagr_value) or cagr_value == '':
                skipped_count += 1; continue
            try:
                cagr_float = float(cagr_value)
                custom_cagr_dict[str(offer).strip()] = cagr_float
            except ValueError:
                skipped_count += 1; continue
        if custom_cagr_dict:
            print(f"✓ Custom CAGR loaded: {len(custom_cagr_dict)} offers")
            return custom_cagr_dict
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


def load_category_mapper(force_rebuild=False):
    """Rep_Code-level category mapper: Rep_Code / catégorie / ID, built by
    combining the hierarchical Rep_Code→catégorie matching (CAGR SVP + All
    Products) with the semantic catégorie→ID matching (Market Size). See
    mapping.py. Returned columns are renamed to match this pipeline's
    conventions: Market_Category (was catégorie), Rep_Code, ID."""
    try:
        mapper = build_category_mapper(force_rebuild=force_rebuild)
        if mapper is None or mapper.empty:
            return None
        mapper = mapper.rename(columns={'catégorie': 'Market_Category'})
        mapper['Market_Category'] = mapper['Market_Category'].fillna('').astype(str).str.rstrip()
        mapper['ID']              = mapper['ID'].fillna('').astype(str).str.rstrip()
        for seg_col in ['Segment', 'Sub-segment 1', 'Sub-segment 2', 'Sub-segment 3']:
            if seg_col in mapper.columns:
                mapper[seg_col] = mapper[seg_col].fillna('---').astype(str).str.rstrip()
        print(f"✓ Category Mapper: {len(mapper)} Rep_Codes "
              f"({(mapper['Market_Category'] != '').sum()} with catégorie, "
              f"{(mapper['ID'] != '').sum()} with ID)")
        return mapper
    except Exception as e:
        print(f"Error loading category mapper: {e}")
        return None


def load_all_data():
    try:
        print("\n" + "="*60)
        print("LOADING ALL DATA")
        print("="*60)

        # ── SVP + CVP mapping ──
        print(f"\n📊 Loading SVP/CVP mapping from All Products...")
        df_prod = pd.read_excel(DATA_ALL_PRODUCTS)
        print(f"  ✓ Loaded {len(df_prod)} rows")
        print(f"  Available columns: {list(df_prod.columns)}")

        df_prod['Offer'] = df_prod['Offer'].astype(str).str.strip()
        if df_prod['Offer'].str.contains('.', regex=False).any():
            df_prod['Offer_Clean'] = df_prod['Offer'].str.split('.', n=1, expand=True)[1].str.strip()
        else:
            df_prod['Offer_Clean'] = df_prod['Offer']

        # Normalize the Rep_Code column name (used to roll the new
        # Rep_Code-level category mapper up to Offer level below).
        rep_code_col = next((c for c in df_prod.columns if c.strip().lower() == 'rep_code'), None)
        if rep_code_col and rep_code_col != 'Rep_Code':
            df_prod = df_prod.rename(columns={rep_code_col: 'Rep_Code'})

        # SVP
        svp_col = None
        for col in df_prod.columns:
            if col.lower() in ['strategic', 'svp', 'strategic value proposition', 'svp category']:
                svp_col = col; break
        offer_to_svp = dict(zip(df_prod['Offer_Clean'], df_prod[svp_col])) if svp_col else {}
        print(f"  SVP column: {svp_col} → {len(offer_to_svp)} offers mapped")

        # CVP
        cvp_col = None
        for col in df_prod.columns:
            if col.lower() in ['cvp', 'customer value proposition', 'cvp category']:
                cvp_col = col; break
        offer_to_cvp = dict(zip(df_prod['Offer_Clean'], df_prod[cvp_col])) if cvp_col else {}
        print(f"  CVP column: {cvp_col} → {len(offer_to_cvp)} offers mapped")

        # Business Line
        business_line_col = None
        for col in df_prod.columns:
            if col.lower() in ['business line', 'business_line', 'businessline']:
                business_line_col = col; break
        offer_to_business_line = dict(zip(df_prod['Offer_Clean'], df_prod[business_line_col])) \
            if business_line_col else {}

        # Domain
        domain_col = None
        for col in df_prod.columns:
            if col.lower() in ['domain', 'sub_domain', 'subdomain']:
                domain_col = col; break
        offer_to_domain = dict(zip(df_prod['Offer_Clean'], df_prod[domain_col])) \
            if domain_col else {}

        # ── Uniqueness check: an Offer must map to exactly 1 Domain and ──
        # ── exactly 1 Business_Line (sub-business line) to be safely      ──
        # ── split into Rep-Codes for the "By Rep_code" aggregation level. ──
        def _offers_with_multiple_values(col_name):
            if not col_name:
                return []
            tmp = df_prod[['Offer_Clean', col_name]].copy()
            tmp[col_name] = tmp[col_name].astype(str).str.strip()
            tmp = tmp[~tmp[col_name].isin(['', 'nan', 'None', '---'])]
            counts = tmp.groupby('Offer_Clean')[col_name].nunique()
            return sorted(counts[counts > 1].index.tolist())

        multi_domain_offers = _offers_with_multiple_values(domain_col)
        multi_bl_offers     = _offers_with_multiple_values(business_line_col)
        if multi_domain_offers:
            print(f"  ⚠ {len(multi_domain_offers)} offers have more than one Domain")
        if multi_bl_offers:
            print(f"  ⚠ {len(multi_bl_offers)} offers have more than one Business_Line")

        # ── Category mapper (Rep_Code → catégorie → ID) ──
        # Built from CAGR SVP (hierarchical match, test1.py idea) + Market
        # Size (semantic match, test2.py idea). See mapping.py.
        print(f"\n📋 Loading Category Mapper...")
        df_repcode_mapper = load_category_mapper()
        if df_repcode_mapper is None:
            print(f"  ✗ Category mapper failed to load"); return None
        print(f"  ✓ Loaded {len(df_repcode_mapper)} Rep_Codes")

        # Roll the Rep_Code-level mapper up to Offer level (first non-empty
        # catégorie/ID per Offer wins) so the existing Offer-level
        # pivot/merge machinery below keeps working. Rep_Code-level detail
        # itself is only needed for revenue splitting, which already
        # happens separately via detailed_revenue in resultat_data().
        if 'Rep_Code' in df_prod.columns:
            df_repcode_mapper = df_repcode_mapper.merge(
                df_prod[['Rep_Code', 'Offer_Clean']].drop_duplicates(subset=['Rep_Code']),
                on='Rep_Code', how='left')
        else:
            df_repcode_mapper['Offer_Clean'] = None
            print("  ⚠ All products requested.xlsx has no Rep_Code column — "
                  "cannot roll the category mapper up to Offer level")

        def _first_nonempty(s):
            vals = [v for v in s if v not in (None, '', 'nan', '---')]
            return vals[0] if vals else ''

        # Roll up every column the mapper produced (catégorie/ID plus the
        # Segment / Sub-segment 1-3 hierarchy), not just Market_Category
        # and ID — previously the .agg() only named those two, which
        # silently dropped the segment columns even when mapping.py had
        # already computed them.
        segment_cols = [c for c in ['Segment', 'Sub-segment 1', 'Sub-segment 2', 'Sub-segment 3']
                        if c in df_repcode_mapper.columns]
        agg_spec = {'Market_Category': ('Market_Category', _first_nonempty),
                    'ID': ('ID', _first_nonempty)}
        for col in segment_cols:
            agg_spec[col] = (col, _first_nonempty)

        df_mapper = (
            df_repcode_mapper.dropna(subset=['Offer_Clean'])
            .groupby('Offer_Clean', as_index=False)
            .agg(**agg_spec)
            .rename(columns={'Offer_Clean': 'Offer'})
        )
        print(f"  ✓ Rolled up to {len(df_mapper)} offers "
              f"(segment columns carried through: {segment_cols or 'none'})")

        df_mapper['SVP']           = df_mapper['Offer'].map(offer_to_svp).fillna('Unknown')
        df_mapper['CVP']           = df_mapper['Offer'].map(offer_to_cvp).fillna('Unknown')
        df_mapper['Business_Line'] = df_mapper['Offer'].map(offer_to_business_line).fillna('---')
        df_mapper['Domain']        = df_mapper['Offer'].map(offer_to_domain).fillna('---')
        print(f"  ✓ Added SVP, CVP, Business_Line, Domain to mapper")

        # ── Market size (pivoted on ID, matched to catégorie semantically) ──
        print(f"\n📈 Loading Market Size...")
        df_ms = pd.read_excel(DATA_MARKET_SIZE,
                              sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
        id_col = next((c for c in df_ms.columns if str(c).strip().lower() == 'id'), None)
        if id_col is None:
            print(f"  ✗ Market Size file has no 'ID' column"); return None
        df_ms = df_ms.rename(columns={id_col: 'ID'})
        df_ms = df_ms[["ID","Year","Million EUR"]].copy()
        df_ms["ID"] = df_ms["ID"].astype(str).str.rstrip()
        df_ms = df_ms.dropna(subset=["ID","Year","Million EUR"])
        df_ms = df_ms.groupby(["ID","Year"], as_index=False).agg({"Million EUR":"sum"})
        df_ms['Year'] = df_ms['Year'].astype(int)
        df_ms_pivot = df_ms.pivot_table(
            index=["ID"], columns="Year",
            values="Million EUR", fill_value=0).reset_index()
        df_ms_pivot.columns = [
            f"Million EUR_{c}" if isinstance(c, int) else c
            for c in df_ms_pivot.columns]
        print(f"  ✓ Pivoted to {len(df_ms_pivot)} market IDs")

        # ── Revenue ──
        print(f"\n💰 Loading Revenue data...")
        if not DATA_REVENUE or not os.path.exists(DATA_REVENUE):
            print(f"  ✗ Revenue file not found"); return None
        df_rev = pd.read_excel(DATA_REVENUE)
        if 'Region' not in df_rev.columns:
            df_rev['Region'] = 'FR'
        required_cols = ["Offer","Period","_S_Revenue_actual","Region"]
        optional_cols = ["Product", "Rep_Code"]
        available_cols = required_cols.copy()
        for col in optional_cols:
            matching_col = next((c for c in df_rev.columns if c.lower() == col.lower()), None)
            if matching_col:
                available_cols.append(matching_col)
        df_rev = df_rev[available_cols]
        df_rev['Period'] = pd.to_datetime(df_rev['Period'], errors='coerce')
        df_rev = df_rev.dropna(subset=['Period'])
        df_rev['Year'] = df_rev['Period'].dt.year
        if df_rev["_S_Revenue_actual"].dtype == object:
            df_rev["_S_Revenue_actual"] = (df_rev["_S_Revenue_actual"]
                .astype(str).str.replace(r"[,$]","",regex=True))
        df_rev["_S_Revenue_actual"] = pd.to_numeric(
            df_rev["_S_Revenue_actual"], errors='coerce').fillna(0)
        df_rev_detailed = df_rev.copy()
        df_rev_grp = df_rev.groupby(['Offer','Year','Region'], as_index=False)[
            '_S_Revenue_actual'].sum()
        df_rev_piv = df_rev_grp.pivot_table(
            index=['Offer','Region'], columns='Year',
            values='_S_Revenue_actual', fill_value=0).reset_index()
        df_rev_piv.columns = [
            f"Revenue_{c}" if isinstance(c, int) else c
            for c in df_rev_piv.columns]
        print(f"  ✓ Pivoted to {len(df_rev_piv)} offer-region combinations")

        # ── Merge ──
        print(f"\n🔗 Merging data...")
        df = df_rev_piv.merge(df_mapper, on='Offer', how='left')
        missing_offers = df[df['Market_Category'].isna() | (df['Market_Category'] == '')]['Offer'].tolist()

        # Clean categoricals
        categorical_cols = ['SVP', 'CVP', 'Market_Category', 'ID',
                            'Business_Line', 'Domain', 'Product', 'Rep_Code',
                            'Segment', 'Sub-segment 1', 'Sub-segment 2', 'Sub-segment 3']
        for col in categorical_cols:
            if col in df.columns:
                df[col] = df[col].fillna('---').replace('nan','---')
                df[col] = df[col].astype(str).str.replace('nan','---')

        million_eur_cols = [c for c in df.columns if c.startswith('Million EUR_')]
        for col in million_eur_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        revenue_cols = [c for c in df.columns if c.startswith('Revenue_')]
        for col in revenue_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df = df.merge(df_ms_pivot, on='ID', how='left')
        for col in million_eur_cols:
            if col in df.columns:
                df[col] = df[col].fillna(0)

        df.attrs['detailed_revenue']     = df_rev_detailed
        df.attrs['missing_offers']       = missing_offers
        df.attrs['multi_domain_offers']  = multi_domain_offers
        df.attrs['multi_bl_offers']      = multi_bl_offers

        # Store CVP options for the frontend
        # ── CVP OPTIONS: Two sections (with offers / without offers) ──
        if cvp_col:
            # All CVPs from products file
            all_cvp_values = df_prod[cvp_col].dropna().astype(str).str.strip()
            all_cvps = sorted([
                v for v in all_cvp_values.unique().tolist()
                if v not in ('---', 'Unknown', 'nan', '', 'NaN')
            ])
            
            # CVPs that have offers with revenue data (merged successfully)
            cvps_with_revenue = set()
            if 'CVP' in df.columns:
                cvps_with_revenue = set([
                    v for v in df['CVP'].dropna().astype(str).str.strip().unique().tolist()
                    if v not in ('---', 'Unknown', 'nan', '', 'NaN')
                ])
            
            # Split into two sections
            cvps_with_offers = sorted([c for c in all_cvps if c in cvps_with_revenue])
            cvps_without_offers = sorted([c for c in all_cvps if c not in cvps_with_revenue])
            
            # Store as dict with sections
            df.attrs['cvp_options'] = {
                'with_offers': cvps_with_offers,
                'without_offers': cvps_without_offers
            }
            
            print(f"\n  📊 CVP Sections:")
            print(f"     With offers: {len(cvps_with_offers)} CVPs")
            for cvp in cvps_with_offers:
                print(f"       ✓ {cvp}")
            print(f"     Without offers: {len(cvps_without_offers)} CVPs")
            for cvp in cvps_without_offers:
                print(f"       ✗ {cvp}")
        else:
            df.attrs['cvp_options'] = {'with_offers': [], 'without_offers': []}


        print(f"\n{'='*60}")
        print(f"✓ FINAL DATASET: {len(df)} rows, {len(df.columns)} columns")
        print(f"{'='*60}\n")
        return df

    except Exception as e:
        print(f"\n✗ ERROR in load_all_data: {e}")
        import traceback; traceback.print_exc()
        return None


def resultat_data(df_all_data, display_start_year, display_end_year,
                  cagr_start_year, cagr_end_year, pred_start_year,
                  selected_svps, region, use_custom_cagr='no',
                  selected_columns=None, subseg_filter='all',
                  aggregation_level='offer', selected_cvps=None):
    try:
        print(f"\n{'='*60}")
        print(f"Display : {display_start_year} → {display_end_year}")
        print(f"CAGR    : {cagr_start_year} → {cagr_end_year}")
        print(f"Pred from: {pred_start_year}")
        print(f"SVPs    : {selected_svps}  CVPs: {selected_cvps}  Region: {region}")
        print(f"Custom CAGR: {use_custom_cagr}")
        print(f"{'='*60}")

        # Load custom CAGR if requested
        custom_cagr_dict = None
        if use_custom_cagr in ('yes', 'both'):
            custom_cagr_dict = load_custom_cagr()
            if custom_cagr_dict:
                print(f"\n📊 Applying custom CAGR (BL) to {len(custom_cagr_dict)} offers\n")

        if not selected_svps and not selected_cvps:
            return {'error': 'Please select at least one SVP or CVP category'}
        if df_all_data is None:
            return {'error': 'Data not loaded.'}

        # ── Filter by SVP OR CVP ──
        svp_mask = df_all_data['SVP'].isin(selected_svps) \
            if selected_svps else pd.Series(False, index=df_all_data.index)
        cvp_mask = df_all_data['CVP'].isin(selected_cvps) \
            if selected_cvps else pd.Series(False, index=df_all_data.index)
        df = df_all_data[svp_mask | cvp_mask].copy()

        # Category-assignment filter: filters on whether a catégorie/ID
        # was successfully assigned (Segment/Sub-segment 1-3 are now
        # carried through separately as display-only columns).
        if subseg_filter == 'with_subseg':
            df = df[df['Market_Category'].notna() &
                    (df['Market_Category'] != '---') &
                    (df['Market_Category'] != 'nan') &
                    (df['Market_Category'] != '')]
        elif subseg_filter == 'no_subseg':
            df = df[df['Market_Category'].isna() |
                    (df['Market_Category'] == '---') |
                    (df['Market_Category'] == 'nan') |
                    (df['Market_Category'] == '')]

        if region == 'France':
            df = df[df['Region'] == 'FR']
        elif region == 'International':
            df = df[df['Region'] == 'INT']

        if len(df) == 0:
            return {'error': 'No data found for selected criteria.'}

        # ── Rep-Code precondition: each Offer must map to exactly 1 Domain ──
        # ── and exactly 1 Sub-Business Line before it can be safely split  ──
        # ── into Rep-Codes (categorical columns represent a hierarchy).    ──
        if aggregation_level == 'rep_code':
            multi_domain = set(df_all_data.attrs.get('multi_domain_offers', []))
            multi_bl     = set(df_all_data.attrs.get('multi_bl_offers', []))
            offending    = sorted((multi_domain | multi_bl) & set(df['Offer'].unique()))
            if offending:
                preview = ', '.join(offending[:10]) + ('…' if len(offending) > 10 else '')
                return {'error': (
                    f"Cannot aggregate by Rep-Code: {len(offending)} offer(s) have more than "
                    f"one Domain or Sub-Business Line, so they cannot be safely split by "
                    f"Rep-Code: {preview}"
                )}

        # Re-aggregate if needed
        if aggregation_level in ['product', 'rep_code'] and \
                'detailed_revenue' in df_all_data.attrs:
            df_rev_detailed = df_all_data.attrs['detailed_revenue']
            selected_offers = df['Offer'].unique()
            df_rev_filtered = df_rev_detailed[
                df_rev_detailed['Offer'].isin(selected_offers)].copy()
            if region == 'France':
                df_rev_filtered = df_rev_filtered[df_rev_filtered['Region'] == 'FR']
            elif region == 'International':
                df_rev_filtered = df_rev_filtered[df_rev_filtered['Region'] == 'INT']
            group_cols = ['Offer', 'Year', 'Region']
            if aggregation_level == 'product' and 'Product' in df_rev_filtered.columns:
                group_cols.append('Product')
            elif aggregation_level == 'rep_code' and 'Rep_code' in df_rev_filtered.columns:
                group_cols.append('Rep_code')
            df_rev_grp = df_rev_filtered.groupby(
                group_cols, as_index=False)['_S_Revenue_actual'].sum()
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
            df = df.drop(columns=[c for c in df.columns if c.startswith('Revenue_')],
                         errors='ignore')
            df = df.merge(df_rev_piv, on=['Offer', 'Region'], how='left')

        rev_cols      = sorted([c for c in df.columns if c.startswith('Revenue_')])
        rev_years     = [int(c.split('_')[1]) for c in rev_cols]
        last_rev_year = rev_years[-1] if rev_years else display_start_year
        all_display_years = list(range(display_start_year, display_end_year + 1))

        # ── CAGR column names ──
        mif_cagr_col_name = f'MIF_CAGR_{cagr_start_year}/{cagr_end_year}'
        bl_cagr_col_name  = 'BL_CAGR'   # shown when use_custom_cagr in (yes, both)

        results_list = []
        cagr_values  = []
        _cagr_seen_offers = set()

        tot_actual    = {y: 0.0 for y in all_display_years}
        tot_real_post = {y: 0.0 for y in all_display_years}
        # For 'both' mode we track two prediction totals
        tot_predicted_mif = {y: 0.0 for y in all_display_years}
        tot_predicted_bl  = {y: 0.0 for y in all_display_years}
        has_real_post     = {y: False for y in all_display_years}

        for _, row in df.iterrows():
            offer           = row['Offer']
            market_category = row.get('Market_Category', '---')
            svp             = row.get('SVP', '---')
            business_line   = row.get('Business_Line', '---')
            domain          = row.get('Domain', '---')
            product  = row.get('Product', '---')  if aggregation_level == 'product'  else None
            rep_code = row.get('Rep_code', '---') if aggregation_level == 'rep_code' else None

            sc = f"Million EUR_{cagr_start_year}"
            ec = f"Million EUR_{cagr_end_year}"
            if sc in df.columns and ec in df.columns:
                ms_s     = row[sc] if not pd.isna(row[sc]) else 0
                ms_e     = row[ec] if not pd.isna(row[ec]) else 0
                mif_cagr = calculate_cagr(ms_s, ms_e, cagr_end_year - cagr_start_year)
            else:
                mif_cagr = 0

            # BL (custom) CAGR
            if use_custom_cagr in ('yes', 'both') and custom_cagr_dict:
                bl_cagr = custom_cagr_dict.get(offer, mif_cagr)
            else:
                bl_cagr = mif_cagr

            # For 'yes' mode the active CAGR is BL; for 'no' it's MIF; for 'both' we keep both
            if use_custom_cagr == 'yes':
                active_cagr = bl_cagr
            else:
                active_cagr = mif_cagr

            if aggregation_level == 'rep_code':
                if offer not in _cagr_seen_offers:
                    _cagr_seen_offers.add(offer)
                    cagr_values.append(active_cagr)
            else:
                cagr_values.append(active_cagr)

            # Actual revenues
            actual = {}
            for y in rev_years:
                col = f"Revenue_{y}"
                if col in row.index:
                    v = row[col]
                    actual[y] = float(v) if not pd.isna(v) else 0.0

            seed_y = pred_start_year
            seed_v = actual.get(seed_y, None)
            if seed_v is None:
                seed_v = actual.get(last_rev_year, 0.0)
                seed_y = last_rev_year

            # ── Build result row ──
            result_row = OrderedDict()

            result_row['Offer']           = offer
            result_row['SVP']             = svp
            result_row['CVP']             = row.get('CVP', '---')  # ✅ ADD THIS LINE
            result_row['Market_Category'] = market_category
            result_row['Business_Line']   = business_line
            result_row['Domain']          = domain
            if product  is not None: result_row['Product']  = product
            if rep_code is not None: result_row['Rep_code'] = rep_code


            # CAGR columns
            if use_custom_cagr in ('yes', 'both'):
                result_row[bl_cagr_col_name] = f"{bl_cagr*100:.2f}%"
            result_row[mif_cagr_col_name] = f"{mif_cagr*100:.2f}%"

            # ── Year columns ──
            for y in all_display_years:
                act_val = actual.get(y, None)

                if y < pred_start_year:
                    if act_val is not None:
                        tot_predicted_mif[y] += act_val
                        tot_predicted_bl[y]  += act_val
                    result_row[f'Revenue_{y}'] = f"{act_val:,.2f}" if act_val is not None else "0.00"
                    if act_val is not None:
                        tot_actual[y] += act_val
                else:
                    # Prediction years
                    steps_mif = y - seed_y
                    pred_mif  = seed_v * ((1 + mif_cagr) ** steps_mif)
                    pred_bl   = seed_v * ((1 + bl_cagr)  ** steps_mif)

                    if y == pred_start_year:
                        pred_mif = act_val if act_val is not None else pred_mif
                        pred_bl  = pred_mif  # same seed

                    if act_val is not None and y == pred_start_year:
                        result_row[f'Real_{y}'] = f"{act_val:,.2f}"
                        tot_actual[y]    += act_val
                        tot_real_post[y] += act_val
                        has_real_post[y]  = True

                    if use_custom_cagr == 'both':
                        result_row[f'Predicted_MIF_{y}'] = f"{pred_mif:,.2f}"
                        result_row[f'Predicted_BL_{y}']  = f"{pred_bl:,.2f}"
                        tot_predicted_mif[y] += pred_mif
                        tot_predicted_bl[y]  += pred_bl
                    else:
                        pred_val = pred_bl if use_custom_cagr == 'yes' else pred_mif
                        result_row[f'Predicted_{y}'] = f"{pred_val:,.2f}"
                        tot_predicted_mif[y] += pred_val
                        tot_predicted_bl[y]  += pred_val

                    if y > pred_start_year:
                        if act_val is not None:
                            # Also record this per-offer row's actual value under
                            # Real_{y} (not just in the tot_real_post aggregate).
                            # Without this, the per-offer "real/historical" curve
                            # and kpiSelectedOfferCagrActual (which reads
                            # Real_{lastYear} per selected row) have no data to
                            # work with when pred_start_year is set to an old year
                            # and one/more offers are selected instead of viewing
                            # the aggregate "Total Revenue by Year" chart.
                            result_row[f'Real_{y}'] = f"{act_val:,.2f}"
                            tot_real_post[y] += act_val
                            has_real_post[y]  = True

            results_list.append(result_row)

        # ── Rep-Code → Offer roll-up ──────────────────────────────────────
        # CAGR is computed and applied at Rep-Code granularity (each
        # Rep-Code keeps its own seed-year revenue while inheriting its
        # parent Offer's Sub-segment 2 / CAGR, since Rep-Codes belong to a
        # single Offer). For visualization we then sum the Rep-Codes back
        # up, grouped by Offer, so the table/chart show one row per Offer.
        if aggregation_level == 'rep_code' and results_list:
            year_col_prefixes = ('Revenue_', 'Real_', 'Predicted_MIF_', 'Predicted_BL_', 'Predicted_')
            grouped = OrderedDict()
            for r in results_list:
                offer = r['Offer']
                if offer not in grouped:
                    agg = OrderedDict()
                    for k, v in r.items():
                        if k == 'Rep_code':
                            continue
                        elif any(k.startswith(p) for p in year_col_prefixes):
                            agg[k] = 0.0
                        else:
                            agg[k] = v  # categorical / CAGR% cols identical across an Offer's Rep-Codes
                    grouped[offer] = agg
                for k, v in r.items():
                    if any(k.startswith(p) for p in year_col_prefixes):
                        grouped[offer][k] += float(str(v).replace(',', '')) if v not in (None, '') else 0.0

            # Re-format the summed year columns back to the "x,xxx.xx" string style
            for offer, agg in grouped.items():
                for k in list(agg.keys()):
                    if any(k.startswith(p) for p in year_col_prefixes):
                        agg[k] = f"{agg[k]:,.2f}"

            results_list = list(grouped.values())

        results_list = sorted(results_list, key=lambda x: x['Offer'])

        # ── Column order ──
        year_cols = []
        for y in all_display_years:
            if any(f'Revenue_{y}' in r for r in results_list):
                year_cols.append(f'Revenue_{y}')
            if any(f'Real_{y}' in r for r in results_list):
                year_cols.append(f'Real_{y}')
            if use_custom_cagr == 'both':
                if any(f'Predicted_MIF_{y}' in r for r in results_list):
                    year_cols.append(f'Predicted_MIF_{y}')
                if any(f'Predicted_BL_{y}' in r for r in results_list):
                    year_cols.append(f'Predicted_BL_{y}')
            else:
                if any(f'Predicted_{y}' in r for r in results_list):
                    year_cols.append(f'Predicted_{y}')

        if selected_columns is None:
            selected_columns = ['Offer','SVP','CVP','Market_Category',
                                'BL_CAGR','MIF_CAGR']
        print(selected_columns)

        filtered_column_order = []
        fixed_cols = ['Offer','SVP','CVP','Market_Category',
                    'Business_Line','Domain','Product','Rep_code',
                    'Segment', 'Sub-segment 1', 'Sub-segment 2',
                    'Sub-segment 3']

        for col in fixed_cols:
            if col == 'Offer' or col in selected_columns:
                filtered_column_order.append(col)

        if 'BL_CAGR' in selected_columns and use_custom_cagr in ('yes','both'):
            filtered_column_order.append(bl_cagr_col_name)
        if 'MIF_CAGR' in selected_columns:
            filtered_column_order.append(mif_cagr_col_name)

        filtered_column_order.extend(year_cols)
        column_order = filtered_column_order


        # Filter rows to selected columns
        results_list_filtered = []
        for r in results_list:
            filtered_row = OrderedDict()
            for col in column_order:
                if col in r:
                    filtered_row[col] = r[col]
            results_list_filtered.append(filtered_row)
        results_list = results_list_filtered

        # ── TOTAL row ──
        total_row = OrderedDict()

        total_row['Offer']           = 'TOTAL'
        total_row['SVP']             = ''
        total_row['CVP']             = ''
        total_row['Market_Category'] = ''
        total_row['Business_Line']   = ''
        total_row['Domain']          = ''
        if 'Product'  in column_order: total_row['Product']  = ''
        if 'Rep_code' in column_order: total_row['Rep_code'] = ''
        if use_custom_cagr in ('yes','both'):
            total_row[bl_cagr_col_name] = ''
        total_row[mif_cagr_col_name] = ''


        for col in year_cols:
            y = int(col.split('_')[-1])
            if col.startswith('Revenue_'):
                total_row[col] = f"{tot_actual[y]:,.2f}"
            elif col.startswith('Real_'):
                total_row[col] = f"{tot_real_post[y]:,.2f}"
            elif col.startswith('Predicted_MIF_'):
                total_row[col] = f"{tot_predicted_mif[y]:,.2f}"
            elif col.startswith('Predicted_BL_'):
                total_row[col] = f"{tot_predicted_bl[y]:,.2f}"
            elif col.startswith('Predicted_'):
                total_row[col] = f"{tot_predicted_mif[y]:,.2f}"

        filtered_total = OrderedDict()
        for col in column_order:
            if col in total_row:
                filtered_total[col] = total_row[col]
        results_list.append(filtered_total)

        # ── KPIs ──
        avg_cagr      = np.mean(cagr_values) if cagr_values else 0
        total_row_data = results_list[-1]
        rev_start_col  = f'Revenue_{display_start_year}'
        rev_end_col    = f'Real_{pred_start_year}'
        Actual_cagr    = 0
        if rev_start_col in total_row_data and rev_end_col in total_row_data:
            s = total_row_data[rev_start_col]
            e = total_row_data[rev_end_col]
            if s and s != '':
                Actual_cagr = calculate_cagr(
                    float(s.replace(',','')), float(e.replace(',','')),
                    pred_start_year - display_start_year)

        # ── Chart arrays ──
        chart_actual      = []
        chart_real_post   = []
        chart_predicted   = []       # MIF (or single)
        chart_predicted_bl = []      # BL  (only for 'both')

        for y in all_display_years:
            if y < pred_start_year:
                chart_actual.append(tot_actual[y])
                chart_real_post.append(None)
                chart_predicted.append(None)
                chart_predicted_bl.append(None)
            elif y == pred_start_year:
                seed_value = tot_actual[y] if tot_actual[y] else tot_real_post[y]
                chart_actual.append(seed_value)
                chart_real_post.append(seed_value)
                chart_predicted.append(seed_value)
                chart_predicted_bl.append(seed_value)
            else:
                chart_actual.append(None)
                chart_real_post.append(tot_real_post[y] if has_real_post[y] else None)
                chart_predicted.append(
                    tot_predicted_mif[y] if tot_predicted_mif[y] > 0 else None)
                chart_predicted_bl.append(
                    tot_predicted_bl[y] if tot_predicted_bl[y] > 0 else None)

        def clean_for_json(obj):
            if isinstance(obj, dict):
                return {k: clean_for_json(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [clean_for_json(item) for item in obj]
            elif isinstance(obj, float):
                return 0 if np.isnan(obj) else obj
            elif obj is None:
                return '---'
            return obj

        result_dict = {
            'cagr_start_year'     : cagr_start_year,
            'cagr_end_year'       : cagr_end_year,
            'display_start_year'  : display_start_year,
            'display_end_year'    : display_end_year,
            'pred_start_year'     : pred_start_year,
            'avg_cagr'            : f"{avg_cagr*100:.2f}%",
            'tot_predicted_mif'   : tot_predicted_mif,
            'tot_predicted_bl'    : tot_predicted_bl,
            'actual_cagr'         : f"{Actual_cagr*100:.2f}%",
            'result_count'        : len(results_list) - 1,
            'chart_labels'        : [str(y) for y in all_display_years],
            'chart_actual'        : chart_actual,
            'chart_real_post'     : chart_real_post,
            'chart_predicted'     : chart_predicted,
            'chart_predicted_bl'  : chart_predicted_bl,
            'use_custom_cagr'     : use_custom_cagr,
            'table_data'          : results_list,
            'column_order'        : column_order,
            'cvp_options'         : df_all_data.attrs.get('cvp_options', []),
        }
        return clean_for_json(result_dict)

    except Exception as e:
        import traceback; traceback.print_exc()
        return {'error': str(e)}
