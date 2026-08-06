import pandas as pd
import numpy as np
import os
import re
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import OrderedDict

from mapping import (
    DATA_CAGR_SVP,
    DATA_MARKET_SIZE,
    DATA_ALL_PRODUCTS,
    DATA_MARKET_HIERARCHIES,
    build_category_mapper,
    normalize_region_label,
    clean_segment_cell,
    SEGMENT_LEVEL_COLS,
    REGION_FR,
    REGION_INTL,
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


# Matches a "CAGR 20XX/20YY" (or "CAGR 20XX / 20YY") style header and
# captures the two years, so the BL CAGR period can be read straight off
# the column name instead of requiring a plain "CAGR" header.
CUSTOM_CAGR_YEAR_PATTERN = re.compile(r'cagr.*?(\d{4})\s*/\s*(\d{4})', re.IGNORECASE)


def load_custom_cagr():
    """
    Loads the custom (BL) CAGR file.

    The 'CAGR' column may instead be named 'CAGR 20XX/20YY' (the period the
    BL CAGR was computed over). When that's the case, the years are
    extracted and the column is renamed back to 'CAGR' — but only on the
    in-memory dataframe; the original file on disk is left untouched.

    Returns a tuple: (custom_cagr_dict, bl_cagr_start_year, bl_cagr_end_year).
    All three are None when no valid custom CAGR file/data is found.
    """
    try:
        filepath = os.path.join(DATA_DIR, 'custom_cagr.xlsx')
        if not os.path.exists(filepath):
            print("No custom CAGR file found")
            return None, None, None
        df = pd.read_excel(filepath)
        df.columns = df.columns.str.strip().str.title()

        offer_col = None
        cagr_col  = None
        bl_start_year = None
        bl_end_year   = None
        for col in df.columns:
            if col.lower() == 'offer':
                offer_col = col
            elif col.lower().startswith('cagr'):
                cagr_col = col

        # Rename the CAGR-with-years column back to plain 'CAGR' on the
        # working dataframe only — the file on disk is never modified.
        if cagr_col != 'Cagr':
            years_cagr = cagr_col.split(' ')[1]
            bl_start_year = years_cagr.split('/')[0]
            bl_end_year = years_cagr.split('/')[1]
            df = df.rename(columns={cagr_col: 'Cagr'})
            cagr_col = 'Cagr'

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
            yr_info = f" ({bl_start_year}/{bl_end_year})" if bl_start_year and bl_end_year else ""
            print(f"✓ Custom CAGR loaded{yr_info}: {len(custom_cagr_dict)} offers")
            return custom_cagr_dict, bl_start_year, bl_end_year
        return None, None, None
    except Exception as e:
        print(f"✗ Error loading custom CAGR: {e}")
        return None, None, None


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
    """Rep_Code-level category mapper: Rep_Code / catégorie / Segment /
    Sub-segment 1-3 (separately for FR and INT — see mapping.py), built by
    combining the hierarchical Rep_Code→catégorie matching (CAGR SVP + All
    Products) with the semantic catégorie→segment matching (Market Size),
    run per region. Returned columns are renamed to match this pipeline's
    conventions: Market_Category (was catégorie), Rep_Code,
    {Segment,Sub-segment 1-3}_{FR,INT}."""
    try:
        mapper = build_category_mapper(force_rebuild=force_rebuild)
        if mapper is None or mapper.empty:
            return None
        mapper = mapper.rename(columns={'catégorie': 'Market_Category'})
        mapper['Market_Category'] = mapper['Market_Category'].fillna('').astype(str).str.rstrip()
        region_seg_cols = [f"{c}_{r}" for r in (REGION_FR, REGION_INTL) for c in SEGMENT_LEVEL_COLS]
        for col in region_seg_cols:
            if col in mapper.columns:
                mapper[col] = mapper[col].apply(clean_segment_cell)
        has_fr  = (mapper[f'Segment_{REGION_FR}']   != '---').sum() if f'Segment_{REGION_FR}'   in mapper.columns else 0
        has_int = (mapper[f'Segment_{REGION_INTL}'] != '---').sum() if f'Segment_{REGION_INTL}' in mapper.columns else 0
        print(f"✓ Category Mapper: {len(mapper)} Rep_Codes "
              f"({(mapper['Market_Category'] != '').sum()} with catégorie, "
              f"{has_fr} with a France segment, {has_int} with an International segment)")
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

        # ── Category mapper (Rep_Code → catégorie → Segment/Sub-segment, by region) ──
        # Built from CAGR SVP (hierarchical match, test1.py idea) + Market
        # Size (semantic match, test2.py idea). See mapping.py.
        print(f"\n📋 Loading Category Mapper...")
        df_repcode_mapper = load_category_mapper()
        if df_repcode_mapper is None:
            print(f"  ✗ Category mapper failed to load"); return None
        print(f"  ✓ Loaded {len(df_repcode_mapper)} Rep_Codes")

        # Roll the Rep_Code-level mapper up to Offer level (first non-empty
        # value per Offer wins) so the existing Offer-level pivot/merge
        # machinery below keeps working. Rep_Code-level detail itself is
        # only needed for revenue splitting, which already happens
        # separately via detailed_revenue in resultat_data().
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

        # Roll up every column the mapper produced (catégorie plus the
        # per-region Segment / Sub-segment 1-3 hierarchy: *_FR and *_INT).
        region_segment_cols = [f"{c}_{r}" for r in (REGION_FR, REGION_INTL) for c in SEGMENT_LEVEL_COLS
                                if f"{c}_{r}" in df_repcode_mapper.columns]
        agg_spec = {'Market_Category': ('Market_Category', _first_nonempty)}
        for col in region_segment_cols:
            agg_spec[col] = (col, _first_nonempty)

        df_mapper = (
            df_repcode_mapper.dropna(subset=['Offer_Clean'])
            .groupby('Offer_Clean', as_index=False)
            .agg(**agg_spec)
            .rename(columns={'Offer_Clean': 'Offer'})
        )
        print(f"  ✓ Rolled up to {len(df_mapper)} offers "
              f"(segment columns carried through: {region_segment_cols or 'none'})")

        df_mapper['SVP']           = df_mapper['Offer'].map(offer_to_svp).fillna('Unknown')
        df_mapper['CVP']           = df_mapper['Offer'].map(offer_to_cvp).fillna('Unknown')
        df_mapper['Business_Line'] = df_mapper['Offer'].map(offer_to_business_line).fillna('---')
        df_mapper['Domain']        = df_mapper['Offer'].map(offer_to_domain).fillna('---')
        print(f"  ✓ Added SVP, CVP, Business_Line, Domain to mapper")

        # ── Market size (pivoted on ID, matched to catégorie semantically) ──
        print(f"\n📈 Loading Market Size...")
        df_ms = pd.read_excel(DATA_MARKET_SIZE,
                              sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
        # ── Market size (pivoted on Segment/Sub-segment hierarchy + Region) ──
        # Previously this was pivoted on a Market Sizing 'ID' column and
        # joined to offers on that ID. That's been replaced with a direct
        # join on the (Segment, Sub-segment 1-3, Region) tuple: 'ID' isn't
        # guaranteed unique per segment/region in the source workbook, and
        # when it wasn't, grouping Million EUR by ID silently summed
        # unrelated rows together — producing inflated market-size figures
        # (e.g. a sub-segment showing ~1900 M€ when the source file only
        # supports ~140 M€ for it). Region is included in the key because
        # France and International carry different Million EUR figures for
        # what's conceptually the same segment tree.
        ms_seg_cols = [c for c in SEGMENT_LEVEL_COLS if c in df_ms.columns]
        if not ms_seg_cols or 'Region' not in df_ms.columns:
            print(f"  ✗ Market Size file missing Segment/Sub-segment or Region columns"); return None
        df_ms = df_ms[[*ms_seg_cols, "Region", "Year", "Million EUR"]].copy()
        for c in ms_seg_cols:
            df_ms[c] = df_ms[c].apply(clean_segment_cell)
        df_ms['Region'] = df_ms['Region'].apply(normalize_region_label)
        df_ms = df_ms.dropna(subset=["Region", "Year", "Million EUR"])
        df_ms = df_ms.groupby([*ms_seg_cols, "Region", "Year"], as_index=False).agg({"Million EUR":"sum"})
        df_ms['Year'] = df_ms['Year'].astype(int)
        df_ms_pivot = df_ms.pivot_table(
            index=[*ms_seg_cols, "Region"], columns="Year",
            values="Million EUR", fill_value=0).reset_index()
        df_ms_pivot.columns = [
            f"Million EUR_{c}" if isinstance(c, int) else c
            for c in df_ms_pivot.columns]
        print(f"  ✓ Pivoted to {len(df_ms_pivot)} segment/region combinations")

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

        # Each revenue row already carries its own Region (FR/INT). Resolve
        # the single Segment/Sub-segment 1-3 hierarchy to use for that row
        # by picking the matching _FR or _INT column the mapper produced —
        # this is what makes market size region-aware end to end.
        for seg_col in SEGMENT_LEVEL_COLS:
            fr_col, int_col = f"{seg_col}_{REGION_FR}", f"{seg_col}_{REGION_INTL}"
            fr_series  = df[fr_col]  if fr_col  in df.columns else pd.Series('---', index=df.index)
            int_series = df[int_col] if int_col in df.columns else pd.Series('---', index=df.index)
            df[seg_col] = np.where(df['Region'] == REGION_FR, fr_series, int_series)
            df[seg_col] = pd.Series(df[seg_col], index=df.index).apply(clean_segment_cell)

        # Clean categoricals
        categorical_cols = ['SVP', 'CVP', 'Market_Category',
                            'Business_Line', 'Domain', 'Product', 'Rep_Code',
                            *SEGMENT_LEVEL_COLS]
        for col in categorical_cols:
            if col in df.columns:
                df[col] = df[col].fillna('---').replace('nan','---')
                df[col] = df[col].astype(str).str.replace('nan','---')

        revenue_cols = [c for c in df.columns if c.startswith('Revenue_')]
        for col in revenue_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df = df.merge(df_ms_pivot, on=[*ms_seg_cols, 'Region'], how='left')
        million_eur_cols = [c for c in df.columns if c.startswith('Million EUR_')]
        for col in million_eur_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

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
        custom_cagr_dict   = None
        bl_cagr_start_year = None
        bl_cagr_end_year   = None
        if use_custom_cagr in ('yes', 'both'):
            custom_cagr_dict, bl_cagr_start_year, bl_cagr_end_year = load_custom_cagr()
            if custom_cagr_dict:
                yr_info = f" ({bl_cagr_start_year}/{bl_cagr_end_year})" \
                    if bl_cagr_start_year and bl_cagr_end_year else ""
                print(f"\n📊 Applying custom CAGR (BL){yr_info} to {len(custom_cagr_dict)} offers\n")

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

        # Sub-segment filter: filters on whether a Market Sizing segment
        # was successfully assigned via the semantic match, for this row's
        # own region (Segment/Sub-segment 1-3 were already resolved to the
        # FR or INT variant per-row in load_all_data, based on each row's
        # Region) — filtering on Market_Category (catégorie) instead was
        # wrong, since catégorie comes from the separate hierarchical
        # match and can be set even when the semantic match to Market
        # Sizing failed (leaving Segment, and the rest of the hierarchy,
        # blank).
        if subseg_filter == 'with_subseg':
            df = df[df['Segment'].notna() &
                    (df['Segment'] != '---') &
                    (df['Segment'] != 'nan') &
                    (df['Segment'] != '')]
        elif subseg_filter == 'no_subseg':
            df = df[df['Segment'].isna() |
                    (df['Segment'] == '---') |
                    (df['Segment'] == 'nan') |
                    (df['Segment'] == '')]

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

        # ── "Real"/actual interval bounds — ALWAYS the true first/last
        # historical revenue years (from data.xlsx), regardless of the
        # display_start_year/display_end_year the person has chosen. These
        # are used for the "CAGR of Actual Interval" KPIs (both per-offer
        # and TOTAL). Deliberately independent of the display window: the
        # display interval only controls which years are *shown*, not
        # which years the real-CAGR calculation is anchored to — anchoring
        # it to display_start_year instead (the old behavior) meant that
        # once the person moved the display window past the true first
        # revenue year, that year's Revenue_ column stopped being emitted
        # at all, and the CAGR calculation would silently fall back to a
        # near-zero start value and blow up (e.g. 2.5% turning into 1000%).
        real_start_year = rev_years[0]  if rev_years else None
        real_end_year   = rev_years[-1] if rev_years else None

        def _actual_interval_cagr(bounds):
            if real_start_year is None or real_end_year is None or real_start_year == real_end_year:
                return 0
            return calculate_cagr(bounds.get(real_start_year, 0.0),
                                   bounds.get(real_end_year, 0.0),
                                   real_end_year - real_start_year)

        # ── CAGR column names ──
        mif_cagr_col_name = f'MIF_CAGR_{cagr_start_year}/{cagr_end_year}'
        # shown when use_custom_cagr in (yes, both); suffixed with the BL
        # CAGR's own period when known (it can differ from the MIF period)
        bl_cagr_col_name  = f'BL_CAGR_{bl_cagr_start_year}/{bl_cagr_end_year}' \
            if bl_cagr_start_year and bl_cagr_end_year else 'BL_CAGR'

        results_list = []
        cagr_values  = []
        _cagr_seen_offers = set()
        # Market size (Million EUR) is a property of the Market Sizing
        # Segment/Sub-segment hierarchy (within a region), not of the
        # offer/rep-code — several offers (or, at rep_code aggregation,
        # several rep-codes of one offer) can share the same segment.
        # Track which (Segment, Sub-segment 1-3, Region) combos have
        # already been counted so the TOTAL row sums each market segment
        # once instead of once per row.
        _market_size_seen_keys = set()

        tot_actual    = {y: 0.0 for y in all_display_years}
        tot_real_post = {y: 0.0 for y in all_display_years}
        # For 'both' mode we track two prediction totals
        tot_predicted_mif = {y: 0.0 for y in all_display_years}
        tot_predicted_bl  = {y: 0.0 for y in all_display_years}
        tot_million_eur   = {y: 0.0 for y in all_display_years}
        has_real_post     = {y: False for y in all_display_years}

        # kpiTotalCagrMif/kpiTotalCagrBL must always be computable from the
        # CAGR interval (cagr_start_year → cagr_end_year), independent of
        # the display interval (all_display_years) — the person can move
        # the display window without losing the total-CAGR KPI. Otherwise
        # tot_predicted_mif/tot_predicted_bl are only keyed over
        # all_display_years, so if either CAGR boundary year falls outside
        # the display window (e.g. display 2022→2029 but CAGR 2025→2030),
        # that key would be missing and the KPI would silently read as
        # undefined on the frontend. Make sure both boundary years always
        # have an entry.
        extra_cagr_years = sorted(
            {cagr_start_year, cagr_end_year} - set(all_display_years))
        for _y in extra_cagr_years:
            tot_predicted_mif.setdefault(_y, 0.0)
            tot_predicted_bl.setdefault(_y, 0.0)

        # Real/actual-interval revenue totals, tracked separately from
        # tot_actual (which is keyed by all_display_years and therefore
        # blind to years outside the display window) — see comment above.
        tot_actual_bounds    = {real_start_year: 0.0, real_end_year: 0.0} \
            if real_start_year is not None else {}
        _offer_actual_bounds = {}   # offer -> {real_start_year: sum, real_end_year: sum}

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

            # Accumulate real-interval bounds (true first/last historical
            # revenue years), independent of the display window.
            if real_start_year is not None:
                sv = actual.get(real_start_year, 0.0)
                ev = actual.get(real_end_year, 0.0)
                tot_actual_bounds[real_start_year] += sv
                tot_actual_bounds[real_end_year]   += ev
                ob = _offer_actual_bounds.setdefault(
                    offer, {real_start_year: 0.0, real_end_year: 0.0})
                ob[real_start_year] += sv
                ob[real_end_year]   += ev

            # Market size (Million EUR of MIF), per display year
            market_key = (
                row.get('Segment', '---'), row.get('Sub-segment 1', '---'),
                row.get('Sub-segment 2', '---'), row.get('Sub-segment 3', '---'),
                row.get('Region', '---'),
            )
            ms_year_vals = {}
            for y in all_display_years:
                msc = f"Million EUR_{y}"
                if msc in row.index:
                    v = row[msc]
                    ms_year_vals[y] = float(v) if not pd.isna(v) else 0.0
            count_ms_in_total = (
                market_key[0] not in (None, '', '---', 'nan') and
                market_key not in _market_size_seen_keys
            )
            if count_ms_in_total:
                _market_size_seen_keys.add(market_key)

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
            result_row['Segment']         = row.get('Segment', '---')
            result_row['Sub-segment 1']   = row.get('Sub-segment 1', '---')
            result_row['Sub-segment 2']   = row.get('Sub-segment 2', '---')
            result_row['Sub-segment 3']   = row.get('Sub-segment 3', '---')
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

                # Market size (Million EUR of MIF) — shown on every row
                # (each row displays its own market segment's size), but
                # only rolled into the TOTAL once per unique (Segment,
                # Sub-segment 1-3, Region) combination.
                if y in ms_year_vals:
                    #result_row[f'Million EUR_{y}'] = f"{ms_year_vals[y]:,.2f}"
                    if count_ms_in_total:
                        tot_million_eur[y] += ms_year_vals[y]

            # Extend tot_predicted_mif/tot_predicted_bl to cover any CAGR
            # boundary year outside the display window (see extra_cagr_years
            # comment above) so the total-CAGR KPIs always resolve, no
            # matter where the display interval is set. Not shown in the
            # table — only rolled into the totals used by the KPI.
            for y in extra_cagr_years:
                act_val = actual.get(y, None)
                if y < pred_start_year:
                    if act_val is not None:
                        tot_predicted_mif[y] += act_val
                        tot_predicted_bl[y]  += act_val
                else:
                    steps_mif = y - seed_y
                    pred_mif  = seed_v * ((1 + mif_cagr) ** steps_mif)
                    pred_bl   = seed_v * ((1 + bl_cagr)  ** steps_mif)
                    if y == pred_start_year:
                        pred_mif = act_val if act_val is not None else pred_mif
                        pred_bl  = pred_mif
                    if use_custom_cagr == 'both':
                        tot_predicted_mif[y] += pred_mif
                        tot_predicted_bl[y]  += pred_bl
                    else:
                        pred_val = pred_bl if use_custom_cagr == 'yes' else pred_mif
                        tot_predicted_mif[y] += pred_val
                        tot_predicted_bl[y]  += pred_val

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
            #if any(f'Million EUR_{y}' in r for r in results_list):
            #    year_cols.append(f'Million EUR_{y}')

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
        total_row['Segment']         = ''
        total_row['Sub-segment 1']   = ''
        total_row['Sub-segment 2']   = ''
        total_row['Sub-segment 3']   = ''
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
            #elif col.startswith('Million EUR_'):
            #    total_row[col] = f"{tot_million_eur[y]:,.2f}"

        filtered_total = OrderedDict()
        for col in column_order:
            if col in total_row:
                filtered_total[col] = total_row[col]
        results_list.append(filtered_total)

        # ── KPIs ──
        avg_cagr      = np.mean(cagr_values) if cagr_values else 0

        # "CAGR of Actual Interval" (TOTAL) — always anchored to the true
        # first/last historical revenue years (real_start_year /
        # real_end_year), never to display_start_year/display_end_year.
        # See the tot_actual_bounds comment above for why.
        Actual_cagr = _actual_interval_cagr(tot_actual_bounds)

        # Per-offer actual-interval CAGR, for the "Selected Offer CAGR of
        # Actual Interval" KPI — computed here (server-side, using
        # calculate_cagr's safe guards) instead of being reconstructed in
        # JS from Revenue_/Real_ table cells, which silently go missing
        # once the display window moves past real_start_year and used to
        # make the KPI blow up to absurd values.
        actual_interval_cagr_by_offer = {
            off: f"{_actual_interval_cagr(bounds)*100:.2f}%"
            for off, bounds in _offer_actual_bounds.items()
        }

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
            'real_cagr_start_year': real_start_year,
            'real_cagr_end_year'  : real_end_year,
            'bl_cagr_start_year'  : bl_cagr_start_year,
            'bl_cagr_end_year'    : bl_cagr_end_year,
            'avg_cagr'            : f"{avg_cagr*100:.2f}%",
            'tot_predicted_mif'   : tot_predicted_mif,
            'tot_predicted_bl'    : tot_predicted_bl,
            'actual_cagr'         : f"{Actual_cagr*100:.2f}%",
            'actual_interval_cagr_by_offer': actual_interval_cagr_by_offer,
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