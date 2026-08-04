"""
Category mapping pipeline.

Combines two ideas:

  * test1.py — hierarchical matching that assigns a "catégorie" to every
    product row by combining hierarchy columns (Business_Line ->
    Sub_Business_Line -> Domain -> Sub_Domain -> Offer_Line -> Offer) and
    falling back to progressively less specific combinations, using the
    multi-sheet CAGR SVP workbook as the category source and "All
    products requested.xlsx" as the Rep_Code dimension table. Since the
    CAGR SVP workbook's own sheets are each tagged with a Region (encoded
    in the sheet name), this match is run once per region, producing
    catégorie_FR and catégorie_INT independently — the same Rep_Code can
    land on a different catégorie label per region. Where no
    region-specific match was found for International, catégorie_INT
    falls back to catégorie_FR.

  * test1.py — a direct, exact-match lookup of catégorie -> Segment /
    Sub-segment 1-3, read straight from a pre-identified hierarchy table
    ("market size identified hierarchies.xlsx", keyed on 'catégorie
    (ID)'). This replaces an earlier semantic-matching approach (TF-IDF
    char n-grams + Truncated SVD/LSA + cosine similarity) that inferred
    the closest Market Sizing segment for each catégorie; the identified
    hierarchies table is now the trusted source instead, so no
    similarity/fuzzy matching happens here at all.

Combined, they produce a single generated mapper keyed at Rep_Code
granularity:

    Rep_Code | catégorie | Segment_FR | Sub-segment 1_FR | Sub-segment 2_FR |
              Sub-segment 3_FR | Segment_INT | Sub-segment 1_INT |
              Sub-segment 2_INT | Sub-segment 3_INT

"catégorie" (the FR one) comes from the hierarchical match. The Segment /
Sub-segment 1-3 columns for each region come from looking up that
region's catégorie (catégorie_FR / catégorie_INT) directly in the
identified hierarchies table — this table only carries one hierarchy per
catégorie ID (its columns happen to be named with a "_FR" suffix), so the
same lookup table is used for both regions; what makes the result
region-aware is that catégorie_FR and catégorie_INT can point at
different rows in that table.

IMPORTANT — region-aware matching: the Market Sizing workbook contains
separate rows (and separate Million EUR figures) for France vs.
International for what is conceptually "the same" segment tree. Once the
Segment/Sub-segment 1-3 tree is known per region (from the identified
hierarchies lookup above), pipeline.py joins Million EUR directly on the
(Segment, Sub-segment 1, Sub-segment 2, Sub-segment 3, Region) tuple
rather than on a Market Sizing 'ID' column, since 'ID' isn't guaranteed
unique per segment/region in the source workbook (a shared ID value
silently summed unrelated rows together and produced wildly inflated
market-size figures downstream).

The CAGR SVP workbook is now the primary source of truth for categories
(it replaces the old flat "CAGR Mapper.xlsx").
"""
import os
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

DATA_CAGR_SVP            = os.path.join(DATA_DIR, "CAGR SVP.xlsx")
DATA_MARKET_SIZE         = os.path.join(DATA_DIR, "Market Sizing - 3Q 2025.xlsx")
DATA_ALL_PRODUCTS        = os.path.join(DATA_DIR, "All products requested.xlsx")
DATA_MARKET_HIERARCHIES  = os.path.join(DATA_DIR, "market size identified hierarchies.xlsx")

# Cache of the built mapper, so it doesn't have to be rebuilt on every
# request. Call build_category_mapper(force_rebuild=True) whenever CAGR
# SVP, All Products, Market Size, or the identified hierarchies file change.
GENERATED_MAPPER_PATH = os.path.join(DATA_DIR, "generated_category_mapper.xlsx")

# Canonical region labels used across the whole pipeline (must match the
# 'FR' / 'INT' values pipeline.py assigns to the revenue data).
REGION_FR   = 'FR'
REGION_INTL = 'INT'
SEGMENT_LEVEL_COLS = ["Segment", "Sub-segment 1", "Sub-segment 2", "Sub-segment 3"]


# ─────────────────────────────────────────────
# Small helpers (from test1.py)
# ─────────────────────────────────────────────

def _check_space_at_the_end(x):
    """Remove trailing space from string (test1.py)."""
    if isinstance(x, str) and x.endswith(' '):
        return x[:-1]
    return x


def clean_segment_cell(v) -> str:
    """Normalize a single Segment/Sub-segment cell.

    Used on BOTH sides of the Million EUR join — once here, when a
    Segment/Sub-segment value from the identified hierarchies lookup is
    copied into the category mapper, and again in pipeline.py, when the
    same cell is re-read directly off the Market Sizing sheet to build
    the Million EUR lookup. Since Million EUR is joined on this value as
    a natural key (see module docstring), both sides MUST clean it
    identically or the join will silently miss rows — so the two call
    sites are kept in lockstep via this one shared function rather than
    reimplemented separately.
    """
    if pd.isna(v):
        return '---'
    s = str(v).rstrip()
    if s in ('', 'nan', 'None'):
        return '---'
    return s


def normalize_region_label(raw):
    """Map a raw Region cell (from the Market Sizing workbook, or from a
    CAGR SVP sheet name) to the pipeline's canonical 'FR' / 'INT' labels.
    Returns None if the value isn't recognized (callers should surface
    unrecognized labels rather than silently guessing, since a wrong
    guess here would misfile an entire region's data)."""
    if pd.isna(raw):
        return None
    s = str(raw).strip().lower()
    if s in ('fr', 'france', 'français', 'francais', 'domestic'):
        return REGION_FR
    if s in ('int', 'intl', "i'ntl", 'international', 'row', 'rest of world',
              'ex-france', 'export', 'hors france'):
        return REGION_INTL
    # Lenient fallback for label variants we haven't enumerated above.
    if 'international' in s or s.startswith('int'):
        return REGION_INTL
    if 'franc' in s or s == 'fr':
        return REGION_FR
    return None


# ─────────────────────────────────────────────
# Step 1 (test1.py's progressive-matching approach): hierarchical
# Rep_Code → catégorie matching
# ─────────────────────────────────────────────

# Order matters: this is also used positionally in load_cagr_svp_categories
# to translate an "Unnamed: N" header back into its real pr_* name, based
# on how many digits precede the dot in that column's own values (e.g. a
# column full of "1.2.3 Something" values is level 3 -> pr_domain).
REQUIRED_COLUMNS_MAPPER = ["pr_business_line", "pr_sub_business_line", "pr_domain",
                           "pr_sub_domain", "pr_offer_line", "pr_offer", "catégorie",
                           "SVP", "Region"]


def load_cagr_svp_categories(path=DATA_CAGR_SVP) -> pd.DataFrame:
    """Load every relevant sheet of the CAGR SVP workbook (skipping MIF*/
    *GLOBAL* sheets), reconstruct the pr_* hierarchy column names from
    each sheet's unnamed/numbered headers, derive SVP/Region from the
    sheet name itself, and concatenate into one dataframe of
    pr_business_line / pr_sub_business_line / pr_domain / pr_sub_domain /
    pr_offer_line / pr_offer / catégorie / SVP / Region (test1.py)."""
    sheets = pd.ExcelFile(path).sheet_names
    relevant_sheets = [s for s in sheets if not s.startswith("MIF") and "GLOBAL" not in s]

    df_list = []
    for sheet_name in relevant_sheets:
        df_temp = pd.read_excel(path, sheet_name=sheet_name, skiprows=2)
        if "catégorie" not in df_temp.columns:
            continue
        df_temp = df_temp.dropna(subset=["catégorie"])
        df_temp["catégorie"] = df_temp["catégorie"].apply(_check_space_at_the_end)

        # SVP / Region come from the sheet name itself, e.g. "SVPName Region".
        s = sheet_name.split(' ')
        df_temp["SVP"] = s[0]
        df_temp["Region"] = s[1] if len(s) > 1 else ''

        for col in list(df_temp.columns):
            if type(col) != str:
                continue
            if col.endswith(' '):
                df_temp.rename(columns={col: col[:-1]}, inplace=True)
                col = col[:-1]
            if "Unnamed: " not in col:
                continue
            if df_temp[col].isna().all():
                df_temp.drop(columns=[col], inplace=True)
                continue
            # How many digits precede the dot (e.g. "1.2.3 Something" -> 3)
            # tells us which pr_* level this unnamed column actually is.
            extracted = df_temp[col].dropna().astype(str).str.extract(r'(\d+)\.')[0].dropna()
            if extracted.empty:
                df_temp.drop(columns=[col], inplace=True)
                continue
            count = extracted.str.len().mode()[0]
            if count <= 0 or count >= 7:
                df_temp.drop(columns=[col], inplace=True)
                continue
            df_temp.rename(columns={col: REQUIRED_COLUMNS_MAPPER[int(count - 1)]}, inplace=True)

        # Sheets that don't carry every pr_* level (e.g. no sub-domain
        # split) still need the column present so concat/matching works.
        for c in REQUIRED_COLUMNS_MAPPER:
            if c not in df_temp.columns:
                df_temp[c] = np.nan

        df_list.append(df_temp[REQUIRED_COLUMNS_MAPPER])

    if not df_list:
        return pd.DataFrame(columns=REQUIRED_COLUMNS_MAPPER)

    all_df = pd.concat(df_list, ignore_index=True)
    all_df["catégorie"] = all_df["catégorie"].fillna('')
    return all_df


DIM_PRO_HIERARCHY_COLS = ['Business_Line', 'Sub_Business_Line', 'Domain',
                          'Sub_Domain', 'Offer_Line', 'Offer']

# Maps df_dim_pro's hierarchy columns to their df_mapper (pr_*) equivalents.
_COLUMN_MAPPING = {
    'Business_Line':     'pr_business_line',
    'Sub_Business_Line': 'pr_sub_business_line',
    'Domain':            'pr_domain',
    'Sub_Domain':        'pr_sub_domain',
    'Offer_Line':        'pr_offer_line',
    'Offer':             'pr_offer',
}


def progressive_matching(df_dim_pro: pd.DataFrame, df_mapper: pd.DataFrame,
                          max_iterations=6) -> pd.DataFrame:
    """Progressively match unmatched rows (test1.py's approach) by
    combining the hierarchy columns into a single text key and reducing
    that combination from the most specific (all 6 columns) down to just
    Business_Line. Only rows still marked "UNMATCHED" are attempted at
    each level, and each successful level records how many columns it
    took to match ('match_level': 0 = most specific)."""
    column_sequences = [
        ['Business_Line', 'Sub_Business_Line', 'Domain', 'Sub_Domain', 'Offer_Line', 'Offer'],
        ['Business_Line', 'Sub_Business_Line', 'Domain', 'Sub_Domain', 'Offer_Line'],
        ['Business_Line', 'Sub_Business_Line', 'Domain', 'Sub_Domain'],
        ['Business_Line', 'Sub_Business_Line', 'Domain'],
        ['Business_Line', 'Sub_Business_Line'],
        ['Business_Line'],
    ][:max_iterations]

    df_dim_pro['match_level'] = None

    for level, cols in enumerate(column_sequences):
        unmatched_mask = df_dim_pro['catégorie'] == "UNMATCHED"
        if not unmatched_mask.any():
            print(f"  ✓ All rows matched at level {level}")
            break

        unmatched_df = df_dim_pro[unmatched_mask].copy()
        unmatched_df['temp_combined'] = unmatched_df[cols] \
            .fillna('---').apply(lambda x: ' '.join(x.astype(str)), axis=1)

        mapper_cols = [_COLUMN_MAPPING[c] for c in cols]
        mapper_level_df = df_mapper[mapper_cols + ['catégorie']].drop_duplicates(subset=mapper_cols)
        mapper_level = dict(zip(
            mapper_level_df[mapper_cols].fillna('---').apply(lambda x: ' '.join(x.astype(str)), axis=1),
            mapper_level_df['catégorie']
        ))

        matched_mask = unmatched_df['temp_combined'].isin(mapper_level.keys())
        matched_indices = unmatched_df[matched_mask].index

        df_dim_pro.loc[matched_indices, 'catégorie'] = unmatched_df.loc[matched_indices, 'temp_combined'].map(mapper_level)
        df_dim_pro.loc[matched_indices, 'match_level'] = level

        matched_count = int(matched_mask.sum())
        print(f"  Level {level} ({len(cols)} cols): matched {matched_count} rows | "
              f"total matched so far: {(df_dim_pro['catégorie'] != 'UNMATCHED').sum()}")

    return df_dim_pro


def assign_categories_to_rep_codes(cagr_svp_path=DATA_CAGR_SVP,
                                    all_products_path=DATA_ALL_PRODUCTS) -> pd.DataFrame:
    """Rep_Code → catégorie_FR / catégorie_INT, using All products
    requested.xlsx as the Rep_Code dimension table and the CAGR SVP
    workbook as the category source, matched via test1.py's progressive
    (multi-column, most- to least-specific) matching.

    The CAGR SVP workbook's sheets are each tagged with a Region (parsed
    out of the sheet name by load_cagr_svp_categories), so the match is
    run twice — once against France-tagged sheets, once against
    International-tagged sheets — producing two independent catégorie
    assignments per Rep_Code. Where a Rep_Code has no region-specific
    International match, catégorie_INT falls back to catégorie_FR.

    Returns columns: Rep_Code, catégorie_FR, catégorie_INT."""
    df_mapper_all = load_cagr_svp_categories(cagr_svp_path)
    df_mapper_all = df_mapper_all.copy()
    df_mapper_all["Region_norm"] = df_mapper_all["Region"].apply(normalize_region_label)

    df_dim_pro = pd.read_excel(all_products_path)
    df_dim_pro.columns = df_dim_pro.columns.str.strip()

    rep_col = next((c for c in df_dim_pro.columns if c.lower() == 'rep_code'), None)
    if rep_col is None:
        raise ValueError("All products requested.xlsx must have a 'Rep_Code' column")
    if rep_col != 'Rep_Code':
        df_dim_pro = df_dim_pro.rename(columns={rep_col: 'Rep_Code'})

    missing = [c for c in DIM_PRO_HIERARCHY_COLS if c not in df_dim_pro.columns]
    if missing:
        raise ValueError(f"All products requested.xlsx is missing required column(s): {missing}")

    base_cols = ['Rep_Code', *DIM_PRO_HIERARCHY_COLS]

    per_region_cat = {}
    for region in (REGION_FR, REGION_INTL):
        region_mapper = df_mapper_all[df_mapper_all["Region_norm"] == region]
        df_region = df_dim_pro[base_cols].copy()
        df_region["catégorie"] = "UNMATCHED"
        if region_mapper.empty:
            print(f"  ⚠ CAGR SVP: no sheets recognized as region '{region}' — "
                  f"catégorie_{region} left unmatched for all Rep_Codes")
        else:
            df_region = progressive_matching(df_region, region_mapper)
        df_region["catégorie"] = df_region["catégorie"].replace("UNMATCHED", '').fillna('')
        df_region = df_region[["Rep_Code", "catégorie"]].drop_duplicates(subset=["Rep_Code"], keep='first')
        per_region_cat[region] = df_region.set_index("Rep_Code")["catégorie"]

    out = df_dim_pro[["Rep_Code"]].drop_duplicates(subset=["Rep_Code"], keep='first').reset_index(drop=True)
    out["catégorie_FR"]  = out["Rep_Code"].map(per_region_cat[REGION_FR]).fillna('')
    out["catégorie_INT"] = out["Rep_Code"].map(per_region_cat[REGION_INTL]).fillna('')
    # Fall back to catégorie_FR wherever no region-specific INT match exists.
    out["catégorie_INT"] = out["catégorie_INT"].mask(out["catégorie_INT"] == '', out["catégorie_FR"])

    print(f"  ✓ Hierarchical (progressive) match: {len(out)} Rep_Codes, "
          f"{(out['catégorie_FR'] != '').sum()} with catégorie_FR, "
          f"{(out['catégorie_INT'] != '').sum()} with catégorie_INT (after FR fallback)")
    return out


# ─────────────────────────────────────────────
# Step 2 (test1.py idea): direct catégorie → market segment lookup,
# read from the pre-identified hierarchies table (no similarity/fuzzy
# matching — a plain exact-match join on catégorie).
# ─────────────────────────────────────────────

def load_category_hierarchies(path=DATA_MARKET_HIERARCHIES) -> dict:
    """Load "market size identified hierarchies.xlsx" (Sheet1) — a table
    that has already been hand/otherwise identified as the correct
    catégorie -> Segment/Sub-segment 1-3 mapping — and return it as
    {catégorie: {'Segment_FR': ..., 'Sub-segment 1_FR': ..., ...}}
    (test1.py's approach).

    The table only carries one hierarchy per catégorie ID (its columns
    happen to be named with a "_FR" suffix); the same lookup table is
    used for both France and International — see build_category_mapper,
    which looks each region's own catégorie (catégorie_FR / catégorie_INT)
    up in this same table, so the region-awareness comes from which
    catégorie is used as the key, not from the table itself.
    """
    df = pd.read_excel(path, sheet_name='Sheet1')
    if 'catégorie (ID)' not in df.columns:
        raise ValueError("market size identified hierarchies.xlsx must have a "
                          "'catégorie (ID)' column")

    hierarchy_cols = [f"{c}_FR" for c in SEGMENT_LEVEL_COLS]
    missing = [c for c in hierarchy_cols if c not in df.columns]
    if missing:
        raise ValueError(f"market size identified hierarchies.xlsx is missing "
                          f"column(s): {missing}")

    for c in hierarchy_cols:
        df[c] = df[c].apply(clean_segment_cell)

    df = df.drop_duplicates(subset=['catégorie (ID)'], keep='first')
    df = df.set_index('catégorie (ID)')
    return df[hierarchy_cols].to_dict('index')


# ─────────────────────────────────────────────
# Orchestration
# ─────────────────────────────────────────────

def build_category_mapper(force_rebuild=False) -> pd.DataFrame:
    """Build (or load the cached) Rep_Code-level category mapper.

    Returns a dataframe with columns: Rep_Code, catégorie (= catégorie_FR),
    catégorie_FR, catégorie_INT, and
    {Segment,Sub-segment 1,Sub-segment 2,Sub-segment 3}_{FR,INT}.
    """
    region_seg_cols = [f"{c}_{r}" for r in (REGION_FR, REGION_INTL) for c in SEGMENT_LEVEL_COLS]
    required_cols = {'Rep_Code', 'catégorie', *region_seg_cols}

    if not force_rebuild and os.path.exists(GENERATED_MAPPER_PATH):
        try:
            cached = pd.read_excel(GENERATED_MAPPER_PATH)
            # A cache written before this fix (e.g. still built from the
            # semantic-matching step, or missing the region-split
            # catégorie columns) must NOT be treated as valid — otherwise
            # the old mapping would keep being served forever after this
            # change.
            if required_cols.issubset(cached.columns):
                return cached
        except Exception:
            pass  # fall through and rebuild

    print("\n" + "=" * 60)
    print("BUILDING CATEGORY MAPPER (Rep_Code → catégorie → Segment/Sub-segment, by region)")
    print("=" * 60)

    # Rep_Code -> catégorie_FR / catégorie_INT (test1.py's progressive
    # matching, run once per CAGR SVP region).
    rep_to_cat = assign_categories_to_rep_codes()

    # catégorie -> Segment/Sub-segment 1-3, read directly from the
    # pre-identified hierarchies table (no semantic/fuzzy matching).
    hierarchy_lookup = load_category_hierarchies()
    matched_fr  = rep_to_cat['catégorie_FR'].isin(hierarchy_lookup.keys()).sum()
    matched_int = rep_to_cat['catégorie_INT'].isin(hierarchy_lookup.keys()).sum()
    print(f"  ✓ Identified hierarchies lookup: {matched_fr}/{len(rep_to_cat)} Rep_Codes' "
          f"catégorie_FR found, {matched_int}/{len(rep_to_cat)} catégorie_INT found")

    mapper = rep_to_cat[["Rep_Code", "catégorie_FR", "catégorie_INT"]].copy()
    mapper["catégorie"] = mapper["catégorie_FR"]  # single display column, kept for pipeline.py

    cat_col_by_region = {REGION_FR: 'catégorie_FR', REGION_INTL: 'catégorie_INT'}
    for region, cat_col in cat_col_by_region.items():
        for seg_col in SEGMENT_LEVEL_COLS:
            out_col = f"{seg_col}_{region}"
            src_col = f"{seg_col}_FR"  # the hierarchies table only has one hierarchy per catégorie
            mapper[out_col] = mapper[cat_col].map(
                lambda c: hierarchy_lookup.get(c, {}).get(src_col, '---') if c else '---'
            )
            mapper[out_col] = mapper[out_col].apply(clean_segment_cell)

    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        mapper.to_excel(GENERATED_MAPPER_PATH, index=False)
        print(f"  ✓ Cached mapper → {GENERATED_MAPPER_PATH}")
    except Exception as e:
        print(f"  ⚠ Could not cache generated category mapper: {e}")

    print("=" * 60 + "\n")
    return mapper


if __name__ == '__main__':
    df = build_category_mapper(force_rebuild=True)
    print(df.head(20))
    print(f"Total: {len(df)} rows | with catégorie_FR: {(df['catégorie_FR'] != '').sum()} "
          f"| with catégorie_INT: {(df['catégorie_INT'] != '').sum()}")
    for region in (REGION_FR, REGION_INTL):
        col = f"Segment_{region}"
        print(f"  with {region} Segment: {(df[col] != '---').sum()}")