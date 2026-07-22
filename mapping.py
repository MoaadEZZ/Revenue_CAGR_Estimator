"""
Category mapping pipeline.

Combines two ideas that were prototyped separately:

  * test1.py — hierarchical matching that assigns a "catégorie" to every
    product row by falling back through pr_offer -> pr_offer_line ->
    pr_business_line, using the multi-sheet CAGR SVP workbook as the
    category source and "All products requested.xlsx" as the Rep_Code
    dimension table.

  * test2.py — semantic matching (TF-IDF char n-grams + Truncated SVD /
    LSA + cosine similarity) that pairs a free-text category label with
    its closest Market Sizing segment.

Combined, they produce a single generated mapper keyed at Rep_Code
granularity:

    Rep_Code | catégorie | ID

"catégorie" comes from the hierarchical match (test1 idea). "ID" is the
identifier column found in the Market Sizing workbook, chosen by
semantically matching each unique "catégorie" against the Market Sizing
segments (test2 idea) — this replaces the old, brittle exact-text join on
"Sub-segment 2".

The CAGR SVP workbook is now the primary source of truth for categories
(it replaces the old flat "CAGR Mapper.xlsx").
"""
import os
import re
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import TruncatedSVD
from sklearn.metrics.pairwise import cosine_similarity

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')

DATA_CAGR_SVP     = os.path.join(DATA_DIR, "CAGR SVP.xlsx")
DATA_MARKET_SIZE  = os.path.join(DATA_DIR, "Market Sizing - 3Q 2025.xlsx")
DATA_ALL_PRODUCTS = os.path.join(DATA_DIR, "All products requested.xlsx")

# Cache of the built mapper, so the (relatively expensive) semantic-matching
# step doesn't have to re-run on every request. Call
# build_category_mapper(force_rebuild=True) whenever CAGR SVP, All Products,
# or Market Size change.
GENERATED_MAPPER_PATH = os.path.join(DATA_DIR, "generated_category_mapper.xlsx")

MIN_SIMILARITY = 0.4     # below this cosine similarity, leave ID blank
EMBED_DIM      = 128     # LSA dimensionality

# Domain-specific vocabulary enhancement, straight from test2.py
DOMAIN_SYNONYMS = {
    'Cloud':        ['IaaS', 'PaaS', 'SaaS', 'Flexible', 'Infrastructure', 'Hosting'],
    'Managed':      ['Service', 'Management', 'Managed Services', 'Support'],
    'Connectivity': ['WAN', 'LAN', 'Network', 'Ethernet', 'Internet', 'Broadband'],
    'Mobile':       ['Traffic', 'Enterprise mobile', 'Devices', 'Wireless', 'Cellular'],
    'Security':     ['Cybersecurity', 'Protection', 'Consulting', 'Firewall', 'Defense'],
    'Data':         ['Analytics', 'Intelligence', 'Management', 'Governance', 'Storage'],
    'Digital':      ['Application', 'Integration', 'Transformation', 'Digitalization'],
    'Voice':        ['Telephony', 'VoIP', 'Communication', 'Calling'],
    'IoT':          ['Internet of Things', 'Connected', 'Sensors', 'M2M'],
}


# ─────────────────────────────────────────────
# Small helpers (from test1.py / test2.py)
# ─────────────────────────────────────────────

def _check_space_at_the_end(x):
    """Remove trailing space from string (test1.py)."""
    if isinstance(x, str) and x.endswith(' '):
        return x[:-1]
    return x


def _clean_text(s) -> str:
    """Normalize text for semantic matching (test2.py)."""
    if pd.isna(s) or s == '':
        return ''
    s = str(s).lower()
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _expand_with_synonyms(label: str) -> str:
    expanded = label
    for key, synonyms in DOMAIN_SYNONYMS.items():
        if key.lower() in label.lower():
            expanded += " " + " ".join(synonyms)
    return expanded


def _find_id_column(df: pd.DataFrame):
    """Locate the 'ID' column in the Market Sizing sheet (exact match first,
    then a lenient case-insensitive fallback)."""
    for col in df.columns:
        if str(col).strip() == 'ID':
            return col
    for col in df.columns:
        if str(col).strip().lower() == 'id':
            return col
    return None


# ─────────────────────────────────────────────
# Step 1 (test1.py idea): hierarchical Rep_Code → catégorie matching
# ─────────────────────────────────────────────

def load_cagr_svp_categories(path=DATA_CAGR_SVP) -> pd.DataFrame:
    """Load every relevant sheet of the CAGR SVP workbook (skipping MIF*/
    *GLOBAL* sheets) and concatenate into one dataframe of
    pr_business_line / pr_offer_line / pr_offer / catégorie."""
    sheets = pd.ExcelFile(path).sheet_names
    relevant_sheets = [s for s in sheets if not s.startswith("MIF") and "GLOBAL" not in s]

    df_list = []
    for sheet_name in relevant_sheets:
        df_temp = pd.read_excel(path, sheet_name=sheet_name, skiprows=2)
        if "catégorie" not in df_temp.columns:
            continue
        df_temp = df_temp.dropna(subset=["catégorie"])
        df_temp["catégorie"] = df_temp["catégorie"].apply(_check_space_at_the_end)
        df_list.append(df_temp)

    if not df_list:
        return pd.DataFrame(columns=["pr_business_line", "pr_offer_line", "pr_offer", "catégorie"])

    all_df = pd.concat(df_list, ignore_index=True)
    all_df["catégorie"] = all_df["catégorie"].fillna('')
    return all_df


def _find_category_hierarchical(row, mapper_df: pd.DataFrame) -> str:
    """3-level fallback match (test1.py's find_category_hierarchical):
    1. pr_offer, 2. pr_offer_line (where pr_offer is blank),
    3. pr_business_line (where pr_offer & pr_offer_line are blank)."""
    match = mapper_df[mapper_df["pr_offer"] == row.get("Offer")]
    if len(match) > 0:
        cats = [c for c in match["catégorie"].unique() if c != '']
        if cats:
            return cats[0]

    match = mapper_df[
        mapper_df["pr_offer"].isna() &
        (mapper_df["pr_offer_line"] == row.get("Offer_Line"))
    ]
    if len(match) > 0:
        cats = [c for c in match["catégorie"].unique() if c != '']
        if cats:
            return cats[0]

    match = mapper_df[
        mapper_df["pr_offer"].isna() &
        mapper_df["pr_offer_line"].isna() &
        (mapper_df["pr_business_line"] == row.get("Business_Line"))
    ]
    if len(match) > 0:
        cats = [c for c in match["catégorie"].unique() if c != '']
        if cats:
            return cats[0]

    return ''


def assign_categories_to_rep_codes(cagr_svp_path=DATA_CAGR_SVP,
                                    all_products_path=DATA_ALL_PRODUCTS) -> pd.DataFrame:
    """Rep_Code → catégorie, using All products requested.xlsx as the
    Rep_Code dimension table and the CAGR SVP workbook as the category
    source. Returns columns: Rep_Code, catégorie."""
    all_df_mapper = load_cagr_svp_categories(cagr_svp_path)

    df_dim_pro = pd.read_excel(all_products_path)
    df_dim_pro.columns = df_dim_pro.columns.str.strip()

    rep_col = next((c for c in df_dim_pro.columns if c.lower() == 'rep_code'), None)
    if rep_col is None:
        raise ValueError("All products requested.xlsx must have a 'Rep_Code' column")

    df_dim_pro["catégorie"] = df_dim_pro.apply(
        lambda row: _find_category_hierarchical(row, all_df_mapper), axis=1
    )

    out = df_dim_pro[[rep_col, "catégorie"]].rename(columns={rep_col: "Rep_Code"})
    out["catégorie"] = out["catégorie"].fillna('')
    out = out.drop_duplicates(subset=["Rep_Code"], keep='first')
    print(f"  ✓ Hierarchical match: {len(out)} Rep_Codes, "
          f"{(out['catégorie'] != '').sum()} assigned a catégorie")
    return out


# ─────────────────────────────────────────────
# Step 2 (test2.py idea): semantic matching of catégorie → market ID
# ─────────────────────────────────────────────

def load_market_segments(path=DATA_MARKET_SIZE) -> pd.DataFrame:
    """Load Market Sizing and return segments with Sub-segment 2."""
    df = pd.read_excel(path, sheet_name="DATA BASE MARKET FORECAST", skiprows=5)
    id_col = _find_id_column(df)
    if id_col is None:
        raise ValueError("Market Sizing file must contain an 'ID' column")

    subseg2_col = next((c for c in df.columns if c.strip() == "Sub-segment 2"), None)
    
    seg_cols = [c for c in ["Segment", "Sub-segment 1", "Sub-segment 2", "Sub-segment 3", "Region"]
                if c in df.columns]
    df["combined_segments"] = df[seg_cols].fillna("").astype(str).apply(
        lambda x: " ".join([v for v in x if v and v != "nan"]), axis=1
    )
    df = df.rename(columns={id_col: "ID"})
    df = df.dropna(subset=["ID"])
    df["ID"] = df["ID"].astype(str).str.rstrip()
    
    # NOTE: seg_cols is itself a list, so it must be *unpacked* into the
    # outer column list here (df[[...]] can't take a nested list as one of
    # its elements — pandas raises "unhashable type: 'list'" if you do,
    # which was silently crashing every rebuild and forcing a fallback to
    # a stale cached mapper).
    return df[["combined_segments", "ID", *seg_cols]].drop_duplicates(subset=["combined_segments"], keep='first')

def semantic_match_categories_to_ids(categories, market_df: pd.DataFrame,
                                      min_similarity=MIN_SIMILARITY,
                                      embed_dim=EMBED_DIM) -> dict:
    """For every unique catégorie label, find the closest Market Sizing
    'ID' using TF-IDF (char n-grams) + LSA + cosine similarity — the same
    technique as test2.py's semantic_match()."""
    if not categories or market_df.empty:
        return {c: '' for c in categories}

    market_labels = market_df["combined_segments"].tolist()
    market_ids    = market_df["ID"].tolist()

    cat_expanded    = [_expand_with_synonyms(c) for c in categories]
    market_expanded = [_expand_with_synonyms(m) for m in market_labels]

    cat_clean    = [_clean_text(c) for c in cat_expanded]
    market_clean = [_clean_text(m) for m in market_expanded]

    vectorizer = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=1, max_features=5000)
    all_texts = cat_clean + market_clean
    tfidf_all = vectorizer.fit_transform(all_texts)

    n_components = min(embed_dim, tfidf_all.shape[1] - 1, tfidf_all.shape[0] - 1)
    n_components = max(n_components, 2)
    svd = TruncatedSVD(n_components=n_components, random_state=42)
    vectors_all = svd.fit_transform(tfidf_all)

    cat_vectors    = vectors_all[:len(categories)]
    market_vectors = vectors_all[len(categories):]

    similarities = cosine_similarity(cat_vectors, market_vectors)

    category_to_id = {}
    for i, cat in enumerate(categories):
        sim_scores = similarities[i]
        best_idx   = int(np.argmax(sim_scores))
        best_score = float(sim_scores[best_idx])
        category_to_id[cat] = market_ids[best_idx] if best_score >= min_similarity else ''
    print(f"  ✓ Semantic match: {sum(1 for v in category_to_id.values() if v)}/"
          f"{len(category_to_id)} catégories matched to a market ID "
          f"(min similarity {min_similarity})")
    return category_to_id


# ─────────────────────────────────────────────
# Orchestration
# ─────────────────────────────────────────────

def build_category_mapper(force_rebuild=False) -> pd.DataFrame:
    """Build (or load the cached) Rep_Code / catégorie / ID mapper.

    Returns a dataframe with columns: Rep_Code, catégorie, ID.
    """
    if not force_rebuild and os.path.exists(GENERATED_MAPPER_PATH):
        try:
            cached = pd.read_excel(GENERATED_MAPPER_PATH)
            if {'Rep_Code', 'catégorie', 'ID'}.issubset(cached.columns):
                return cached
        except Exception:
            pass  # fall through and rebuild

    print("\n" + "=" * 60)
    print("BUILDING CATEGORY MAPPER (Rep_Code → catégorie → ID)")
    print("=" * 60)

    rep_to_cat = assign_categories_to_rep_codes()

    market_df = load_market_segments()
    unique_categories = sorted([c for c in rep_to_cat["catégorie"].unique() if c])
    cat_to_id = semantic_match_categories_to_ids(unique_categories, market_df)

    rep_to_cat["ID"] = rep_to_cat["catégorie"].map(cat_to_id).fillna('')
    mapper = rep_to_cat[["Rep_Code", "catégorie", "ID"]].copy()

    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        mapper.to_excel(GENERATED_MAPPER_PATH, index=False)
        print(f"  ✓ Cached mapper → {GENERATED_MAPPER_PATH}")
    except Exception as e:
        print(f"  ⚠ Could not cache generated category mapper: {e}")

    # Attach Segment / Sub-segment 1-3 (not just Sub-segment 2) so they're
    # selectable/displayable downstream, keyed off the same ID that the
    # semantic match already produced.
    for seg_col in ["Segment", "Sub-segment 1", "Sub-segment 2", "Sub-segment 3"]:
        if seg_col in market_df.columns:
            id_to_seg = dict(zip(market_df["ID"], market_df[seg_col]))
            mapper[seg_col] = mapper["ID"].map(id_to_seg).fillna('---')
    print("=" * 60 + "\n")
    return mapper


if __name__ == '__main__':
    df = build_category_mapper(force_rebuild=True)
    print(df.head(20))
    print(f"Total: {len(df)} rows | with catégorie: {(df['catégorie'] != '').sum()} | "
          f"with ID: {(df['ID'] != '').sum()}")
