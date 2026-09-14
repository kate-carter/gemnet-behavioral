#analysis script for task 2

#qualtrics behavioral processing - impossibility check

#imports
import pandas as pd
import os
import re
import openpyxl
import numpy as np

# CONFIG - edit this section
filepath = "/Users/Kate/Documents/Binghamton-SPiNS/word-embedding-project/gemnet-2/gemini-2/behavioral_task/task-2/"
filename = "qualtrics_exp/Truck_8-4.xlsx"
output_filename = "task-2-processed/truck_master.xlsx"
output_path = os.path.join(filepath, output_filename)

# set concept manually - must match a key in impossible_lists below
CONCEPT = "truck"

# SUBTYPE ALIASES — map older/alternate Qualtrics FL_*_DO_* tags onto the
# canonical name used in later exports. Counts and meta keys use the canonical
# name so early and late files merge instead of creating duplicate zero columns.
# Add new entries as: "old_or_alternate_tag": "canonical_tag"
SUBTYPE_ALIASES = {
    #bed

    # bread (early exports dropped the trailing "bread")
    "whole-grain": "whole-grainbread",
    "glutenfree": "glutenfreebread",
    "rye": "ryebread",
    "sliced": "slicedbread",
    # potato
    "idaho": "idahopotato",
}


def canonical_subtype(name):
    """Return the canonical subtype label for counting / merging."""
    if name is None:
        return None
    return SUBTYPE_ALIASES.get(str(name), str(name))


def consolidate_subtype_counts(counts_dict):
    """Sum counts that share the same canonical subtype name."""
    merged = {}
    for name, n in counts_dict.items():
        key = canonical_subtype(name)
        merged[key] = merged.get(key, 0) + int(n)
    return merged


def consolidate_counts_df(df):
    """Rename alias columns to canonical names and sum duplicates."""
    if df is None or df.empty:
        return df
    renamed = df.rename(columns={c: canonical_subtype(c) for c in df.columns})
    return renamed.T.groupby(level=0).sum().T


# IMPOSSIBLE LISTS - placeholder empty lists, will fill these in once lists are finalized

impossible_lists = {
    "bed": ['delicious', 'flavorful', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'causes pollution', 'used as an additive', 'contains eggs', 'contains gluten','has a straw', 'has buds', 'has candles', 'has exhaust', 'has eyes', 'has frosting', 'has icing', 'has headlights', 'has petals', 'has seeds', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
    "bread": ['electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a straw', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has large wheels', 'has lights', 'has memory foam', 'has mirrors', 'has or needs a key'],
    "cake": ['electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has blankets', 'has doors', 'has exhaust', 'has headlights', 'has memory foam', 'has or needs a key', 'has roots'],
    "chair": ['delicious', 'flavorful', 'edible', 'used as an additive', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a straw', 'has a trailer', 'has candles', 'has exhaust', 'has headlights', 'has icing', 'has or needs a key', 'has seeds', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
    "computer": ['delicious', 'flavorful', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has a mattress', 'has a seatbelt', 'has a straw', 'has a trailer', 'has blankets', 'has candles', 'has frosting', 'has icing', 'has leaves', 'has memory foam', 'has petals', 'has sheets', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'fried'],
    "cup": ['used as an additive', 'has four wheels', 'has a headboard', 'has a mattress', 'has a mouse', 'has a processor (CPU)', 'has a seatbelt', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has blankets', 'has doors', 'has exhaust', 'has headlights', 'has memory foam', 'has pillows', 'has sheets', 'has tubers'],
    "flower": ['easily melted', 'has a code', 'has a frame', 'has a headboard', 'has a keyboard', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has doors', 'has headlights', 'has memory foam', 'has mirrors', 'has or needs a key'],
    "lock": ['edible', 'farmed', 'grows in soil', 'grows on trees', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has a straw', 'has blankets', 'has arm rests', 'has candles', 'has frosting', 'has icing', 'has memory foam', 'has petals', 'has roots', 'has sprinkles', 'has tubers', 'fried', 'cooked', 'boiled'],
    "pear": ['electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows', 'has windows'],
    "plastic": ['ceramic', 'has a processor (CPU)', 'has a trailer', 'has arm rests', 'has blankets', 'delicious', 'flavorful', 'ancient', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'has a seatbelt'],
    "potato": ['has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has exhaust', 'has four legs', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows'],
    "spice": ['ceramic', 'electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has eyes', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key'],
    "spoon": ['has four wheels', 'has a code', 'has a cushion', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key', 'has sheets', 'has windows'],
    "stone": ['contains eggs', 'contains flour', 'contains gluten', 'contains yeast', 'has four wheels', 'has a code', 'has a keyboard', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a straw', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has candles', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam'],
    "truck": ['boiled', 'cooked', 'grows on trees'],
}

feature_list = [
    "airy", "beige", "black", "bland", "bright", "brown", "clear", "cold",
    "colorful", "creaky", "creamy", "crispy", "crumbly", "crunchy", "delicious",
    "dense", "earthy", "flat on bottom", "concave", "curved", "cylindrical",
    "flavorful", "flexible", "fluffy", "fragrant", "grainy", "gray", "green",
    "hard", "heavy", "accessible", "adjustable", "ancient", "bendable", "brittle",
    "celebratory", "cheap", "ceramic", "comfortable", "complex",
    "convenient", "crushed", "cultural", "decorated", "decorative", "delicate",
    "disposable", "durable", "easily melted", "edible", "electrical", "expensive",
    "fake", "farmed", "fast", "foldable", "grows in soil", "grows on trees",
    "handheld", "harmful", "causes pollution", "used as an additive", "chocolate",
    "contains eggs", "contains flour", "contains gluten", "contains grain",
    "contains yeast", "glass", "golden", "has four wheels", "has a code",
    "has a cushion", "has a frame", "has a handle", "has a headboard", "has a hood",
    "has a keyboard", "has a keyhole", "has a lid", "has a mattress", "has a monitor",
    "has a mouse", "has a processor (CPU)", "has a scent", "has a seat",
    "has a seatbelt", "has a speaker", "has a steering wheel", "has a stem",
    "has a straw", "has a touchpad", "has a touchscreen", "has a trailer",
    "has arm rests", "has arms", "has blankets", "has buds", "has candles",
    "has crust", "has doors", "has exhaust", "has eyes", "has fillings",
    "has four legs", "has frosting", "has headlights", "has icing",
    "has large wheels", "has layers", "has leaves", "has legs", "has lights",
    "has memory foam", "has mirrors", "has numbers", "has or needs a key",
    "has petals", "has pillows", "has roots", "has seeds", "has sheets", "has skin",
    "has sprinkles", "has toppings", "has tubers", "has windows", "baked", "boiled",
    "can be closed", "clean", "cooked", "cut", "dirty", "fresh", "fried",
]

#load data from qualtrics and convert to matrix
path = os.path.join(filepath, filename)
if filename.lower().endswith(".csv"):
    data = pd.read_csv(path, header=None, low_memory=False)
else:
    data = pd.read_excel(path, header=None)
matrix = data.to_numpy()

# collect header info
header_info = matrix[0:17, 0:3]

# Flow-order columns look like FL_3_DO_gemstone / FL_4_DO_twinbed (number varies by survey).
FLOW_DO_RE = re.compile(r"^FL_\d+_DO_(.+)$")

def flow_concept(item):
    if not isinstance(item, str):
        return None
    match = FLOW_DO_RE.match(item)
    return match.group(1) if match else None

#pull last n items of column one to create subconcept index
# raw names must stay in FL/block order for indexing into block_starts
row_1 = matrix[0, :]
row_2 = matrix[1, :]
subconcepts_raw = [c for item in row_1 if (c := flow_concept(item)) is not None]
if not subconcepts_raw:
    raise ValueError("No FL_*_DO_* flow-order columns found in this export.")

# canonical subtype columns for the count matrix (unique, order preserved)
subconcepts = []
for c in subconcepts_raw:
    cc = canonical_subtype(c)
    if cc not in subconcepts:
        subconcepts.append(cc)
# Feature blocks start at each concept's Q2_1 column (first block may
# include a leading "Required"/Q2 consent column before that).
fl_start = next(j for j, item in enumerate(row_1) if flow_concept(item) is not None)
q2_1_cols = [j for j in range(17, fl_start) if row_1[j] == "Q2_1"]
if row_1[17] == "Q2":
    block_starts = [17] + q2_1_cols[1:]
else:
    block_starts = q2_1_cols
block_ends = block_starts[1:] + [fl_start]

# Feature labels aligned to columns within the first block (same order
# across concepts). Prefer reading the header at each selected column
# so we never index features by raw block position.


def is_selected(value):
    """Qualtrics numeric export uses 1; CSV often stores it as '1'."""
    return value in (1, 1.0, "1") or (isinstance(value, str) and value.strip() == "1")

def feature_name(header):
    return header.split("- ")[-1].strip()

# Impossibility check


def check_impossible(sub_features, concept):
    """ Returns list of invalid features if more than 3 of the participant's selected
    features appear on the impossible list across all concepts. """
    # normalize both lists to lowercase as a safety net in case of
    # capitalization mismatches between the survey and impossible list
    impossible = impossible_lists.get(concept.lower(), [])
    if not impossible:
        # no impossible list yet - default to valid
        return []
    impossible_cleaned = []
    for feature in impossible:
        impossible_cleaned.append(feature.strip().lower())

    trial_invalid = []
    for feature in sub_features:
        if feature.strip().lower() in impossible_cleaned:
            trial_invalid.append(feature)
    # True = valid, False = invalid
    return trial_invalid




# resolve MTurk ID column once (same logic as task-2-imp.py)
MTURK_PROMPTS = (
    "Please enter your MTurk ID and then press the button to begin.",
    "Please enter your MTurk ID and press the button to begin",
)
mturk_index = None
for prompt in MTURK_PROMPTS:
    if prompt in matrix[1, :]:
        mturk_index = np.where(matrix[1, :] == prompt)[0][0]
        break
if mturk_index is None:
    if CONCEPT == "cake":
        mturk_index = np.where(matrix[0, :] == "Q151")[0][0]
    elif CONCEPT in ("computer", "pear", "spoon"):
        mturk_index = np.where(matrix[0, :] == "Q119")[0][0]
    elif CONCEPT == "cup":
        mturk_index = np.where(matrix[0, :] == "Q131")[0][0]
    elif CONCEPT in ("flower", "lock", "plastic", "spice"):
        mturk_index = np.where(matrix[0, :] == "Q149")[0][0]
    elif CONCEPT == "potato":
        mturk_index = np.where(matrix[0, :] == "Q126")[0][0]
    else:
        mturk_index = np.where(matrix[0, :] == "Q139")[0][0]

# ResponseId fallback when MTurk ID is blank (common in older exports)
response_id_matches = np.where(matrix[0, :] == "ResponseId")[0]
if len(response_id_matches) == 0:
    raise ValueError("No ResponseId column found in this export.")
response_id_index = int(response_id_matches[0])


def build_meta(n_valid_total, source_files, mturk_records, subtype_counts):
    """Long-format meta sheet: totals, source files, mturk ids (+ date), per-subtype Ns."""
    rows = [{"field": "n_valid_total", "value": n_valid_total, "date": None}]
    for f in source_files:
        rows.append({"field": "source_file", "value": f, "date": None})
    for rec in mturk_records:
        rows.append({
            "field": "mturk_id",
            "value": rec["id"],
            "date": rec.get("date"),
        })
    for subtype, n in subtype_counts.items():
        rows.append({"field": f"subtype_n:{subtype}", "value": int(n), "date": None})
    return pd.DataFrame(rows)


def load_meta(path):
    """Return (source_files, mturk_ids, subtype_counts) from an existing workbook."""
    meta = pd.read_excel(path, sheet_name="meta")
    source_files = (
        meta.loc[meta["field"] == "source_file", "value"].astype(str).tolist()
    )
    mturk_records = []
    mturk_rows = meta.loc[meta["field"] == "mturk_id"]
    has_date = "date" in meta.columns
    for _, row in mturk_rows.iterrows():
        rec = {"id": str(row["value"])}
        if has_date and pd.notna(row.get("date")):
            rec["date"] = str(row["date"])
        else:
            rec["date"] = None
        mturk_records.append(rec)
    subtype_counts = {}
    subtype_rows = meta[meta["field"].astype(str).str.startswith("subtype_n:")]
    for _, row in subtype_rows.iterrows():
        subtype = canonical_subtype(str(row["field"]).split(":", 1)[1])
        subtype_counts[subtype] = subtype_counts.get(subtype, 0) + int(row["value"])
    return source_files, mturk_records, subtype_counts


# load prior meta so repeat workers can be skipped before counting
if os.path.exists(output_path):
    try:
        source_files, mturk_records, subtype_counts = load_meta(output_path)
    except (ValueError, KeyError, IndexError):
        source_files, mturk_records, subtype_counts = [], [], {}
else:
    source_files, mturk_records, subtype_counts = [], [], {}

subtype_counts = consolidate_subtype_counts(subtype_counts)

seen_mturk_ids = {rec["id"] for rec in mturk_records}

# ensure every subtype in this file has a counter (new subtypes start at 0)
for subtype in subconcepts:
    subtype_counts.setdefault(subtype, 0)

#initialize records
records = []
mturk_records_this_run = []

#initialize analysis dataframe
counts = pd.DataFrame(0, columns=subconcepts, index=feature_list)


# Main processing loop
for row_idx in range(3, matrix.shape[0]):
    total_invalid = []
    participant_trials = []
    participant_row = matrix[row_idx, :]
    end_date = pd.to_datetime(participant_row[1], errors="coerce")
    if pd.notna(end_date):
        time_completed = end_date.strftime("%-m/%-d/%y %-H:%M")
        id_date = end_date.strftime("%Y-%m-%d")
    else:
        time_completed = None
        id_date = None

    mturk_id = participant_row[mturk_index]
    if pd.isna(mturk_id):
        mturk_id = None
    else:
        mturk_id = str(mturk_id).strip() or None

    # older exports often leave MTurk blank — fall back to Qualtrics ResponseId
    if not mturk_id:
        response_id = participant_row[response_id_index]
        if pd.isna(response_id):
            mturk_id = None
        else:
            mturk_id = str(response_id).strip() or None

    # use flow order to determine which were seen
    flow_order = []

    for j, item in enumerate(row_1):
        concept = flow_concept(item)
        if concept is None:
            continue
        order = participant_row[j]
        if pd.notna(order) and str(order).strip() != "":
            flow_order.append((int(float(order)), concept))
    flow_order.sort(key=lambda x: x[0])


    # pull data for each concept (concept loop)
    for presentation_num, (order, concept) in enumerate(flow_order, start=1):
        block = subconcepts_raw.index(concept)
        block_start = block_starts[block]
        block_end = block_ends[block]
        block_headers = row_2[block_start:block_end]
        block_data = participant_row[block_start:block_end]

        sub_features = [
            feature_name(header)
            for header, value in zip(block_headers, block_data)
            if isinstance(header, str) and "Please" in header and is_selected(value)
        ]

        trial_invalid = check_impossible(sub_features, CONCEPT)
        for feature in trial_invalid:
            total_invalid.append(feature)

        participant_trials.append({
            "presentation_num": presentation_num,
            "concept": canonical_subtype(concept),
            "sub_features": sub_features,
            "trial_invalid": trial_invalid,
            "complete": len(sub_features) >= 10,
            })

    # skip invalid participants (3+ impossible features total)
    if len(total_invalid) >= 3:
        continue

    # need an ID (MTurk or ResponseId) to dedupe across files / within this run
    if not mturk_id:
        continue

    # skip workers already recorded in this master file (or earlier in this run)
    if mturk_id in seen_mturk_ids:
        continue

    # "used" = valid, new ID, and at least one complete trial enters the matrix
    if not any(trial["complete"] for trial in participant_trials):
        continue

    seen_mturk_ids.add(mturk_id)
    mturk_records_this_run.append({"id": mturk_id, "date": id_date})

    for trial in participant_trials:
        record = {
            "presentation_num": trial["presentation_num"],
            "concept": trial["concept"],
            "complete": trial["complete"],
            "sub_features": trial["sub_features"],
            "mturk_id": mturk_id,
            "date": id_date,
        }

        # add feature columns at the end
        for i, feat in enumerate(trial["sub_features"], start=1):
            record[f"feat. {i}"] = feat

        records.append(record)

# count features once, after all new valid participants are collected
for participant in records:
    if not participant["complete"]:
        continue
    concept = participant["concept"]
    subtype_counts[concept] = subtype_counts.get(concept, 0) + 1
    for feature in feature_list:
        if feature in participant["sub_features"]:
            counts.loc[feature, concept] += 1

# merge counts / meta with existing workbook if present
if os.path.exists(output_path):
    existing = pd.read_excel(output_path, sheet_name="cleaned", index_col=0)
    existing = consolidate_counts_df(existing)
    counts = existing.add(counts, fill_value=0)

counts = consolidate_counts_df(counts)

if filename not in source_files:
    source_files.append(filename)
mturk_records.extend(mturk_records_this_run)
n_valid_total = len(mturk_records)

counts = counts.fillna(0)
meta = build_meta(n_valid_total, source_files, mturk_records, subtype_counts)

# proportions: feature count / N participants who contributed that subtype
proportions = counts.astype(float).copy()
for col in proportions.columns:
    n = int(subtype_counts.get(col, 0))
    if n == 0:
        # no valid complete trials for this subtype (incl. 0/0) → leave zeros
        proportions[col] = 0.0
    else:
        proportions[col] = proportions[col] / n

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    counts.to_excel(writer, sheet_name="cleaned", index=True)
    proportions.to_excel(writer, sheet_name="proportions", index=True)
    meta.to_excel(writer, sheet_name="meta", index=False)
