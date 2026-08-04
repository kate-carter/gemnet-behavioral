#qualtrics behavioral processing - impossibility check

#imports
import pandas as pd
import os
import re
import openpyxl
import numpy as np

# CONFIG - edit this section
filepath = "/path/to/directory"
filename = "input.xlsx"
output_filename = "output.xlsx"

# set concept manually - must match a key in impossible_lists below
CONCEPT = "bed"

# IMPOSSIBLE LISTS - placeholder empty lists, will fill these in once lists are finalized

impossible_lists = {
    "bed": ['delicious', 'flavorful', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'handheld', 'causes pollution', 'used as an additive', 'contains eggs', 'has a straw', 'has buds', 'has candles', 'has exhaust', 'has eyes', 'has frosting', 'has icing', 'has headlights', 'has petals', 'has seeds', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
    "bread": ['electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a straw', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has large wheels', 'has lights', 'has memory foam', 'has mirrors', 'has or needs a key'],
    "cake": ['electrical', 'has four wheels', 'has a code', 'has a frame', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has blankets', 'has doors', 'has exhaust', 'has headlights', 'has memory foam', 'has or needs a key', 'has roots'],
    "chair": ['delicious', 'flavorful', 'edible', 'used as an additive', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a straw', 'has a trailer', 'has candles', 'has exhaust', 'has fillings', 'has headlights', 'has icing', 'has or needs a key', 'has seeds', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
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
row_1 = matrix[0, :]
row_2 = matrix[1, :]
subconcepts = [c for item in row_1 if (c := flow_concept(item)) is not None]
if not subconcepts:
    raise ValueError("No FL_*_DO_* flow-order columns found in this export.")

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


# Main processing loop

records = []

for row_idx in range(3, matrix.shape[0]):
    total_invalid = []
    participant_trials = []
    participant_row = matrix[row_idx, :]
    time_completed = pd.to_datetime(participant_row[1], errors="coerce")
    if pd.notna(time_completed):
        time_completed = time_completed.strftime("%-m/%-d/%y %-H:%M")
    else:
        time_completed = None

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

    #pull mturk id
    mturk_index = np.where(matrix[0, :] == 'Q139')[0][0]
    mturk_id = participant_row[mturk_index]

    # pull data for each concept (concept loop)
    for presentation_num, (order, concept) in enumerate(flow_order, start=1):
        block = subconcepts.index(concept)
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
            "concept": concept, 
            "sub_features": sub_features,
            "trial_invalid": trial_invalid,
            "complete": len(sub_features) >= 10,
            })

    #validity decision per participant
    if len(total_invalid) >= 3:
        valid = False
    else:
        valid = True


    for trial in participant_trials:
        record = {
            "participant number": row_idx - 2,
            "mturk_id": mturk_id,
            "time_completed": time_completed,
            "presentation_num": trial["presentation_num"],
            "concept": trial["concept"],
            "complete": trial["complete"],
            "valid": valid,  # False for all concepts if 3+ impossible features total per participant
            "invalid features": ", ".join(trial["trial_invalid"]),
        }

        # add feature columns at the end
        for i, feat in enumerate(trial["sub_features"], start=1):
            record[f"feat. {i}"] = feat

        records.append(record)

df = pd.DataFrame(records)
df.to_excel(os.path.join(filepath, output_filename), sheet_name="cleaned", index=False)

