#qualtrics behavioral processing - impossibility check

#imports
import pandas as pd
import os
import re
import openpyxl
import numpy as np

# CONFIG - edit this section

filepath = "/path/to/gemini-2/behavioral_task/task-2/qualtrics_exp/"
filename = "filename.csv"
output_filename = "filename-imp.xlsx"

# set concept manually - must match a key in impossible_lists below
CONCEPT = "cake"

# IMPOSSIBLE LISTS - placeholder empty lists, will fill these in once lists are finalized

impossible_lists = {
    "bed": ['creamy', 'crispy', 'delicious', 'flavorful', 'ceramic', 'easily melted', 'edible', 'farmed', 'fast', 'grows in soil', 'grows on trees', 'handheld', 'causes pollution', 'used as an additive', 'contains eggs', 'chocolate', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has a mouse', 'has a steering wheel', 'has a stem', 'has a straw', 'has a trailer', 'has buds', 'has candles', 'has doors', 'has exhaust', 'has eyes', 'has frosting', 'has icing', 'has headlights', 'has petals', 'has roots', 'has seeds', 'has skin', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
    "bread": ['clear', 'creaky', 'ceramic', 'easily melted', 'electrical', 'grows in soil', 'causes pollution', 'has four wheels', 'has a code', 'has a cushion', 'has a handle', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a stem', 'has a straw', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has blankets', 'has buds', 'has doors', 'has exhaust', 'has eyes', 'has four legs', 'has headlights', 'has large wheels', 'has lights', 'has memory foam', 'has mirrors', 'has or needs a key'],
    "cake": ['clear', 'creaky', 'ceramic', 'electrical', 'farmed', 'grows in soil', 'grows on trees', 'causes pollution', 'glass', 'has four wheels', 'has a code', 'has a cushion', 'has a handle', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has blankets', 'has doors', 'has exhaust', 'has four legs', 'has large wheels', 'has headlights', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows', 'has roots', 'has sheets', 'has skin', 'has tubers', 'has windows'],
    "chair": ['creamy', 'delicious', 'flavorful', 'edible', 'used as an additive', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a stem', 'has a straw', 'has a trailer', 'has buds', 'has candles', 'has crust', 'has exhaust', 'has eyes', 'has fillings', 'has headlights', 'has icing', 'has or needs a key', 'has seeds', 'has sheets', 'has sprinkles', 'has toppings', 'has tubers', 'baked', 'boiled', 'cooked', 'fried'],
    "computer": ['creamy', 'delicious', 'flavorful', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'chocolate', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has four wheels', 'has a mattress', 'has a seatbelt', 'has a straw', 'has a trailer', 'has blankets', 'has buds', 'has candles', 'has frosting', 'has icing', 'has large wheels', 'has leaves', 'has memory foam', 'has petals', 'has sheets', 'has skin', 'has sprinkles', 'has tubers', 'baked', 'boiled', 'fried'],
    "cup": ['crispy', 'crunchy', 'edible', 'used as an additive', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has four wheels', 'has a headboard', 'has a mattress', 'has a mouse', 'has a processor (CPU)', 'has a seatbelt', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has armrests', 'has blankets', 'has doors', 'has exhaust', 'has frosting', 'has headlights', 'has memory foam', 'has pillows', 'has sheets', 'has sprinkles', 'has tubers', 'fried'],
    "flower": ['easily melted', 'causes pollution', 'has a code', 'has a frame', 'has a headboard', 'has a keyboard', 'has a mattress', 'has a monitor', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has candles', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has lights', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows', 'has sheets', 'has windows'],
    "lock": ['creamy', 'crispy', 'crumbly', 'delicious', 'flavorful', 'edible', 'farmed', 'grows in soil', 'grows on trees', 'chocolate', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has a straw', 'has buds', 'has blankets', 'has arm rests', 'has candles', 'has frosting', 'has icing', 'has memory foam', 'has petals', 'has roots', 'has sprinkles', 'has tubers', 'fried', 'cooked', 'boiled'],
    "pear": ['ceramic', 'easily melted', 'electrical', 'foldable', 'has four wheels', 'has a code', 'has a cushion', 'has a frame', 'has a handle', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has doors', 'has exhaust', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows', 'has sheets', 'has windows'],
    "plastic": ['delicious', 'edible', 'ceramic', 'grows in soil', 'grows on trees', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has a processor (CPU)', 'has a trailer', 'has arm rests', 'has blankets', 'has buds', 'has candles', 'has exhaust', 'has four legs', 'has frosting', 'has icing', 'has mirrors', 'has sprinkles', 'has tubers'],
    "potato": ['ceramic', 'easily melted', 'electrical', 'chocolate', 'has four wheels', 'has a code', 'has a cushion', 'has a frame', 'has a handle', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has exhaust', 'has four legs', 'has large wheels', 'has memory foam', 'has mirrors', 'has or needs a key', 'has pillows', 'has sheets', 'has sprinkles'],
    "spice": ['ceramic', 'electrical', 'chocolate', 'contains gluten', 'contains flour', 'has four wheels', 'has a code', 'has a cushion', 'has a frame', 'has a handle', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has arms', 'has blankets', 'has exhaust', 'has eyes', 'has fillings', 'has four legs', 'has headlights', 'has large wheels', 'has memory foam', 'has mirrors', 'has numbers', 'has or needs a key', 'has sheets'],
    "spoon": ['creamy', 'chocolate', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has four wheels', 'has a code', 'has a cushion', 'has a headboard', 'has a hood', 'has a keyboard', 'has a keyhole', 'has a lid', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seat', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has buds', 'has candles', 'has doors', 'has exhaust', 'has four legs', 'has frosting', 'has headlights', 'has icing', 'has large wheels', 'has memory foam', 'has mirrors', 'has numbers', 'has or needs a key', 'has petals', 'has pillows', 'has roots', 'has seeds', 'has sheets', 'has skin', 'has sprinkles', 'has toppings', 'has tubers', 'has windows'],
    "stone": ['edible', 'electrical', 'farmed', 'grows in soil', 'grows on trees', 'chocolate', 'contains eggs', 'contains flour', 'contains gluten', 'contains grain', 'contains yeast', 'has four wheels', 'has a code', 'has a keyboard', 'has a keyhole', 'has a mattress', 'has a monitor', 'has a mouse', 'has a processor (CPU)', 'has a seatbelt', 'has a speaker', 'has a steering wheel', 'has a stem', 'has a straw', 'has a touchpad', 'has a touchscreen', 'has a trailer', 'has arm rests', 'has blankets', 'has buds', 'has candles', 'has four legs', 'has frosting', 'has headlights', 'has icing', 'has large wheels', 'has memory foam', 'has sprinkles'],
    "truck": ['creamy', 'contains gluten', 'contains yeast', 'has a mattress', 'has a straw', 'has buds', 'has candles', 'has frosting', 'has petals', 'has roots', 'has seeds', 'has skin', 'has sheets', 'has tubers', 'boiled', 'cooked'],
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
    """ Returns False (invalid) if more than 5 of the participant's selected
    features appear on the impossible list for this concept. """
    # normalize both lists to lowercase as a safety net in case of
    # capitalization mismatches between the survey and impossible list
    impossible = impossible_lists.get(concept.lower(), [])
    if not impossible:
        # no impossible list yet - default to valid
        return True  
    impossible_cleaned = []
    for feature in impossible:
        impossible_cleaned.append(feature.strip().lower())

    impossible_count = 0
    for feature in sub_features:
        if feature.strip().lower() in impossible_cleaned:
            impossible_count += 1
    # True = valid, False = invalid
    return impossible_count <= 5  

# Main processing loop

records = []

for row_idx in range(3, matrix.shape[0]):
    participant_row = matrix[row_idx, :]
    participant_id = participant_row[8]  # ResponseId
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

    # pull data for each concept
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

        record = {
            "participant number": row_idx - 2,
            "participant": participant_id,
            "time_completed": time_completed,
            "presentation_num": presentation_num,
            "concept": concept,
            "complete": len(sub_features) >= 10,
            "valid": check_impossible(sub_features, CONCEPT),  # False if > 5 impossible features
        }

        # add feature columns at the end
        for i, feat in enumerate(sub_features, start=1):
            record[f"feat. {i}"] = feat

        records.append(record)

# Output

#create overall report for each subconcept

#count how many times each subconcept was seen

#count how many times each feature was picked for each subconcept

df = pd.DataFrame(records)
df.to_excel(os.path.join(filepath, output_filename), sheet_name="cleaned", index=False)
