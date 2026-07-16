#qualtrics behavioral processing


#imports
import pandas as pd
import os
import openpyxl
import numpy as np


#data
filepath = "path/to/gemnet-2/gemini-2/behavioral_task/task-2/qualtrics_exp/"
filename = "Stone_7-16.csv"  # will eventually change to be formulaic
output_filename = "qualtrics_task-2-Stone-7-16-cleaned.xlsx"

#load data from qualtrics and convert to matrix
path = os.path.join(filepath, filename)
if filename.lower().endswith(".csv"):
    data = pd.read_csv(path, header=None, low_memory=False)
else:
    data = pd.read_excel(path, header=None)
matrix = data.to_numpy()

#critical feature: SET PER CONCEPT
critical_feature = "heavy"

#collect header info
header_info = matrix[0:17, 0:3]

#pull last n items of column one to create subconcept index
row_1 = matrix[0, :]
row_2 = matrix[1, :]
subconcepts = [
    item.replace("FL_3_DO_", "")
    for item in row_1
    if isinstance(item, str) and item.startswith("FL_3_DO_")
]

# Feature blocks start at each concept's Q2_1 column (first block may
# include a leading "Required"/Q2 consent column before that).
fl_start = next(
    j for j, item in enumerate(row_1)
    if isinstance(item, str) and item.startswith("FL_3_DO_")
)
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
    return header.split("- ")[-1]


records = []

for row_idx in range(3, matrix.shape[0]):
    participant_row = matrix[row_idx, :]
    participant_id = participant_row[8]  # ResponseId
    time_completed = participant_row[1]

    # use flow order to determine which were seen
    flow_order = []
    for j, item in enumerate(row_1):
        if isinstance(item, str) and item.startswith("FL_3_DO_"):
            concept = item.replace("FL_3_DO_", "")
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

        # start with fixed metadata columns
        record = {
            "participant number": row_idx - 2,
            "participant": participant_id,
            "time_completed": time_completed,
            "presentation_num": presentation_num,
            "concept": concept,
            "complete": len(sub_features) >= 10,
            "critical": critical_feature in sub_features,
        }

        # add feature columns at the end
        for i, feat in enumerate(sub_features, start=1):
            record[f"feat. {i}"] = feat

        records.append(record)

#create overall report for each subconcept

#count how many times each subconcept was seen

#count how many times each feature was picked for each subconcept


df = pd.DataFrame(records)
df.to_excel(os.path.join(filepath, output_filename), sheet_name="cleaned", index=False)
