# Correlates human behavioral master matrices (feature × subtype counts)
# instead of Gemini embeddings. Each feature is a vector over subtypes;
# the output is a feature × feature similarity/correlation matrix.

import numpy as np
import pandas as pd
import time
from scipy.spatial.distance import pdist, squareform
import argparse
import logging
import os

# Concepts that have (or will have) a {concept}_master.xlsx in task-2-screened/
concept_name = [
    "bed",
    "bread",
    "cake",
    "chair",
    "computer",
    "cup",
    "flower",
    "lock",
    "pear",
    "plastic",
    "potato",
    "spice",
    "spoon",
    "stone",
    "truck",
]

# Base paths
input_dir = (
    "/Users/Kate/Documents/Binghamton-SPiNS/word-embedding-project/"
    "gemnet-2/gemini-2/behavioral_task/task-2/task-2-processed/"
)
output_dir = (
    "/Users/Kate/Documents/Binghamton-SPiNS/word-embedding-project/"
    "gemnet-2/gemini-2/behavioral_task/task-2/corr_h/"
)

# Which sheet to use from the master workbook ("cleaned" = raw counts;
# use "percentage" if you have a proportion sheet instead).
sheet_name = "proportions"

os.makedirs(output_dir, exist_ok=True)

start_time = time.time()
parser = argparse.ArgumentParser()
parser.add_argument(
    "--method",
    type=str,
    default="cosine",
    help="Method: pearson, spearman, mahalanobis, cosine",
)
args = parser.parse_args()
method = args.method

results = []

for concept in concept_name:
    file_start = time.time()
    filename = f"{concept}_master.xlsx"
    input_path = os.path.join(input_dir, filename)
    print(f"Processing {concept} (human master)")

    if not os.path.exists(input_path):
        logging.warning(f"  Skipping {concept}: {filename} not found")
        continue

    try:
        # rows = features, columns = subtypes
        matrix_counts = pd.read_excel(input_path, sheet_name=sheet_name, index_col=0)
        matrix_counts = matrix_counts.apply(pd.to_numeric, errors="coerce").fillna(0)

        if matrix_counts.shape[0] == 0 or matrix_counts.shape[1] == 0:
            logging.warning(f"  Skipping {concept}: empty {sheet_name} sheet")
            continue

        # Drop all-zero features (no profile across subtypes)
        nonzero = matrix_counts.loc[(matrix_counts != 0).any(axis=1)]
        if nonzero.empty:
            logging.warning(f"  Skipping {concept}: all feature counts are zero")
            continue

        # features × subtypes — one row per feature (item to correlate)
        matrix_emb = nonzero
        feature_labels = matrix_emb.index.astype(str).tolist()
        print(f"  Loaded shape: {matrix_counts.shape} (features × subtypes)")
        print(
            f"  Correlating {len(feature_labels)} features over "
            f"{matrix_emb.shape[1]} subtypes"
        )

        X = matrix_emb.values

        if method == "pearson":
            # .T.corr correlates the original rows (= features)
            matrix_corr = matrix_emb.T.corr(method="pearson").values
            print(f"  Correlation matrix shape: {matrix_corr.shape}")
        elif method == "spearman":
            matrix_corr = matrix_emb.T.corr(method="spearman").values
            print(f"  Correlation matrix shape: {matrix_corr.shape}")
        elif method == "mahalanobis":
            VI = np.linalg.pinv(np.cov(X, rowvar=False))
            dists = pdist(X, metric="mahalanobis", VI=VI)
            d = squareform(dists)
            off_diag_distances = d[np.triu_indices_from(d, k=1)]

            sigma = np.median(off_diag_distances)
            if sigma > 0:
                sim01 = np.exp(-(d ** 2) / (2 * sigma ** 2))
            else:
                sim01 = np.ones_like(d)
            matrix_corr = sim01
            print(f"  Mahalanobis similarity matrix shape: {matrix_corr.shape}")
        elif method == "cosine":
            dists = pdist(X, metric="cosine")
            matrix_corr = 1 - squareform(dists)
            print(f"  Cosine similarity matrix shape: {matrix_corr.shape}")
        else:
            raise ValueError(f"Unknown method: {method}")

        # Min/max from off-diagonal only (diagonal must not set the scale)
        iu = np.triu_indices_from(matrix_corr, k=1)
        off_diag = matrix_corr[iu]
        min_val = np.min(off_diag)
        max_val = np.max(off_diag)
        print(f"  min/max before normalization: {min_val}, {max_val}")

        # Normalize to [-1, 1]
        if max_val > min_val:
            matrix_corr = 2 * ((matrix_corr - min_val) / (max_val - min_val)) - 1
        else:
            matrix_corr = np.zeros_like(matrix_corr, dtype=float)

        iu = np.triu_indices_from(matrix_corr, k=1)
        off_diag2 = matrix_corr[iu]
        print(
            f"  min/max after normalization: {np.min(off_diag2)}, {np.max(off_diag2)}"
        )

        # Convert diagonal from 1's to 0's
        matrix_corr = matrix_corr.copy()
        np.fill_diagonal(matrix_corr, 0)

        # Save with feature labels for interpretability
        output_file = os.path.join(output_dir, f"{concept}_human_{method}.csv")
        pd.DataFrame(
            matrix_corr, index=feature_labels, columns=feature_labels
        ).to_csv(output_file)

        elapsed = time.time() - file_start
        print(f"  ✓ Saved: {concept}_human_{method}.csv ({elapsed:.2f}s)")
        results.append({
            "concept": concept,
            "method": method,
            "matrix_shape": matrix_corr.shape,
        })
    except Exception as e:
        logging.error(f"ERROR processing {concept}: {e}")
        continue

total_time = time.time() - start_time
print("\n" + "=" * 60)
print(
    f"Completed! Processed {len(results)} / {len(concept_name)} concepts "
    f"in {total_time / 60:.2f} minutes"
)
print(f"Output directory: {output_dir}")
