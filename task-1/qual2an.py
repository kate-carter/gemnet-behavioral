# Python file for processing raw Qualtrics output into analysis-ready files.
# this file was written partially  by Claude

from __future__ import annotations

import json
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font


INPUT_DIRECTORY = Path(
    "/path/to/gemnet-2/gemini-2/behavioral_task/"
)
INPUT_FILE = "gemnet_task-1_4-22-2.xlsx"
OUTPUT_SUBDIR = "concept_exports"

# Concepts are assumed to appear in this fixed order in both feature and subtype blocks.
CONCEPTS = [
    "cup",
    "pear",
    "truck",
    "spice",
    "cake",
    "flower",
    "lock",
    "spoon",
    "plastic",
    "computer",
    "bread",
    "bed",
    "potato",
    "chair",
    "stone",
]

BLOCK_SIZE = 15


def parse_import_id(cell: str) -> str:
    """Extract Qualtrics ImportId when metadata cell is JSON."""
    try:
        return json.loads(str(cell)).get("ImportId", "")
    except Exception:
        return ""


def extract_mm_dd(filename: str) -> str:
    """Extract MM-DD from filename; fallback to today's date."""
    match = re.search(r"(\d{1,2})-(\d{1,2})", filename)
    if match:
        month, day = int(match.group(1)), int(match.group(2))
        return f"{month:02d}-{day:02d}"
    return datetime.now().strftime("%m-%d")


def load_qualtrics_workbook(path: Path) -> tuple[pd.DataFrame, list[str], list[str], list[str]]:
    """Load metadata/header rows plus response rows from Qualtrics export."""
    meta = pd.read_excel(path, nrows=3, header=None, dtype=str).fillna("")
    raw_names = meta.iloc[0].tolist()
    labels = meta.iloc[1].tolist()
    import_meta = meta.iloc[2].tolist()

    data = pd.read_excel(path, skiprows=3, header=None, dtype=str).fillna("")
    data.columns = [f"col_{i}" for i in range(data.shape[1])]

    # Drop fully empty rows.
    data = data.loc[data.apply(lambda r: "".join(r.astype(str)).strip() != "", axis=1)].copy()

    # Keep only finished responses when available.
    if "Finished" in labels:
        finished_idx = labels.index("Finished")
        data = data[
            data[f"col_{finished_idx}"].astype(str).str.lower().isin(["true", "1", "yes"])
        ].copy()

    data = data.reset_index(drop=True)
    return data, raw_names, labels, import_meta


def build_participant_metadata(df: pd.DataFrame, labels: list[str]) -> pd.DataFrame:
    participant_labels = [f"Participant {i + 1}" for i in range(len(df))]
    participant_meta = pd.DataFrame({"participant": participant_labels})

    meta_cols = [
        "Response ID",
        "Start Date",
        "End Date",
        "Recorded Date",
        "Duration (in seconds)",
        "Finished",
        "Progress",
        "IP Address",
        "User Language",
    ]
    meta_idx = {name: labels.index(name) for name in meta_cols if name in labels}
    for name, idx in meta_idx.items():
        participant_meta[name] = df[f"col_{idx}"].tolist()

    if "Duration (in seconds)" in participant_meta.columns:
        participant_meta["Duration (minutes)"] = (
            pd.to_numeric(participant_meta["Duration (in seconds)"], errors="coerce") / 60.0
        )
    return participant_meta


def export_concept_files(
    src_path: Path,
    out_dir: Path,
    df: pd.DataFrame,
    raw_names: list[str],
    labels: list[str],
    import_meta: list[str],
) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    participant_labels = [f"Participant {i + 1}" for i in range(len(df))]
    participant_meta = build_participant_metadata(df, labels)

    feature_start = next(i for i, v in enumerate(labels) if str(v).strip() == "Feature 1")
    subtype_start = next(i for i, v in enumerate(labels) if str(v).strip() == "Sub type 1")

    feature_blocks_in_file = sum(1 for v in labels if str(v).strip() == "Feature 1")
    subtype_blocks_in_file = sum(1 for v in labels if str(v).strip() == "Sub type 1")

    for i, concept in enumerate(CONCEPTS):
        feat_idxs = list(range(feature_start + i * BLOCK_SIZE, feature_start + (i + 1) * BLOCK_SIZE))
        sub_idxs = list(range(subtype_start + i * BLOCK_SIZE, subtype_start + (i + 1) * BLOCK_SIZE))

        # Sheet 1: columns = participants, rows = features.
        feat_by_participant = pd.DataFrame(index=[f"Feature {j}" for j in range(1, BLOCK_SIZE + 1)])
        for p_idx, p_name in enumerate(participant_labels):
            feat_by_participant[p_name] = [
                df.iloc[p_idx, idx] if idx < df.shape[1] else "" for idx in feat_idxs
            ]

        # Sheet 2: columns = participants, rows = sub types.
        sub_by_participant = pd.DataFrame(index=[f"Sub type {j}" for j in range(1, BLOCK_SIZE + 1)])
        for p_idx, p_name in enumerate(participant_labels):
            sub_by_participant[p_name] = [
                df.iloc[p_idx, idx] if idx < df.shape[1] else "" for idx in sub_idxs
            ]

        # Sheet 3: columns = features, rows = all answers (participants).
        feat_all_answers = pd.DataFrame(
            {
                f"Feature {j + 1}": [
                    df.iloc[p_idx, feat_idxs[j]] if feat_idxs[j] < df.shape[1] else ""
                    for p_idx in range(len(df))
                ]
                for j in range(BLOCK_SIZE)
            },
            index=participant_labels,
        )
        feat_all_answers.index.name = "participant"

        # Sheet 4: columns = sub types, rows = all answers (participants).
        sub_all_answers = pd.DataFrame(
            {
                f"Sub type {j + 1}": [
                    df.iloc[p_idx, sub_idxs[j]] if sub_idxs[j] < df.shape[1] else ""
                    for p_idx in range(len(df))
                ]
                for j in range(BLOCK_SIZE)
            },
            index=participant_labels,
        )
        sub_all_answers.index.name = "participant"

        # Sheet 5: metadata.
        mapping_rows = []
        for j in range(BLOCK_SIZE):
            fidx = feat_idxs[j]
            sidx = sub_idxs[j]
            mapping_rows.append(
                {
                    "concept": concept,
                    "slot": j + 1,
                    "feature_column_index": fidx,
                    "feature_header_label": labels[fidx] if fidx < len(labels) else "",
                    "feature_variable": raw_names[fidx] if fidx < len(raw_names) else "",
                    "feature_import_id": parse_import_id(import_meta[fidx]) if fidx < len(import_meta) else "",
                    "subtype_column_index": sidx,
                    "subtype_header_label": labels[sidx] if sidx < len(labels) else "",
                    "subtype_variable": raw_names[sidx] if sidx < len(raw_names) else "",
                    "subtype_import_id": parse_import_id(import_meta[sidx]) if sidx < len(import_meta) else "",
                }
            )
        mapping_df = pd.DataFrame(mapping_rows)

        run_info = pd.DataFrame(
            [
                {"key": "source_file", "value": str(src_path)},
                {"key": "concept", "value": concept},
                {"key": "participant_count", "value": len(df)},
                {"key": "feature_blocks_in_file", "value": feature_blocks_in_file},
                {"key": "subtype_blocks_in_file", "value": subtype_blocks_in_file},
                {"key": "concept_blocks_exported", "value": len(CONCEPTS)},
                {
                    "key": "unused_feature_block_index_1_based",
                    "value": feature_blocks_in_file if feature_blocks_in_file > len(CONCEPTS) else "",
                },
                {
                    "key": "unused_subtype_block_index_1_based",
                    "value": subtype_blocks_in_file if subtype_blocks_in_file > len(CONCEPTS) else "",
                },
            ]
        )

        out_file = out_dir / f"{i + 1:02d}_{concept}.xlsx"
        with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
            feat_by_participant.to_excel(writer, sheet_name="1_features_by_participant")
            sub_by_participant.to_excel(writer, sheet_name="2_subtypes_by_participant")
            feat_all_answers.to_excel(writer, sheet_name="3_features_all_answers")
            sub_all_answers.to_excel(writer, sheet_name="4_subtypes_all_answers")
            run_info.to_excel(writer, index=False, sheet_name="5_metadata", startrow=0)
            participant_meta.to_excel(
                writer, index=False, sheet_name="5_metadata", startrow=len(run_info) + 2
            )
            mapping_df.to_excel(
                writer,
                index=False,
                sheet_name="5_metadata",
                startrow=len(run_info) + len(participant_meta) + 5,
            )


def build_compact_summary_table(
    df: pd.DataFrame,
    labels: list[str],
    block_label_prefix: str,
) -> tuple[pd.DataFrame, dict[str, list[int]]]:
    """
    Build stacked answers per concept for either Features or Sub types.
    Returns:
      - summary DataFrame (columns=concepts)
      - per-concept list of 0-based row indices to bold (first response per participant)
    """
    block_start = next(
        i for i, v in enumerate(labels) if str(v).strip() == f"{block_label_prefix} 1"
    )
    answers_by_concept: dict[str, list[str]] = {}
    bold_rows_by_concept: dict[str, list[int]] = {}

    for i, concept in enumerate(CONCEPTS):
        idxs = list(range(block_start + i * BLOCK_SIZE, block_start + (i + 1) * BLOCK_SIZE))
        values: list[str] = []
        bold_rows: list[int] = []
        for p_idx in range(len(df)):
            participant_start_len = len(values)
            for idx in idxs:
                value = df.iloc[p_idx, idx] if idx < df.shape[1] else ""
                # Preserve exact slot order (1..15) without dropping blanks.
                values.append(value)
            bold_rows.append(participant_start_len)
        answers_by_concept[concept] = values
        bold_rows_by_concept[concept] = bold_rows

    max_len = max((len(v) for v in answers_by_concept.values()), default=0)
    padded_answers = {
        concept: values + [""] * (max_len - len(values))
        for concept, values in answers_by_concept.items()
    }
    return pd.DataFrame(padded_answers), bold_rows_by_concept


def apply_bold_markers(ws, bold_rows_by_concept: dict[str, list[int]]) -> None:
    """Bold first response row for each participant in each concept column."""
    bold_font = Font(bold=True)
    for col_idx, concept in enumerate(CONCEPTS, start=1):
        for data_row_idx in bold_rows_by_concept[concept]:
            # Header is row 1, so data row 0 maps to Excel row 2.
            ws.cell(row=2 + data_row_idx, column=col_idx).font = bold_font


def export_summary_file(src_path: Path, df: pd.DataFrame, labels: list[str]) -> Path:
    """
    Create a single summary workbook:
    - columns: concepts
    - rows: all feature answers in participant order (direct copy)
    - first feature for each participant is bolded in each concept column
    """
    features_df, features_bold_rows = build_compact_summary_table(df, labels, "Feature")
    subtypes_df, subtypes_bold_rows = build_compact_summary_table(df, labels, "Sub type")
    mm_dd = extract_mm_dd(src_path.name)
    summary_path = src_path.parent / f"gemnet_task-1_{mm_dd}_summary.xlsx"

    with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
        features_df.to_excel(writer, index=False, sheet_name="features")
        subtypes_df.to_excel(writer, index=False, sheet_name="sub types")
        apply_bold_markers(writer.book["features"], features_bold_rows)
        apply_bold_markers(writer.book["sub types"], subtypes_bold_rows)

    return summary_path


def main() -> None:
    src_path = INPUT_DIRECTORY / INPUT_FILE
    out_dir = INPUT_DIRECTORY / OUTPUT_SUBDIR
    if not src_path.exists():
        raise FileNotFoundError(f"Input file not found: {src_path}")

    df, raw_names, labels, import_meta = load_qualtrics_workbook(src_path)
    export_concept_files(src_path, out_dir, df, raw_names, labels, import_meta)
    summary_path = export_summary_file(src_path, df, labels)
    print(f"Created {len(CONCEPTS)} concept files in: {out_dir}")
    print(f"Created summary file: {summary_path}")


if __name__ == "__main__":
    main()
