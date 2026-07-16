"""Create incremental summary outputs for newly added Qualtrics participants.

Usage:
  python qual_update.py --old /path/to/previous.xlsx --new /path/to/latest.xlsx
"""
# this file was written partially by Claude

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font

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

# Static directory for behavioral task exports.
DATA_DIR = Path(
    "/path/to/gemnet-2/gemini-2/behavioral_task/"
)

# Default filenames for quick swapping between exports.
# You can edit these directly, or override with --old / --new at runtime.
OLD_FILE = "gemnet_task-1_4-27.xlsx"
NEW_FILE = "gemnet_task-1_4-29.xlsx"


def extract_mm_dd(filename: str) -> str:
    match = re.search(r"(\d{1,2})-(\d{1,2})", filename)
    if match:
        month, day = int(match.group(1)), int(match.group(2))
        return f"{month:02d}-{day:02d}"
    return datetime.now().strftime("%m-%d")


def load_qualtrics_data(path: Path) -> tuple[pd.DataFrame, list[str]]:
    """Load Qualtrics workbook data rows and label row."""
    meta = pd.read_excel(path, nrows=3, header=None, dtype=str).fillna("")
    labels = meta.iloc[1].tolist()

    data = pd.read_excel(path, skiprows=3, header=None, dtype=str).fillna("")
    data.columns = [f"col_{i}" for i in range(data.shape[1])]
    data = data.loc[data.apply(lambda r: "".join(r.astype(str)).strip() != "", axis=1)].copy()

    if "Finished" in labels:
        finished_idx = labels.index("Finished")
        data = data[
            data[f"col_{finished_idx}"].astype(str).str.lower().isin(["true", "1", "yes"])
        ].copy()

    data = data.reset_index(drop=True)
    return data, labels


def row_signature(row: pd.Series) -> str:
    return "|".join(row.astype(str).tolist())


def get_new_participants(
    old_df: pd.DataFrame, new_df: pd.DataFrame, old_labels: list[str], new_labels: list[str]
) -> pd.DataFrame:
    """Return only newly added participants from new_df."""
    if "Response ID" in old_labels and "Response ID" in new_labels:
        old_rid_idx = old_labels.index("Response ID")
        new_rid_idx = new_labels.index("Response ID")
        old_key_col = f"col_{old_rid_idx}"
        new_key_col = f"col_{new_rid_idx}"
        old_ids = set(old_df[old_key_col].astype(str))
        mask = ~new_df[new_key_col].astype(str).isin(old_ids)
        return new_df.loc[mask].copy().reset_index(drop=True)

    # Fallback: compare full-row signatures if Response ID is unavailable.
    shared_col_count = min(old_df.shape[1], new_df.shape[1])
    shared_cols = [f"col_{i}" for i in range(shared_col_count)]
    old_sigs = set(old_df[shared_cols].apply(row_signature, axis=1))
    new_sigs = new_df[shared_cols].apply(row_signature, axis=1)
    return new_df.loc[~new_sigs.isin(old_sigs)].copy().reset_index(drop=True)


def build_compact_summary_table(
    df: pd.DataFrame, labels: list[str], block_label_prefix: str
) -> tuple[pd.DataFrame, dict[str, list[int]]]:
    """Build strict 15-slot stacked table per concept; preserve blanks and duplicates."""
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
                values.append(df.iloc[p_idx, idx] if idx < df.shape[1] else "")
            bold_rows.append(participant_start_len)

        answers_by_concept[concept] = values
        bold_rows_by_concept[concept] = bold_rows

    max_len = max((len(v) for v in answers_by_concept.values()), default=0)
    padded = {
        concept: values + [""] * (max_len - len(values))
        for concept, values in answers_by_concept.items()
    }
    return pd.DataFrame(padded), bold_rows_by_concept


def apply_bold_markers(ws, bold_rows_by_concept: dict[str, list[int]]) -> None:
    bold_font = Font(bold=True)
    for col_idx, concept in enumerate(CONCEPTS, start=1):
        for data_row_idx in bold_rows_by_concept[concept]:
            ws.cell(row=2 + data_row_idx, column=col_idx).font = bold_font


def build_metadata_sheet(
    old_path: Path, new_path: Path, new_participants_df: pd.DataFrame, labels: list[str]
) -> tuple[pd.DataFrame, pd.DataFrame]:
    run_info = pd.DataFrame(
        [
            {"key": "old_file", "value": str(old_path)},
            {"key": "new_file", "value": str(new_path)},
            {"key": "new_participant_count", "value": len(new_participants_df)},
            {"key": "concept_count", "value": len(CONCEPTS)},
            {"key": "slots_per_concept", "value": BLOCK_SIZE},
        ]
    )

    participant_labels = [f"Participant {i + 1}" for i in range(len(new_participants_df))]
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
        participant_meta[name] = new_participants_df[f"col_{idx}"].tolist()

    if "Duration (in seconds)" in participant_meta.columns:
        participant_meta["Duration (minutes)"] = (
            pd.to_numeric(participant_meta["Duration (in seconds)"], errors="coerce") / 60.0
        )

    return run_info, participant_meta


def export_update_summary(old_path: Path, new_path: Path, output_path: Path | None = None) -> Path:
    old_df, old_labels = load_qualtrics_data(old_path)
    new_df, new_labels = load_qualtrics_data(new_path)

    new_participants_df = get_new_participants(old_df, new_df, old_labels, new_labels)

    features_df, features_bold = build_compact_summary_table(
        new_participants_df, new_labels, "Feature"
    )
    subtypes_df, subtypes_bold = build_compact_summary_table(
        new_participants_df, new_labels, "Sub type"
    )
    run_info, participant_meta = build_metadata_sheet(
        old_path, new_path, new_participants_df, new_labels
    )

    if output_path is None:
        mm_dd = extract_mm_dd(new_path.name)
        output_path = new_path.parent / f"gemnet_task-1_{mm_dd}_update_summary.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        features_df.to_excel(writer, index=False, sheet_name="features")
        subtypes_df.to_excel(writer, index=False, sheet_name="sub types")
        run_info.to_excel(writer, index=False, sheet_name="metadata", startrow=0)
        participant_meta.to_excel(writer, index=False, sheet_name="metadata", startrow=len(run_info) + 2)
        apply_bold_markers(writer.book["features"], features_bold)
        apply_bold_markers(writer.book["sub types"], subtypes_bold)

    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare two Qualtrics files and summarize only new participants."
    )
    parser.add_argument(
        "--old",
        type=str,
        default=OLD_FILE,
        help=f"Previous Qualtrics export filename or absolute path (default: {OLD_FILE}).",
    )
    parser.add_argument(
        "--new",
        type=str,
        default=NEW_FILE,
        help=f"Latest Qualtrics export filename or absolute path (default: {NEW_FILE}).",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help="Optional output .xlsx path. Default: beside new file with *_update_summary.xlsx name.",
    )
    return parser.parse_args()


def resolve_input_path(path_or_name: str) -> Path:
    """Resolve input against DATA_DIR when a filename is provided."""
    candidate = Path(path_or_name)
    if candidate.is_absolute():
        return candidate
    return DATA_DIR / candidate


def main() -> None:
    args = parse_args()
    old_path = resolve_input_path(args.old)
    new_path = resolve_input_path(args.new)

    if not old_path.exists():
        raise FileNotFoundError(f"Old file not found: {old_path}")
    if not new_path.exists():
        raise FileNotFoundError(f"New file not found: {new_path}")

    out_path = export_update_summary(old_path, new_path, args.out)
    print(f"Created update summary: {out_path}")


if __name__ == "__main__":
    main()
