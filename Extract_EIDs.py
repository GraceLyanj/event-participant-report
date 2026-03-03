import os
import sys
from typing import Set, List


def _check_imports() -> None:
    """Ensure required third‑party packages are available."""
    missing: List[str] = []
    try:
        import pandas  # type: ignore
    except ImportError:
        missing.append("pandas")

    if missing:
        exe = sys.executable
        print("Missing required package(s):", ", ".join(missing))
        print("Using Python:", exe)
        print("Run this in the same terminal you use to run the script:")
        print(f'  "{exe}" -m pip install {" ".join(missing)}')
        sys.exit(1)


_check_imports()

import pandas as pd  # noqa: E402


def _find_eid_columns(columns: List[str]) -> List[str]:
    """Return column names that look like EID fields."""
    eid_cols: List[str] = []
    for col in columns:
        low = col.strip().lower()
        if low == "eid" or "eid" in low:
            eid_cols.append(col)
    return eid_cols


def extract_eids_from_csv(path: str) -> Set[str]:
    """Extract EIDs from a single CSV file, if it has EID-like columns."""
    try:
        df = pd.read_csv(path)
    except Exception as exc:  # pragma: no cover - defensive
        print(f"Skipping {os.path.basename(path)} (could not read CSV: {exc})")
        return set()

    df.columns = df.columns.str.strip()
    eid_cols = _find_eid_columns(list(df.columns))
    if not eid_cols:
        print(f"No EID column found in {os.path.basename(path)}; skipping.")
        return set()

    eids: Set[str] = set()
    for col in eid_cols:
        series = df[col]
        for val in series:
            # Skip true missing values
            if pd.isna(val):
                continue
            # Normalize to string
            s = str(val).strip()
            if not s:
                continue
            low = s.lower()
            if low in {"nan", "na", "none"}:
                continue
            eids.add(s)
    print(f"Found {len(eids)} unique EID(s) in {os.path.basename(path)}.")
    return eids


def extract_eids_from_directory(directory: str) -> Set[str]:
    """Walk a directory (non‑recursive) and collect unique EIDs from all CSV files."""
    if not os.path.isdir(directory):
        raise NotADirectoryError(f"Not a directory: {directory}")

    all_eids: Set[str] = set()
    csv_files = [
        os.path.join(directory, name)
        for name in os.listdir(directory)
        if name.lower().endswith(".csv")
    ]

    if not csv_files:
        print("No CSV files found in directory.")
        return set()

    print(f"Scanning {len(csv_files)} CSV file(s) in {directory}")
    for csv_path in csv_files:
        all_eids |= extract_eids_from_csv(csv_path)

    print(f"Total unique EID(s) across directory: {len(all_eids)}")
    return all_eids


def write_eids_to_csv(eids: Set[str], out_path: str) -> None:
    """Write unique EIDs to a single‑column CSV."""
    if not eids:
        print("No EIDs to write; output file will not be created.")
        return
    df = pd.DataFrame(sorted(eids), columns=["EID"])
    df.to_csv(out_path, index=False)
    print(f"Unique EIDs written to: {out_path}")


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python Extract_EIDs.py <directory_with_csv_files>")
        sys.exit(1)

    directory = sys.argv[1]
    try:
        eids = extract_eids_from_directory(directory)
    except NotADirectoryError as exc:
        print(exc)
        sys.exit(1)

    if not eids:
        print("No EIDs found in any CSV file.")
        sys.exit(0)

    out_path = os.path.join(os.path.abspath(directory), "all_EIDs.csv")
    write_eids_to_csv(eids, out_path)


if __name__ == "__main__":
    main()

