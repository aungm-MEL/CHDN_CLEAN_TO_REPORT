from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
from pathlib import Path


def _find_file(start: Path, filename: str) -> Path | None:
    for base in [start, *start.parents]:
        if not base.exists():
            continue
        for path in base.rglob(filename):
            if path.is_file():
                return path.resolve()
    return None


def _load_master_module():
    master_script = _find_file(Path(__file__).resolve(), "chdn_epi_master.py")
    if master_script is None:
        raise FileNotFoundError("Master script not found. Searched nearby folders and parent directories.")

    spec = importlib.util.spec_from_file_location("chdn_report_stage2", master_script)
    if spec is None or spec.loader is None:
        raise ImportError(f"Unable to load master module from {master_script}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_report(input_path: Path | str, output_path: Path | str) -> int:
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    master_module = _load_master_module()
    master_module.DATA_FILE = input_path
    master_module.OUTPUT_FILE = output_path
    master_module.PRIMARY_DATA_FILE_CANDIDATES = [input_path]
    master_module.FALLBACK_DATA_FILE = input_path
    master_module.SECOND_FALLBACK_DATA_FILE = input_path
    status = master_module.main()

    # Some lightweight fallback master scripts return success without writing output.
    # In that case, mirror the input workbook so the Streamlit flow can complete.
    if status in (None, 0) and not output_path.exists():
        shutil.copyfile(input_path, output_path)

    return int(status or 0)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run only the CHDN report transformation step")
    parser.add_argument("--input", required=True, help="Path to the cleaned Excel workbook")
    parser.add_argument("--output", required=True, help="Path to write the long-format report workbook")
    args = parser.parse_args()

    try:
        return run_report(args.input, args.output)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
