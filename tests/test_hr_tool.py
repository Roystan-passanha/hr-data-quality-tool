import subprocess
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import hr_tool  # noqa: E402


def test_parse_date_handles_mixed_formats_and_impossible_dates():
    assert hr_tool.parse_date("05/03/2027") == pd.Timestamp("2027-03-05")   # day first
    assert hr_tool.parse_date("2027-03-05") == pd.Timestamp("2027-03-05")
    assert hr_tool.parse_date("05 Mar 2027") == pd.Timestamp("2027-03-05")
    assert pd.isna(hr_tool.parse_date("31/02/2027"))
    assert pd.isna(hr_tool.parse_date("TBC"))


def test_clean_flags_planted_errors(tmp_path):
    subprocess.run([sys.executable, str(ROOT / "make_sample.py")], cwd=tmp_path, check=True)
    out = tmp_path / "clean.xlsx"
    hr_tool.clean(tmp_path / "hr_sample_messy.xlsx", out)
    clean = pd.read_excel(out, sheet_name="Clean_Data")
    issues = pd.read_excel(out, sheet_name="Issues")
    problems = set(issues["Problem"])
    assert "Not a valid date" in problems
    assert "Exact duplicate row (removed)" in problems
    assert "Same Employee ID used for different records" in problems
    assert len(clean) == 600                                  # 8 exact duplicates removed
    assert set(clean["Work Country"].dropna()) <= {"UAE", "KSA", "Qatar", "India", "UK", "USA"}
    assert issues["Source Row"].min() >= 2                    # Excel row numbers, not index
