from pathlib import Path

import pandas as pd

from dispatch.or_liste import parse_or_liste, group_by_day


def test_parse_and_group(tmp_path: Path) -> None:
    data = {
        0: ["name", "Kuersad", "Serghei S."],
        1: ["date", "2025-07-01", "2025-07-02"],
        2: ["weekday", None, None],
        3: ["name", "Alice", None],
        4: ["date", "2025-07-03", None],
        5: ["weekday", None, None],
    }
    df = pd.DataFrame(data)
    file = tmp_path / "Liste.xlsx"
    df.to_excel(file, sheet_name="Juni_25", header=False, index=False)

    assignments = parse_or_liste(file)
    assert list(assignments.columns) == ["date", "tech"]
    assert len(assignments) == 3
    assert assignments.loc[0, "tech"] == "Kürşad"

    grouped = group_by_day(assignments)
    expected = {
        pd.Timestamp("2025-07-01"): ["Kürşad"],
        pd.Timestamp("2025-07-02"): ["Serghei S."],
        pd.Timestamp("2025-07-03"): ["Alice"],
    }
    assert grouped == expected
