import datetime as dt
from pathlib import Path
import pandas as pd
from process_calls import process_report


def test_process_report_filters_and_classifies(tmp_path):
    today = dt.date.today()
    old_date = today - dt.timedelta(days=2)

    data1 = pd.DataFrame(
        {
            "Techniker": [
                "Ahmad, Daniyal (Keskin)",
                "Ahmad, Daniyal (Keskin)",
                "Andere Person",
            ],
            "Callnr": ["17500001", "18000001", "17500002"],
            "Erstellt": [
                today.strftime("%d.%m.%Y"),
                today.strftime("%d.%m.%Y"),
                today.strftime("%d.%m.%Y"),
            ],
        }
    )
    data2 = pd.DataFrame(
        {
            "Techniker": ["Ahmad, Daniyal (Keskin)", "Ahmad, Daniyal (Keskin)"],
            "Callnr": ["17500003", "17500004"],
            "Erstellt": [
                old_date.strftime("%d.%m.%Y"),
                today.strftime("%d.%m.%Y"),
            ],
        }
    )

    file_path = tmp_path / "report.xlsx"
    with pd.ExcelWriter(file_path) as writer:
        data1.to_excel(writer, index=False, sheet_name="Sheet1")
        data2.to_excel(writer, index=False, sheet_name="Sheet2")

    df = process_report(file_path, "Ahmad, Daniyal (Keskin)")

    assert set(df["Callnr"]) == {"17500001", "17500003", "17500004"}

    status_map = dict(zip(df["Callnr"], df["Status"]))
    assert status_map["17500003"] == "alt"
    assert status_map["17500001"] == "neu"
    assert status_map["17500004"] == "neu"
