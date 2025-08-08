import datetime as dt
from pathlib import Path
import pandas as pd
import pytest

from dispatch.process_calls import process_report, vorheriger_werktag


@pytest.mark.parametrize(
    "excel_name, query_name",
    [
        ("Ahmad, Daniyal (Keskin)", "Ahmad, Daniyal (Keskin)"),
        ("Ahmad, Daniyal (Keskin)", "daniyal"),
        ("AHMAD, DANIYAL (KESKIN)", " AhMaD, Daniyal (Keskin) "),
        ("Ahmad, Daniyal (Keskin)", "Daniyal Ahmad"),
        ("Doe, John (Team)", "john"),
    ],
)
def test_process_report_filters_and_classifies(tmp_path, excel_name, query_name):
    report_date = dt.date(2024, 3, 15)
    prev_day = vorheriger_werktag(report_date)
    old_date = prev_day - dt.timedelta(days=2)

    data1 = pd.DataFrame(
        {
            "Techniker": [excel_name, excel_name, "Andere Person"],
            "Callnr": ["17500001", "18000001", "17500002"],
            "Erstellt": [
                prev_day.strftime("%d.%m.%Y"),
                prev_day.strftime("%d.%m.%Y"),
                prev_day.strftime("%d.%m.%Y"),
            ],
        }
    )
    data2 = pd.DataFrame(
        {
            "Techniker": [excel_name, excel_name],
            "Callnr": ["17500003", "17500004"],
            "Erstellt": [
                old_date.strftime("%d.%m.%Y"),
                prev_day.strftime("%d.%m.%Y"),
            ],
        }
    )
    meta = pd.DataFrame({"Berichtstag": [report_date.strftime("%d.%m.%Y")]})

    file_path = tmp_path / "report.xlsx"
    with pd.ExcelWriter(file_path) as writer:
        data1.to_excel(writer, index=False, sheet_name="Sheet1")
        data2.to_excel(writer, index=False, sheet_name="Sheet2")
        meta.to_excel(writer, index=False, sheet_name="Meta")

    df = process_report(file_path, query_name)

    assert set(df["Callnr"]) == {"17500001", "17500003", "17500004"}

    status_map = dict(zip(df["Callnr"], df["Status"]))
    assert status_map["17500003"] == "alt"
    assert status_map["17500001"] == "neu"
    assert status_map["17500004"] == "neu"


def test_process_report_missing_file():
    with pytest.raises(FileNotFoundError):
        process_report("does_not_exist.xlsx", "Ahmad, Daniyal (Keskin)")


def test_process_report_vergangener_berichtstag(tmp_path):
    report_date = dt.date(2024, 3, 20)
    prev_day = vorheriger_werktag(report_date)

    data = pd.DataFrame(
        {
            "Techniker": ["Ahmad, Daniyal (Keskin)", "Ahmad, Daniyal (Keskin)"],
            "Callnr": ["17500005", "17500006"],
            "Erstellt": [
                prev_day.strftime("%d.%m.%Y"),
                (prev_day - dt.timedelta(days=1)).strftime("%d.%m.%Y"),
            ],
        }
    )
    meta = pd.DataFrame({"Berichtstag": [report_date.strftime("%d.%m.%Y")]})

    file_path = tmp_path / "report_past.xlsx"
    with pd.ExcelWriter(file_path) as writer:
        data.to_excel(writer, index=False, sheet_name="Daten")
        meta.to_excel(writer, index=False, sheet_name="Meta")

    df = process_report(file_path, "Ahmad, Daniyal (Keskin)")

    status_map = dict(zip(df["Callnr"], df["Status"]))
    assert status_map["17500005"] == "neu"
    assert status_map["17500006"] == "alt"
