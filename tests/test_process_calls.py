import datetime as dt
import pandas as pd
import pytest

from process_calls import process_report


@pytest.mark.parametrize(
    "excel_name, query_name",
    [
        ("Ahmad, Daniyal (Keskin)", "Ahmad, Daniyal (Keskin)"),
        ("Ahmad, Daniyal (Keskin)", "daniyal"),
        ("AHMAD, DANIYAL (KESKIN)", " AhMaD, Daniyal (Keskin) "),
        ("Ahmad, Daniyal (Keskin)", "Daniyal Ahmad"),
    ],
)
def test_process_report_filters_and_classifies(tmp_path, excel_name, query_name):
    today = dt.date.today()
    old_date = today - dt.timedelta(days=2)

    data1 = pd.DataFrame(
        {
            "Techniker": [excel_name, excel_name, "Andere Person"],
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
            "Techniker": [excel_name, excel_name],
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

    df = process_report(file_path, query_name)

    assert set(df["Callnr"]) == {"17500001", "17500003", "17500004"}

    status_map = dict(zip(df["Callnr"], df["Status"]))
    assert status_map["17500003"] == "alt"
    assert status_map["17500001"] == "neu"
    assert status_map["17500004"] == "neu"


def test_process_report_missing_file():
    with pytest.raises(FileNotFoundError):
        process_report("does_not_exist.xlsx", "Ahmad, Daniyal (Keskin)")
