import pandas as pd
import openpyxl
from summarize_calls import summarize_report


def test_summarize_report(tmp_path):
    excel = tmp_path / "report.xlsx"
    with pd.ExcelWriter(excel) as writer:
        pd.DataFrame({"first": ["Daniyal"], "last": ["Ahmad"]}).to_excel(
            writer, index=False, sheet_name="Techniker DK"
        )
        pd.DataFrame(columns=["first", "last"]).to_excel(
            writer, index=False, sheet_name="Berlin dk"
        )
        pd.DataFrame(columns=["first", "last"]).to_excel(
            writer, index=False, sheet_name="Berlin rest"
        )
        data = pd.DataFrame(
            [
                ["Daniyal Ahmad", "", "17500001", "", "", "", "", "30-06-2025 09:04"],
                ["Daniyal Ahmad", "", "17500002", "", "", "", "", "26-06-2025 15:03"],
                ["Unknown Tech", "", "17500003", "", "", "", "", "30-06-2025 09:04"],
                ["Daniyal Ahmad", "", "18000001", "", "", "", "", "30-06-2025 09:04"],
            ],
            columns=list("ABCDEFGH"),
        )
        data.to_excel(writer, index=False, sheet_name="West Central", startrow=20, header=False)

    wb = openpyxl.load_workbook(excel)
    ws = wb["West Central"]
    ws["A2"] = "01.07.25 07:01"
    wb.save(excel)
    wb.close()

    result = summarize_report(excel)
    assert len(result) == 1
    row = result.iloc[0]
    assert row["technician"] == "Daniyal Ahmad"
    assert row["new"] == 1
    assert row["old"] == 1
    assert row["total"] == 2
    assert str(row["date"]) == "2025-07-01"
