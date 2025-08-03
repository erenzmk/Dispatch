# Dispatch

This repository contains daily call reports for field technicians.  The
`process_reports.py` script summarises the calls and updates `Liste.xlsx` with
per-technician statistics.  It expects a morning report (file name contains
`7`) and an evening report (`19`) inside the given day directory.  By comparing
both reports the script determines how many calls were completed during the
day.

## Usage

```bash
python process_reports.py Juli_25/01.07 Liste.xlsx
```

The command expects the directory for the day (e.g. `Juli_25/01.07`) and the
path to the master workbook (`Liste.xlsx`).  The script requires
[openpyxl](https://openpyxl.readthedocs.io/) to be installed.

## Troubleshooting

### Empty summaries
If no calls are written for a technician, ensure the daily report follows the
expected template. The script searches for a row containing `Employee ID` to
locate the header. Files lacking this marker produce empty summaries.

### Missing files
The day directory must contain a morning report (`*7*.xlsx`) and may optionally
include an evening report (`*19*.xlsx`). If either file is missing or named
incorrectly, processing will fail with `StopIteration` or `FileNotFoundError`.

### Incorrect sheet names
`Liste.xlsx` needs a worksheet named after the target month, such as `Juli_25`.
A `KeyError` indicates the sheet is missing or misnamed. Verify that the sheet
matches the pattern `<German month>_<yy>`.

### Verbose logging
To gain more insight into the processing steps, enable Python's logging at the
debug level when running the script:

```bash
LOGLEVEL=DEBUG python process_reports.py Juli_25/01.07 Liste.xlsx
```

This configuration outputs additional details either to the console or to a
file if logging is redirected.
