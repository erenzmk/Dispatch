 Dispatch

This repository contains daily call reports for field technicians.  The
`process_reports.py` script summarises the calls and updates `Liste.xlsx` with
per-technician statistics.

## Usage

```bash
python process_reports.py Juli_25/01.07 Liste.xlsx
```

The command expects the directory for the day (e.g. `Juli_25/01.07`) and the
path to the master workbook (`Liste.xlsx`).  The script requires
[openpyxl](https://openpyxl.readthedocs.io/) to be installed.
