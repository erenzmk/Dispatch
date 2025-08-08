@echo off
set /p month=Monat (YYYY-MM) eingeben:
python "%~dp0run_current_month.pyw" %month%
