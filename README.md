# KOLAO daily sales report

Builds the daily KOLAO sales reports (order status, sales, stock, new models).

- `Setup.bat` installs what the report needs (run once).
- `Run.bat` downloads today's exports from KolaoERP into `Input` (`erp_export.py`; the ERP is closed after a finished download), then builds the report (`sales_report.py`). The download asks which data in a tick window; `erp_export.py --all` takes everything, `--only DL,KLM,LXC,KDC` and `--types Sales,Stock,Order` take a part (Diana uses these). Cancel = exit code 2, nothing downloaded.
- **Diana** (her own add-on, in `D:\Diana`) is the 8-bit pixel girl who runs all this for you: click her on the desktop and choose what to do. When she is installed, `Run.bat` opens her instead. The report's Yes/No questions then come up in her chat panel (through `DIANA_PORT`); without her they are normal Yes/No boxes.
- `how-to-use.html` has the step-by-step guide. `workflow.html` and `workflow-diagram.html` show how it works.
