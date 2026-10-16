# Diana The KOLAO sale department reporter bot

Diana builds the daily KOLAO sales reports (order status, sales, stock, new models).

- `Setup.bat` installs what Diana needs (run once).
- `Run.bat` runs Diana: she downloads the ERP exports into `Input`, builds the report and makes the pictures, while her 8-bit animation (`diana.py`) plays in the bottom-right corner. She wears a new outfit every day from a list of 33 (casual, modern, office, sporty, anime sailor, Lao sinh, Korean hanbok) and starts over after the last one. Add more in `OUTFITS` in `diana.py`. Try it without running the job: `python diana.py --demo` or `--demo fail` (add a number to try that outfit).
- `how-to-use.html` has the step-by-step guide. `workflow.html` and `workflow-diagram.html` show how it works.
