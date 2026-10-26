"""Uploads the signed car permits (D:\\Auto sign\\signed\\<date>\\<frame no.>.pdf) to their bookings in KolaoERP.

  1. starts KolaoERP and logs in as KDC (credentials.json), opens VS_Booking_Sale: SRM ADM, Book Date one year back
  2. each PDF of the chosen date folders: Frame# = the file name, Query, drags the PDF from an Explorer window into the
     Sale Document tab, ticks MN Chk, saves, and reads the row back
  3. a folder whose PDFs are all in the ERP moves to signed\\uploaded, so it is not offered again

STOP: the small red window on top, or the Esc key, stops the run at once. The run also stops by itself, clicking
nothing, as soon as something is not what it expects (a message box, a drop or tick that does not take). A PDF whose
frame number finds no booking, or more than one, is skipped and its folder stays.
"""
import datetime as dt
import shutil
from pathlib import Path

SIGNED = Path(r"D:\Auto sign\signed")    # Auto sign's output: one folder per date (MM-DD-YYYY)
UPLOADED = SIGNED / "uploaded"           # date folders whose permits are all in the ERP


def _by_date(folder):
    try:
        return 0, dt.datetime.strptime(folder.name, "%m-%d-%Y"), folder.name
    except ValueError:  # not a date: after the dates
        return 1, dt.datetime.min, folder.name


def pending_folders(signed=SIGNED):
    """The date folders in signed with PDFs in them, oldest first; uploaded and empty ones left out."""
    if not signed.exists():
        return []
    found = [p for p in signed.iterdir() if p.is_dir() and p.name.lower() != UPLOADED.name and any(p.glob("*.pdf"))]
    return sorted(found, key=_by_date)


def move_done(folder, uploaded=UPLOADED):
    """Moves a fully uploaded date folder into uploaded. A folder of that name there already (the date signed again
    later) gets the new files; the ones it holds are kept."""
    to = uploaded / folder.name
    uploaded.mkdir(parents=True, exist_ok=True)
    if not to.exists():
        shutil.move(str(folder), str(to))
        return to
    for f in folder.iterdir():
        if not (to / f.name).exists():
            shutil.move(str(f), str(to / f.name))
    shutil.rmtree(folder)
    return to
