"""Uploads the signed car permits (D:\\Auto sign\\signed\\<date>\\<frame no.>.pdf) to their bookings in KolaoERP.

  1. starts KolaoERP and logs in as KDC (credentials.json), opens VS_Booking_Sale: SRM ADM, Book Date one year back
  2. each PDF of the chosen date folders: Frame# = the file name, Query, drags the PDF from an Explorer window into the
     Sale Document tab, ticks MN Chk, saves, and reads the row back
  3. a folder whose PDFs are all in the ERP moves to signed\\uploaded, so it is not offered again

STOP: the small red window on top, or the Esc key, stops the run at once. The run also stops by itself, clicking
nothing, as soon as something is not what it expects (a message box, a drop or tick that does not take). A PDF whose
frame number finds no booking, or more than one, is skipped and its folder stays.
"""
import ctypes
import datetime as dt
import json
import shutil
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SIGNED = Path(r"D:\Auto sign\signed")    # Auto sign's output: one folder per date (MM-DD-YYYY)
UPLOADED = SIGNED / "uploaded"           # date folders whose permits are all in the ERP

# ---- settings ---------------------------------------------------------------------------------
ACCOUNT = "KDC"                          # the login in credentials.json
MENU_TOP, FORM_NAME, FORM = "Vehicle Sales", "VS_Booking_Sale", "SDB496"
SRM = "ADM"
QUERY_PATIENCE = 30                      # seconds for a Query's bookings to come in
# -----------------------------------------------------------------------------------------------


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


# ---- the ERP ------------------------------------------------------------------------------------
import erp_export as erp  # noqa: E402  (its helpers; nothing of it runs on import)
from erp_export import Stopped, check, desk, fail, field, front, main_window, message_boxes, wait  # noqa: E402


def uia(handle):
    return desk("uia").window(handle=handle).wrapper_object()


def centre(ctrl):
    r = ctrl.rectangle()
    return (r.left + r.right) // 2, (r.top + r.bottom) // 2


def no_box(when):
    """Any message from the ERP stops the run: it is left on screen, nothing clicked."""
    if message_boxes():
        box = message_boxes()[0]
        fail(f"The ERP showed a message {when}: {box.window_text()} {' '.join(c.window_text() for c in box.children())}"
             " (left as it is, nothing clicked)")


def again(read):
    """A read of the ERP's lists, tried again when the ERP has just replaced one of its windows."""
    def run(*args):
        for attempt in range(3):
            try:
                return read(*args)
            except Stopped:
                raise
            except Exception:
                if attempt == 2:
                    raise
                erp._found.clear()
                wait(1)
    return run


def open_form():
    """Vehicle Sales -> VS_Booking_Sale: looked for by name in the menu and its sub-menus, then waits for the form."""
    from pywinauto import mouse
    w = erp.menu_bar()
    top = [c for c in w.descendants(control_type="MenuItem")
           if c.parent().window_text() == "MenuStrip1" and c.window_text() == MENU_TOP]
    if not top:
        fail(f"Menu '{MENU_TOP}' not found - the ERP menu changed.")
    mouse.click(coords=centre(top[0]))
    wait(0.8)
    items = [c for c in top[0].children(control_type="MenuItem") if c.rectangle().width() > 0]
    found = [c for c in items if c.window_text() == FORM_NAME]
    for item in items if not found else []:  # not on the first level: open each sub-menu in turn
        mouse.move(coords=centre(item))
        wait(0.6)
        found = [c for c in item.children(control_type="MenuItem")
                 if c.window_text() == FORM_NAME and c.rectangle().width() > 0]
        if found:
            break
    if not found:
        fail(f"'{FORM_NAME}' not found under '{MENU_TOP}' - the ERP menu changed.")
    mouse.click(coords=centre(found[0]))
    for _ in range(60):
        check()
        if erp.form(FORM):
            wait(1)
            return
        time.sleep(0.5)
    fail(f"{FORM_NAME} did not open.")


def shown(ctrl):
    """The text a DevExpress box shows (its own value, or its text box's)."""
    for c in [uia(ctrl.handle)] + uia(ctrl.handle).descendants():
        for v in (c.legacy_properties().get("Value"), c.window_text()):
            if v and v.strip():
                return v.strip()
    return ""


def set_srm():
    box = [c for c in field(FORM, "f_show_code").children() if c.element_info.automation_id == "ComboBoxEdit1"]
    box = box[0] if box else fail("SRM has no list - the ERP screen changed.")
    if shown(box) == SRM:
        return
    front(main_window())
    erp.click_at(box.rectangle())
    wait(1)
    ps = erp.popups()
    items = [c for c in uia(ps[0].handle).descendants(control_type="ListItem")
             if c.window_text().split()[:1] == [SRM]] if ps else []
    if not items:
        fail(f"The SRM list has no '{SRM}'.")
    items[0].click_input()
    wait(0.8)
    if shown(box) != SRM:
        fail(f"SRM shows '{shown(box)}', not {SRM}.")


def open_documents():
    """The Sale Document tab, where the PDFs go."""
    page = field(FORM, "XtraTabPage3")
    if page.is_visible() and page.rectangle().width() > 100:
        return
    tabs = [c for c in uia(field(FORM, "XtraTabControl1").handle).descendants(control_type="TabItem")
            if c.window_text() == "Sale Document"]
    if not tabs:
        fail("The Sale Document tab was not found - the ERP screen changed.")
    tabs[0].click_input()
    wait(1)
    if not (page.is_visible() and page.rectangle().width() > 100):
        fail("The Sale Document tab did not open.")


def doc_grid():
    gs = [c for c in field(FORM, "XtraTabPage3").descendants() if c.element_info.automation_id == "Grid1"]
    return gs[0] if gs else fail("The Sale Document tab has no list - the ERP screen changed.")


@again
def bookings():
    """Book# of each booking the Query found (the form's first list)."""
    return [c.legacy_properties().get("Value") for c in uia(field(FORM, "Grid1").handle).descendants(control_type="DataItem")
            if c.window_text().startswith("Book# row")]


@again
def rows():
    """{row no.: {column: cell}} of the Sale Document list."""
    out = {}
    for c in uia(doc_grid().handle).descendants(control_type="DataItem"):
        col, _, k = c.window_text().rpartition(" row ")
        if k.isdigit():
            out.setdefault(int(k), {})[col] = c
    return out


def files():
    return [r["File Name"].legacy_properties().get("Value") for _, r in sorted(rows().items()) if "File Name" in r]


def ticked(cell):
    """MN Chk as drawn: a tick is dark pixels, the empty box (grey, or orange when selected) has none. (Read through UI
    automation the box keeps its old value until saved.)"""
    from PIL import ImageGrab
    x, y = centre(cell)
    return sum(1 for p in ImageGrab.grab((x - 9, y - 9, x + 9, y + 9)).convert("L").tobytes() if p < 90) >= 6


def settle(read, quiet=1.0, patience=30):
    """read() until it gives the same answer for `quiet` seconds."""
    last, since, end = read(), time.time(), time.time() + patience
    while time.time() < end:
        wait(0.3)
        now = read()
        if now != last:
            last, since = now, time.time()
        elif time.time() - since >= quiet:
            break
    return last


def query(frame_no):
    """Frame# -> Query. Returns the bookings found, or None when the list did not change (a Query that did not run)."""
    before = bookings()
    box = erp.edit_box(FORM, "f_frame_no")
    front(main_window())
    box.click_input()
    wait(0.2)
    box.type_keys("^a{BACKSPACE}" + frame_no)
    wait(0.3)
    if box.window_text() != frame_no:
        fail(f"Frame# shows '{box.window_text()}', not {frame_no}.")
    box.type_keys("{TAB}")  # the ERP takes the number only on leaving the box: Query(Q) does not take the focus
    wait(0.5)
    desk("uia").window(title_re=f".*{erp.MAIN_TITLE}").child_window(title="Query(Q)", control_type="Button",
                                                                     found_index=0).click_input()
    end = time.time() + QUERY_PATIENCE
    while time.time() < end:  # the bookings come in late, while the ERP keeps answering
        wait(0.4)
        no_box("after Query")
        if bookings() != before:
            return settle(bookings)
    return None


def drag(pdf):
    """A real mouse drag of the PDF from an Explorer window onto the Sale Document list (the list takes no other kind
    of drop). The Explorer window is parked left of the list and closed again."""
    from pywinauto import mouse
    u = ctypes.windll.user32
    grid = doc_grid().rectangle()
    explorers = lambda: [w for w in desk().windows() if w.class_name() == "CabinetWClass"]  # noqa: E731
    had = {w.handle for w in explorers()}
    subprocess.Popen(["explorer", "/select,", str(pdf)])
    ex = None
    for _ in range(40):
        wait(0.25)
        ws = [w for w in explorers() if w.window_text().startswith(pdf.parent.name + " - ")]
        if ws:
            ex = ws[0]
            break
    if ex is None:
        fail("The Explorer window did not open.")
    try:
        u.ShowWindow(ex.handle, 9)  # SW_RESTORE
        u.MoveWindow(ex.handle, 0, grid.top, max(300, grid.left - 10), grid.height(), True)
        front(ex)
        wait(1.2)
        items = [c for c in uia(ex.handle).descendants(control_type="ListItem") if c.window_text() in (pdf.stem, pdf.name)]
        if not items:
            fail(f"{pdf.name} is not in the Explorer window.")
        r = items[0].rectangle()
        sx, sy = r.left + 40, (r.top + r.bottom) // 2
        tx, ty = (grid.left + grid.right) // 2, (grid.top + grid.bottom) // 2
        mouse.press(coords=(sx, sy))
        wait(0.3)
        for k in range(1, 31):  # in small steps: Windows starts a drag only once the mouse moves with the button down
            mouse.move(coords=(sx + (tx - sx) * k // 30, sy + (ty - sy) * k // 30))
            time.sleep(0.03)
        wait(0.8)
        mouse.release(coords=(tx, ty))
        wait(2)
    finally:
        if ex.handle not in had:  # one of yours stays open
            u.PostMessageW(ex.handle, 0x0010, 0, 0)  # WM_CLOSE
    wait(0.5)


def upload(pdf, log):
    """One permit. Returns 'OK', 'ALREADY' or 'SKIP'; anything else stops the run (Stopped)."""
    from pywinauto import mouse
    no = pdf.stem
    front(main_window())
    no_box("before the next permit")
    found = query(no)
    if found is None:
        log(f"  SKIP    {no}: the booking list did not change after Query")
        return "SKIP"
    if len(found) != 1:
        log(f"  SKIP    {no}: Query found {len(found)} bookings {found}")
        return "SKIP"
    book = found[0]
    have = settle(files)
    if pdf.name in have:
        log(f"  ALREADY {no}: in {book}")
        return "ALREADY"
    drag(pdf)
    no_box("after the drop")
    now = settle(files, quiet=0.8, patience=8)
    if len(now) != len(have) + 1 or now.count(pdf.name) != 1:
        fail(f"{no}: after the drop {book} lists {now}, not one new {pdf.name} - not saved.")
    k = [i for i, r in sorted(rows().items()) if r["File Name"].legacy_properties().get("Value") == pdf.name][0]
    front(main_window())
    wait(0.3)
    for _ in range(2):  # a click ticks it; the first may only select the cell: looked at on screen after each
        if ticked(rows()[k]["MN Chk"]):
            break
        mouse.click(coords=centre(rows()[k]["MN Chk"]))
        wait(0.8)
        no_box("after ticking MN Chk")
    if not ticked(rows()[k]["MN Chk"]):
        fail(f"{no}: MN Chk did not tick - not saved ({book}, row {k + 1}).")
    desk("uia").window(title_re=f".*{erp.MAIN_TITLE}").child_window(title="Save(S)", control_type="Button",
                                                                     found_index=0).click_input()
    wait(2.5)
    no_box("after Save")
    row = settle(rows, quiet=0.8, patience=8)[k]
    name = row["Regional MN Name"].legacy_properties().get("Value")
    if not ticked(row["MN Chk"]) or not name:
        fail(f"{no}: after Save {book} shows MN Chk {'ticked' if ticked(row['MN Chk']) else 'not ticked'}, "
             f"name '{name}' - check it in the ERP.")
    log(f"  OK      {no}: uploaded to {book}, MN Chk by {name.strip()}")
    return "OK"


def run(folders, log):
    """Returns the number of permits skipped."""
    cred = json.loads(erp.CREDENTIALS.read_text(encoding="utf-8")) if erp.CREDENTIALS.exists() else fail(
        f"{erp.CREDENTIALS.name} is missing: run Setup.bat to enter the ERP logins.")
    c = cred.get(ACCOUNT) or {}
    if not c.get("user") or not c.get("password"):
        fail(f"No ERP login for {erp.LOGINS[ACCOUNT]} on this PC: run Setup.bat to enter it.")
    erp.login(ACCOUNT, c, log)
    open_form()
    set_srm()
    erp.set_date(FORM, "fr_dt", "year_ago")
    erp.set_date(FORM, "to_dt", "today")
    open_documents()
    log(f"{FORM_NAME} open: SRM {SRM}, Book Date {erp.day_for('year_ago')} - {erp.day_for('today')}")
    skipped = 0
    for folder in folders:
        log(f"{folder.name}:")
        got = [upload(pdf, log) for pdf in sorted(folder.glob("*.pdf"))]
        if "SKIP" in got:
            skipped += got.count("SKIP")
            log(f"  {folder.name} stays in signed: {got.count('SKIP')} permit(s) not uploaded")
        else:
            log(f"  all in the ERP: {folder.name} -> {move_done(folder)}")
    return skipped


def main():
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    lines = []

    def log(msg):
        print(msg, flush=True)
        lines.append(msg)

    if "--list" in sys.argv:  # the folders still to upload
        print("\n".join(p.name for p in pending_folders()))
        return 0
    wanted = sys.argv[sys.argv.index("--folders") + 1].split(",") if "--folders" in sys.argv[:-1] else None
    folders = [p for p in pending_folders() if wanted is None or p.name in wanted]
    if not folders:
        print("No signed permits to upload" + (f" in {', '.join(wanted)}" if wanted else "") + " - nothing done.")
        return 2
    try:  # real screen pixels for every window, so clicks and the STOP box land where they are meant to
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except OSError:
        pass
    threading.Thread(target=erp._stop_window, args=("Permit upload",), daemon=True).start()
    log(f"Permit upload {stamp}  (STOP window or Esc stops it)")
    log("Folders: " + ", ".join(f"{p.name} ({len(list(p.glob('*.pdf')))})" for p in folders))
    code = 1
    try:
        skipped = run(folders, log)
        log("")
        log("Done." if not skipped else f"Done, {skipped} permit(s) skipped - see SKIP above.")
        subprocess.run(["taskkill", "/IM", erp.ERP_PROCESS, "/F"], capture_output=True)  # finished: close the ERP
        log("KolaoERP closed.")
        code = 0 if not skipped else 1
    except Stopped as e:  # the ERP stays open as it is, to look at
        log("")
        log(f"STOPPED: {e}")
    except Exception as e:
        log("")
        log(f"STOPPED (unexpected): {type(e).__name__}: {e}")
    finally:
        erp._stop.set()
        (ROOT / "output").mkdir(exist_ok=True)
        (ROOT / "output" / f"ERP upload {stamp}.txt").write_text("\n".join(lines), encoding="utf-8-sig")
    return code


if __name__ == "__main__":
    sys.coinit_flags = 2  # single-threaded COM, as the Explorer drag and the ERP's lists need
    sys.exit(main())
