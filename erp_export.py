"""Downloads today's exports from KolaoERP into Input, ready for sales_report.py.

  1. starts KolaoERP and logs in with the account in credentials.json
  2. opens each report (the dealers, then KLM = MBS, LXC = Geely, KDC/KDCX), sets its filters, presses Query and exports it as Excel (Xls)
  3. checks every file: saved, has rows, and sales_report.py recognises it as the part it should be

STOP: the small red window on top, or the Esc key, stops the run at once. The script also stops by itself as soon as
something is not what it expects (a window that does not open, a filter that does not take, an empty or unknown file).
It only uses Query and Export in the ERP - never New, Save or Delete.
"""
import ctypes
import datetime as dt
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "Input"
CREDENTIALS = ROOT / "credentials.json"

# ---- settings ---------------------------------------------------------------------------------
ERP_SHORTCUT = r"C:\Users\Public\Desktop\KolaoERP.lnk"
ERP_PROCESS = "System8.Load.exe"
MAIN_TITLE = "KolaoERP Local"           # the main window's title ends with this
ACCOUNTS = ["DL", "KDC"]                # logins in credentials.json, in this order: the dealers first (their stock moves fastest)
EXCLUDE = {"MC", "EV MC", "EV TRUCK", "EV TUK TUK"}            # car / parts types the report does not carry
SALE_STATUS = ("only", {"Confirm", "Wait Sales Approval"})     # sales: confirmed and waiting for approval
# each report: (account, menu path, form code, file name in Input, parts sales_report may recognise in it, filters).
# Filters, applied in this order:
#   ("checks", field, rule)  a drop-down check list: "all", ("only", {names}) or ("except", {names})
#   ("pick", field, value)   a drop-down with one choice, e.g. NEW
#   ("date", field, day)     "month_start" (1st of this month), "today", "year_ago" (today one year back)
#   ("tick" / "untick", field)  a tick box
KDC_MENU = ("Vehicle Sales", "Reports")
KLM_MENU = ("KLM", "Vehicle Sales", "Report Vehicles")
LXC_MENU = ("LXC", "Vehicle Sales", "Report vihicles")
SALE_DATES = [("date", "fr_dt", "month_start"), ("date", "to_dt", "today"), ("pick", "old_new", "NEW")]  # new cars only
# NEW / USED as the workbook always held it: sales, KDC orders and KDC / DL stock NEW only; the orders of the dealers,
# KLM and LXC NEW and USED (no filter)
ORDER_DATES = [("date", "book_dt_fr", "year_ago"), ("date", "book_dt_to", "today")]
NEW_ONLY = ("pick", "old_new", "NEW")
GEELY = {"Geely", "BMW"}  # LXC holds Geely and BMW (and BMW can have no rows)
REPORTS = [
    # dealers (login DL)
    ("DL", KDC_MENU + ("Sales Detail",), "JSD990", "Sales Report DL.xls", {("Sale", "Dealer")},
     [("checks", "wh_cd", "all"), *SALE_DATES, ("checks", "stat_bc", SALE_STATUS)]),
    ("DL", KDC_MENU + ("Car Stock",), "JSD910", "Stock Report DL.xls", {("Stock", "Dealer")},
     [("checks", "sr_cd", "all"), ("checks", "itm_ty", ("except", EXCLUDE)), NEW_ONLY]),
    ("DL", KDC_MENU + ("Order Status",), "JSD970", "Order Status DL.xls", {("Order", "Dealer")},
     [("checks", "sr_cd", "all"), ("tick", "chk_dt")]),
    # KLM = MBS
    ("KDC", KLM_MENU + ("[KLM]Sale details",), "KLR200", "Sales Report KLM.xls", {("Sale", "MBS")},
     [("checks", "wh_cd", "all"), *SALE_DATES, ("checks", "stat_bc", SALE_STATUS), ("checks", "car_type", ("except", EXCLUDE))]),
    ("KDC", KLM_MENU + ("[KLM]Car stock detail",), "KLR100", "Stock Report KLM.xls", {("Stock", "MBS")},
     [("checks", "sr_cd", "all"), ("checks", "itm_ty", ("except", EXCLUDE))]),
    ("KDC", KLM_MENU + ("[KLM]Booking details",), "KLR300", "Order Report KLM.xls", {("Order", "MBS")},
     [("checks", "sr_cd", "all"), *ORDER_DATES]),
    # LXC = Geely (+ BMW)
    ("KDC", LXC_MENU + ("[BMW]Sale detail",), "BMR200", "Sales Report LXC.xls", {("Sale", b) for b in GEELY},
     [("checks", "wh_cd", "all"), *SALE_DATES, ("checks", "stat_bc", SALE_STATUS), ("checks", "car_type", ("except", EXCLUDE))]),
    ("KDC", LXC_MENU + ("[BMW]Car stock detail",), "BMR100", "Stock Report LXC.xls", {("Stock", b) for b in GEELY},
     [("checks", "sr_cd", "all"), ("checks", "itm_ty", ("except", EXCLUDE))]),
    ("KDC", LXC_MENU + ("[BMW]Booking detail",), "BMR300", "Order Report LXC.xls", {("Order", b) for b in GEELY},
     [("checks", "sr_cd", "all"), *ORDER_DATES]),
    # KDC + KDCX
    ("KDC", KDC_MENU + ("Sales Detail",), "SDB990", "Sales Report KDC KDCX.xls", {("Sale", "KDC")},
     [("checks", "wh_cd", "all"), *SALE_DATES, ("checks", "stat_bc", SALE_STATUS), ("checks", "car_type", ("except", EXCLUDE))]),
    ("KDC", KDC_MENU + ("Car Stock New",), "SDB915", "Stock Report KDC KDCX.xls", {("Stock", "KDC")},
     [("checks", "sr_cd", "all"), ("checks", "itm_ty", ("except", EXCLUDE)), NEW_ONLY]),
    ("KDC", KDC_MENU + ("Order Status",), "SDB970", "Order Status KDC KDCX.xls", {("Order", "KDC")},
     [("checks", "sr_cd", "all"), ("tick", "chk_dt"), NEW_ONLY]),  # a date range here keeps only 19 of ~500 open orders
]
QUERY_TIMEOUT = 600                     # seconds a Query may take
# -----------------------------------------------------------------------------------------------


class Stopped(Exception):
    pass


# ---- stop button / Esc ----------------------------------------------------------------------------
_stop = threading.Event()
_stop_reason = ["pressed STOP"]


def _stop_window():
    import tkinter as tk
    root = tk.Tk()
    root.overrideredirect(True)  # no title bar: a small box
    root.attributes("-topmost", True)
    # top right, over the empty right end of the ERP ribbon: clear of the menus, lists and save windows the run uses
    root.geometry(f"260x64+{root.winfo_screenwidth() - 280}+100")
    box = tk.Frame(root, bg="#c62828", bd=0)
    box.pack(fill="both", expand=True)
    tk.Button(box, text="STOP  (Esc)", font=("Segoe UI", 13, "bold"), bg="#c62828", fg="white", activebackground="#8e0000",
              bd=0, command=_stop.set).pack(fill="x", padx=4, pady=(4, 0))
    tk.Label(box, text="ERP download running - don't touch mouse / keyboard", font=("Segoe UI", 8), bg="#c62828",
             fg="white").pack(padx=4, pady=(0, 4))

    def poll():
        if ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000:  # Esc held down
            _stop_reason[0] = "pressed Esc"
            _stop.set()
        if _stop.is_set():
            root.destroy()
        else:
            root.after(50, poll)
    poll()
    root.mainloop()


def check():
    """Called between all steps: stops the run when STOP or Esc was pressed."""
    if _stop.is_set():
        raise Stopped(f"You {_stop_reason[0]}.")


def wait(seconds):
    end = time.time() + seconds
    while time.time() < end:
        check()
        time.sleep(0.05)


def fail(msg):
    raise Stopped(msg)


def front(w, tries=4):
    """Brings a window to the front. Windows refuses that when another program has just popped up in front
    (e.g. a WhatsApp message): then a zero-distance mouse move makes this script the last one with input,
    which Windows accepts, and it tries again."""
    for attempt in range(tries):
        try:
            w.set_focus()
            return w
        except Exception:
            if attempt == tries - 1:
                raise
            print(f"  (Windows kept the ERP from the front - try {attempt + 2} of {tries})", flush=True)
            ctypes.windll.user32.mouse_event(0x0001, 0, 0, 0, 0)  # MOUSEEVENTF_MOVE by 0, 0
            wait(1)


# ---- ERP helpers ------------------------------------------------------------------------------------
def desk(backend="win32"):
    from pywinauto import Desktop
    return Desktop(backend=backend)


def title(handle):
    """A window's title as Windows keeps it: unlike pywinauto's window_text() this does not ask the window, so it
    still reads while the ERP is busy (Not Responding) instead of coming back empty."""
    buf = ctypes.create_unicode_buffer(512)
    ctypes.windll.user32.GetWindowTextW(handle, buf, 512)
    return buf.value


def main_window(patience=10, busy_patience=180):
    """The ERP main window; while it redraws (a form opening) it can be missing for a moment, so look again for a while.
    A heavy report (e.g. [KLM]Sale details) can freeze the ERP for many seconds while it opens: then wait for it
    to answer again, and while the ERP program still runs give it longer before calling it closed."""
    start = time.time()
    while True:
        ws = [w for w in desk().windows() if title(w.handle).endswith(MAIN_TITLE) and w.class_name().startswith("WindowsForms10")]
        waited = time.time() - start
        if ws and not ctypes.windll.user32.IsHungAppWindow(ws[0].handle):
            return ws[0]
        if waited > busy_patience or (not ws and waited > patience and not erp_running()):
            return None
        check()
        time.sleep(0.5)


def erp_running():
    out = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {ERP_PROCESS}", "/NH"], capture_output=True, text=True).stdout
    return ERP_PROCESS.lower() in out.lower()


def message_boxes():
    """Error or warning boxes the ERP shows (standard Windows dialogs other than the save box)."""
    return [w for w in desk().windows() if w.is_visible() and w.class_name() == "#32770"]


def popups():
    """Drop-down lists and the like: ERP windows without a title of their own."""
    return [w for w in desk().windows() if w.is_visible() and w.class_name().startswith("WindowsForms10")
            and MAIN_TITLE not in w.window_text() and w.window_text() == ""]


def click_at(rect, dx=None):
    from pywinauto import mouse
    x = rect.right - 10 if dx is None else rect.left + dx
    mouse.click(coords=(x, (rect.top + rect.bottom) // 2))


def form(code):
    w = main_window() or fail("The ERP window closed." if not erp_running() else "The ERP stopped answering for 3 minutes.")
    fs = [c for c in w.descendants() if c.element_info.automation_id == code]
    return fs[0] if fs else None


def field(code, name):
    f = form(code) or fail(f"Form {code} is not open.")
    fs = [c for c in f.descendants() if c.element_info.automation_id == name]
    return fs[0] if fs else fail(f"Form {code} has no field '{name}' - the ERP screen changed.")


def login(account, cred, log, ours=False):
    """ours: the ERP open now was started by this run (the previous login), so it is closed without asking."""
    if erp_running() and ours:
        subprocess.run(["taskkill", "/IM", ERP_PROCESS, "/F"], capture_output=True)
        wait(3)
    if erp_running():
        from sales_report import ask
        if not ask("KolaoERP is already open.\n\nClose it and log in again for the download?\n"
                   "(Anything not saved in the ERP is lost.)"):
            fail("KolaoERP was already open and you chose not to close it.")
        subprocess.run(["taskkill", "/IM", ERP_PROCESS, "/F"], capture_output=True)
        wait(3)
    log(f"Starting KolaoERP, login {account} ({cred['user']})")
    os.startfile(ERP_SHORTCUT)
    lg = desk().window(title="Login", class_name_re="WindowsForms10.*")
    for _ in range(120):
        check()
        if lg.exists() and lg.is_visible():
            break
        time.sleep(0.5)
    else:
        fail("The ERP login window did not open within a minute.")
    wait(1)
    front(lg)
    lg.child_window(auto_id="id").set_edit_text(cred["user"])
    pwd = lg.child_window(auto_id="pwd")
    pwd.set_edit_text(cred["password"])
    pwd.set_focus()
    pwd.type_keys("{ENTER}")
    for _ in range(120):
        check()
        if message_boxes():
            box = message_boxes()[0]
            fail(f"Login {account} refused: {box.window_text()} {' '.join(c.window_text() for c in box.children())}")
        if main_window(patience=0, busy_patience=0) and not lg.exists():
            break
        time.sleep(0.5)
    else:
        fail(f"Login {account} did not finish within a minute (wrong password? run Setup.bat to enter it again).")
    wait(2)
    front(main_window())


def menu_bar():
    """The ERP window for UI automation, with its menu bar unfolded."""
    from pywinauto import mouse
    front(main_window())
    wait(0.3)
    w = desk("uia").window(title_re=f".*{MAIN_TITLE}").wrapper_object()
    shown = lambda: any(c.parent().window_text() == "MenuStrip1" and c.rectangle().width() > 0
                        for c in w.descendants(control_type="MenuItem"))
    if not shown():  # after login the ribbon is folded: the small arrow at the right end of its tab row unfolds it
        tabs = [c for c in w.descendants(control_type="Tab") if c.window_text() == "Ribbon Tabs"]
        if not tabs:
            fail("The ERP ribbon was not found.")
        r = tabs[0].rectangle()
        mouse.click(coords=(r.right - 16, (r.top + r.bottom) // 2))
        wait(1.5)
        if not shown():
            fail("The ERP menu bar did not open (top-right arrow of the ribbon).")
    return w


def menu(*names):
    """Clicks through the menu bar by name, e.g. menu('KLM', 'Vehicle Sales', 'Report Vehicles', '[KLM]Sale details').
    Each name is looked for only inside the item clicked before it: several menus have sub-menus of the same name."""
    from pywinauto import mouse
    w = menu_bar()
    level = [c for c in w.descendants(control_type="MenuItem") if c.parent().window_text() == "MenuStrip1"]
    for name in names:
        check()
        it = [c for c in level if c.window_text() == name]
        if not it:
            fail(f"Menu item '{name}' not found - the ERP menu changed.")
        r = it[0].rectangle()
        mouse.click(coords=((r.left + r.right) // 2, (r.top + r.bottom) // 2))
        wait(0.8)
        level = it[0].children(control_type="MenuItem")


def open_report(path, code):
    """Opens the report at that menu path and waits for its form."""
    menu(*path)
    for _ in range(60):
        check()
        if form(code):
            wait(1.5)
            return
        time.sleep(0.5)
    fail(f"'{' -> '.join(path)}' did not open.")


def close_report(code):
    """File -> Close closes the report in front (the one just exported), so the next one starts fresh."""
    menu("File", "Close")
    for _ in range(20):
        if not form(code):
            return
        wait(0.5)
    fail("The report did not close.")


def set_checks(code, name, rule):
    """Sets a drop-down check list and presses OK. Returns (ticked names, number of items).

    The list only shows the script the items scrolled into view, so it ticks '(All)' (which reaches every item), then
    scrolls through the whole list to untick / tick single items, and reads every item back the same way."""
    from pywinauto import mouse
    if rule == "all":
        want = lambda n: True
    else:
        how, names = rule
        want = (lambda n: n in names) if how == "only" else (lambda n: n not in names)
    front(main_window())  # the clicks below go to the screen: the ERP must be in front
    wait(0.2)
    click_at(field(code, name).rectangle())
    pop = None
    for _ in range(25):
        wait(0.2)
        ps = popups()
        if ps:
            pop = ps[0]
            break
    if pop is None:
        fail(f"The list '{name}' did not open.")
    u = desk("uia").window(handle=pop.handle).wrapper_object()
    boxes = lambda: {c.window_text(): c for c in u.descendants(control_type="CheckBox")}
    for _ in range(25):  # wait for the items
        if len(boxes()) > 1:
            break
        wait(0.2)
    else:
        fail(f"The list '{name}' is empty.")
    lst = [c for c in pop.children() if c.class_name().startswith("WindowsForms10.Window.8")]
    area = (lst[0] if lst else pop).rectangle()
    centre = ((area.left + area.right) // 2, (area.top + area.bottom) // 2)

    def scroll(steps):
        mouse.scroll(coords=centre, wheel_dist=steps)
        time.sleep(0.1)

    def top():  # a big wheel turn is not applied in full: scroll up in steps until '(All)' is in view
        for _ in range(60):
            if "(All)" in boxes():
                return
            scroll(10)
        fail(f"Could not scroll to the top of the list '{name}'.")

    def sweep(act=None):
        """Top to bottom through the whole list; act(name, box) on each item seen. Returns {name: state}."""
        top()
        seen, still = {}, 0
        while still < 2:
            check()
            new = False
            for n, c in boxes().items():
                if n == "(All)" or not n:
                    continue
                if act:
                    act(n, c)
                if n not in seen:
                    new = True
                seen[n] = c.get_toggle_state()
            still = 0 if new else still + 1
            scroll(-8)
        return seen

    def visible():
        return {n: c.get_toggle_state() for n, c in boxes().items() if n and n != "(All)"}

    state = None
    if rule == "all" or rule[0] == "except":
        # Showroom / Dealer / Car Type: (All) first. (All) ticks the whole list, so only the items in view are checked:
        # no scrolling to the end
        for _ in range(2):  # (All) ticks the list, or clears it when it was all ticked already: then once more
            top()
            boxes()["(All)"].click_input()
            wait(0.4)
            if all(v == 1 for v in visible().values()):
                break
        else:
            fail(f"'(All)' did not tick the list '{name}'.")

    def act(n, c):
        if c.get_toggle_state() != int(want(n)):
            c.toggle()
            time.sleep(0.05)
    if rule != "all":
        # Car Type: one pass down the list unticks the ones left out. Status: no (All) - the same pass ticks the
        # wanted ones and unticks any other. The pass reads every item back as it goes
        state = sweep(act)
    ok = [c for c in pop.children() if c.window_text() == "OK"]
    if not ok:
        fail(f"The list '{name}' has no OK button.")
    if state is None:  # 'all': (All) was pressed and checked
        ok[0].click_input()
        wait(0.8)
        return ["(All)"], "all"
    if rule[0] == "only" and not rule[1] <= set(state):
        fail(f"The list '{name}' has no {sorted(rule[1] - set(state))}.")
    wrong = [n for n, v in state.items() if v != int(want(n))]
    if not state or wrong:
        fail(f"Could not set the list '{name}': {wrong or 'no items'}.")
    ok[0].click_input()
    wait(0.8)
    return [n for n, v in state.items() if v], len(state)


def ticked(ctrl, box_left=False):
    """A DevExpress tick box shows no state to UI automation: its square is dark when ticked.
    The square sits at the right end of the control ('All' [x]) or, with box_left, the left end ([x] Open after saving)."""
    from PIL import ImageGrab
    r = ctrl.rectangle()
    x = r.left if box_left else r.right - 22
    box = ImageGrab.grab((x, r.top, x + 22, r.bottom)).convert("L")
    return sum(1 for p in box.getdata() if p < 90) >= 6


def tick(code, name, on=True):
    c = field(code, name)
    if ticked(c) != on:
        c.click_input()
        wait(0.6)
    if ticked(c) != on:
        fail(f"Could not {'tick' if on else 'untick'} '{name}'.")


def edit_box(code, name):
    boxes = [c for c in field(code, name).descendants() if "EDIT" in c.class_name()]
    return boxes[0] if boxes else fail(f"'{name}' has no text box.")


def pick(code, name, value):
    """A drop-down with one choice (NEW / USED): opens it, clicks the item by name, reads the box back."""
    front(main_window())
    wait(0.2)
    click_at(field(code, name).rectangle())
    wait(1)
    ps = popups()
    if not ps:
        fail(f"The list '{name}' did not open.")
    items = [c for c in desk("uia").window(handle=ps[0].handle).wrapper_object().descendants(control_type="ListItem")
             if c.window_text() == value]
    if not items:
        fail(f"The list '{name}' has no '{value}'.")
    items[0].click_input()
    wait(0.8)
    if edit_box(code, name).window_text() != value:
        fail(f"'{name}' did not take '{value}'.")


def day_for(rule):
    today = dt.date.today()
    if rule == "month_start":
        return today.replace(day=1)
    if rule == "year_ago":  # 29 Feb has no day a year back: 28 Feb
        return today.replace(year=today.year - 1, day=28) if (today.month, today.day) == (2, 29) else today.replace(year=today.year - 1)
    return today


def set_date(code, name, rule):
    """Types the date (M/D/YYYY, as the ERP shows it) and reads it back."""
    day = day_for(rule)
    text = f"{day.month}/{day.day}/{day.year}"
    box = edit_box(code, name)
    if not box.is_enabled():
        fail(f"The date '{name}' is greyed out.")
    box.click_input()
    wait(0.2)
    # a masked box filled part by part: Home, then month, day and year as digits (MMDDYYYY) - typed slashes or a
    # select-all are not understood by it
    box.type_keys("{HOME}")
    box.type_keys(f"{day.month:02}{day.day:02}{day.year}")
    box.type_keys("{TAB}")
    wait(0.6)
    if box.window_text() != text:
        fail(f"The date '{name}' shows '{box.window_text()}', not {text}.")
    return text


def responsive(w):
    """True when the ERP window answers messages (it hangs while a Query runs)."""
    res = ctypes.c_ulong()
    return bool(ctypes.windll.user32.SendMessageTimeoutW(w.handle, 0, 0, 0, 0x2, 1000, ctypes.byref(res)))


def query(log):
    w = main_window()
    btn = desk("uia").window(title_re=f".*{MAIN_TITLE}").child_window(title="Query(Q)", control_type="Button", found_index=0)
    btn.click_input()
    start = time.time()
    wait(1)
    calm_since = None
    while calm_since is None or time.time() - calm_since < 1.5:  # answering for 1.5 s in a row = done
        check()
        if message_boxes():
            box = message_boxes()[0]
            fail(f"The ERP showed a message after Query: {box.window_text()} {' '.join(c.window_text() for c in box.children())}")
        if time.time() - start > QUERY_TIMEOUT:
            fail(f"Query took longer than {QUERY_TIMEOUT // 60} minutes.")
        if not responsive(w):
            calm_since = None
        elif calm_since is None:
            calm_since = time.time()
        wait(0.25)
    log(f"  Query done in {time.time() - start:.0f} s")


def export(code, path, log=print):
    """Right-click the grid -> Export... -> Excel (Xls) ..., saved at path."""
    from pywinauto import mouse
    if path.exists():
        path.unlink()
    t0 = time.time()
    g = field(code, "Grid1").rectangle()
    mouse.right_click(coords=((g.left + g.right) // 2, g.top + 120))
    for _ in range(50):  # until the menu is there (not a fixed pause); a busy ERP can take a few seconds
        wait(0.2)
        menus = [e for e in desk("uia").windows() if e.element_info.control_type == "Menu"]
        item = [c for m in menus for c in m.descendants() if c.window_text() == "Export..."]  # by name: the same menu has Paste / Delete
        if item:
            break
    else:
        fail("The right-click menu has no 'Export...'.")
    item[0].click_input()
    box = desk().window(title="Export", class_name_re="WindowsForms10.*")
    for _ in range(40):
        check()
        if box.exists() and box.is_visible():
            break
        time.sleep(0.5)
    else:
        fail("The Export window did not open.")
    t1 = time.time()
    keep_closed = box.child_window(auto_id="chk_open")  # 'Open after saving' would start Excel: untick it first
    if ticked(keep_closed, box_left=True):
        keep_closed.click_input()
        wait(0.4)
    if ticked(keep_closed, box_left=True):
        fail("Could not untick 'Open after saving' in the Export window.")
    box.child_window(auto_id="btn_xls").click_input()  # 'Excel (Xls) ...': the .xls the report is built on
    for _ in range(50):  # until the save window is there
        wait(0.2)
        saves = message_boxes()
        if saves:
            wait(0.3)  # let it finish drawing before typing into it
            break
    else:
        fail("The save window did not open.")
    save = desk().window(handle=saves[0].handle)
    t2 = time.time()
    name_box = save.child_window(class_name="Edit", found_index=0)
    name_box.set_edit_text(str(path))
    wait(0.3)
    # Save given to the save window as its OK command: unlike a click on the button (BM_CLICK, often ignored while
    # the window is not the active one) it needs no focus, and nothing on top can block it
    ctypes.windll.user32.PostMessageW(save.handle, 0x0111, 1, 0)  # WM_COMMAND, IDOK
    for attempt in range(2):  # should it still not register: Enter in the file name box
        for _ in range(10):  # a press that works closes the window well within 2 s
            if not save.exists():
                break
            wait(0.2)
        else:
            if attempt == 0:
                front(save)
                name_box.set_focus()
                name_box.type_keys("{ENTER}")
                log("  (Save press did not register - pressed Enter)")
                continue
            fail("The save window did not close after Save and Enter.")
        break
    for _ in range(360):
        check()
        if path.exists() and path.stat().st_size > 0:
            break
        time.sleep(0.5)
    else:
        fail(f"{path.name} was not saved within 3 minutes.")
    t3 = time.time()
    size = -1
    while size != path.stat().st_size:  # written completely once the size stops changing
        size = path.stat().st_size
        wait(0.4)
    if box.exists():
        box.child_window(auto_id="btn_exit").click_input()
        wait(0.5)
    log(f"  export timing: open Export {t1 - t0:.1f}s, open save window {t2 - t1:.1f}s, "
        f"Save until file written {t3 - t2:.1f}s, closing {time.time() - t3:.1f}s")


def signatures():
    import sales_report as sr
    main = [p for p in sr.MAIN.glob("*.xls*") if not p.name.startswith("~$")]
    if len(main) != 1:
        fail(f"Main must hold exactly one workbook to check the files against (found {len(main)}).")
    return sr.main_signatures(main[0])


def part_of(path, sig):
    """(part sales_report.py sees in the file, its rows); (None, []) for a file it cannot read."""
    import sales_report as sr
    try:
        header, rows = sr.read_raw(path)
        return sr.classify(path, header, rows, sig), rows
    except Exception:
        return None, []


def verify(path, parts, sig, log):
    """The file must have rows and be one of the parts sales_report.py should see in it. Returns the part."""
    got, rows = part_of(path, sig)
    if not rows:
        fail(f"{path.name} has no rows.")
    if got not in parts:
        fail(f"{path.name} looks like {got or 'nothing the report knows'}, expected "
             + " or ".join(f"{k} {b}" for k, b in sorted(parts)) + ".")
    log(f"  OK {path.name}: {len(rows)} rows, recognised as {got[0]} - {got[1]}")
    return got


# ---- which downloads --------------------------------------------------------------------------------
def group_of(report):
    account, menu_path = report[0], report[1]
    if account == "DL":
        return "DL (dealers)"
    return {"KLM": "KLM (MBS)", "LXC": "LXC (Geely / BMW)"}.get(menu_path[0], "KDC + KDCX")


def choose_reports():
    """A window listing every download, all ticked: the person running it picks which to fetch.
    Returns the chosen REPORTS, or None for Cancel."""
    import tkinter as tk
    root = tk.Tk()
    root.title("ERP export - choose the data to download")
    root.attributes("-topmost", True)
    root.resizable(False, False)
    tk.Label(root, text="Tick the data to download from KolaoERP into Input:", font=("Segoe UI", 10, "bold")).grid(
        row=0, column=0, columnspan=4, padx=12, pady=(10, 6), sticky="w")
    marks, groups = [], {}
    for r in REPORTS:
        groups.setdefault(group_of(r), []).append(r)
    for col, (title, reports) in enumerate(groups.items()):
        box = tk.LabelFrame(root, text=title, font=("Segoe UI", 9, "bold"), padx=8, pady=4)
        box.grid(row=1, column=col, padx=6, pady=4, sticky="n")
        for r in reports:
            v = tk.BooleanVar(value=True)
            tk.Checkbutton(box, text=r[3].rsplit(".", 1)[0], variable=v, font=("Segoe UI", 9)).pack(anchor="w")
            marks.append((r, v))
    chosen = [None]

    def start():
        chosen[0] = [r for r, v in marks if v.get()]
        if not chosen[0]:
            tk.messagebox.showwarning("ERP export", "Tick at least one, or press Cancel.", parent=root)
            return
        root.destroy()
    import tkinter.messagebox  # noqa: F401  (used in start)
    bar = tk.Frame(root)
    bar.grid(row=2, column=0, columnspan=4, pady=(6, 10))
    for text, cmd in (("All", lambda: [v.set(True) for _, v in marks]), ("None", lambda: [v.set(False) for _, v in marks]),
                      ("Start", start), ("Cancel", root.destroy)):
        tk.Button(bar, text=text, width=10, command=cmd, font=("Segoe UI", 9, "bold" if text == "Start" else "normal")).pack(side="left", padx=4)
    root.bind("<Return>", lambda e: start())
    root.eval("tk::PlaceWindow . center")
    root.mainloop()
    return chosen[0]


# ---- run --------------------------------------------------------------------------------------------
def run(log, reports=REPORTS):
    cred = json.loads(CREDENTIALS.read_text(encoding="utf-8")) if CREDENTIALS.exists() else fail(
        f"{CREDENTIALS.name} is missing: run Setup.bat to enter the ERP logins.")
    INPUT.mkdir(exist_ok=True)
    sig = signatures()
    saved, logged_in = [], False
    for account in ACCOUNTS:
        if not any(r[0] == account for r in reports):  # nothing chosen for this login: not logged in at all
            continue
        c = cred.get(account) or {}
        if not c.get("user") or not c.get("password"):
            fail(f"No ERP login for {LOGINS.get(account, account)} on this PC: run Setup.bat to enter it.")
        login(account, c, log, ours=logged_in)
        logged_in = True
        for acc, menu_path, code, name, parts, filters in reports:
            if acc != account:
                continue
            check()
            t_report = time.time()
            log(f"{' > '.join(menu_path)}:")
            open_report(menu_path, code)
            for f in filters:
                if f[0] == "checks":
                    on, total = set_checks(code, f[1], f[2])  # read back after ticking
                    log(f"  {f[1]}: (All) ticked" if total == "all" else f"  {f[1]}: {len(on)} of {total} ticked" + (f" ({', '.join(on)})" if len(on) <= 4 else "")
                        + (f" - not ticked: {', '.join(sorted(f[2][1]))}" if f[2] != "all" and f[2][0] == "except" else ""))
                elif f[0] == "pick":
                    pick(code, f[1], f[2])
                    log(f"  {f[1]}: {f[2]}")
                elif f[0] == "date":
                    log(f"  {f[1]}: {set_date(code, f[1], f[2])}")
                else:
                    tick(code, f[1], f[0] == "tick")
                    log(f"  {f[1]}: {f[0]}ed")
            query(log)
            path = INPUT / name
            tmp = INPUT / f"~erp {name}"  # saved under a temporary name, put in place only once it checks out
            export(code, tmp, log)
            verify(tmp, parts, sig, log)
            # what it replaces: the same name, or an older download of the same part under another name
            from sales_report import recycle
            old = [p for p in INPUT.iterdir() if p.suffix.lower() in (".xls", ".xlsx") and p != tmp and p not in saved
                   and not p.name.startswith("~") and (p.stem.lower() == path.stem.lower() or part_of(p, sig)[0] in parts)]
            if old:
                recycle(old)
                log(f"  old {', '.join(p.name for p in old)} -> Recycle Bin")
            tmp.rename(path)
            saved.append(path)
            log(f"  saved {path}")
            # closed once saved: left open, every report adds its grid to the ERP window that each later step
            # searches through, and the run got slower with each one
            try:
                close_report(code)
            except Stopped:
                if _stop.is_set():
                    raise
                log("  (could not close the report - carrying on with it open)")
            log(f"  {time.time() - t_report:.0f} s for this report")


def main():
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    lines = []

    def log(msg):
        print(msg, flush=True)
        lines.append(msg)

    try:  # real screen pixels for every window, so the STOP box lands where it is meant to
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except OSError:
        pass
    # which downloads: all of them with --all (e.g. for a scheduled run); --only DL,KLM and / or --types Sales,Stock
    # (Diana's picker: groups as in the choice window, by their first word, and the first word of the file name);
    # otherwise the choice window
    def listed(name):
        return sys.argv[sys.argv.index(name) + 1].split(",") if name in sys.argv[:-1] else None
    who, what = listed("--only"), listed("--types")
    if "--all" in sys.argv:
        reports = REPORTS
    elif who or what:
        reports = [r for r in REPORTS if (not who or group_of(r).split()[0] in who)
                   and (not what or r[3].split()[0] in what)]
        if not reports:
            print(f"Nothing matches --only {who} --types {what} - nothing downloaded.")
            return None
    else:
        reports = choose_reports()
    if reports is None:
        print("ERP export cancelled - nothing downloaded.")
        return None
    threading.Thread(target=_stop_window, daemon=True).start()
    log(f"ERP export {stamp}  (STOP window or Esc stops it)")
    log("Chosen: " + ("all" if len(reports) == len(REPORTS) else ", ".join(r[3].rsplit(".", 1)[0] for r in reports)))
    ok = False
    try:
        run(log, reports)
        ok = True
        log("")
        log(f"Done: {len(reports)} files in {INPUT}")
        subprocess.run(["taskkill", "/IM", ERP_PROCESS, "/F"], capture_output=True)  # finished: close the ERP
        log("KolaoERP closed.")
    except Stopped as e:
        log("")
        log(f"STOPPED: {e}")
    except Exception as e:  # anything unexpected stops too, with the reason
        log("")
        log(f"STOPPED (unexpected): {type(e).__name__}: {e}")
    finally:
        _stop.set()
        for p in INPUT.glob("~erp *"):  # half-finished downloads never stay in Input
            p.unlink(missing_ok=True)
        (ROOT / "output").mkdir(exist_ok=True)
        (ROOT / "output" / f"ERP export {stamp}.txt").write_text("\n".join(lines), encoding="utf-8-sig")
    return ok


LOGINS = {"DL": "the dealers' data (DL)",
          "KDC": "car sales, stock and order data (KDC / KDCX, KLM, LXC)"}


def ask_logins():
    """Setup.bat: asks for the ERP logins of this PC's user and keeps them in credentials.json (only on this PC)."""
    from getpass import getpass
    cred = json.loads(CREDENTIALS.read_text(encoding="utf-8")) if CREDENTIALS.exists() else {}
    if all((cred.get(a) or {}).get("user") and (cred.get(a) or {}).get("password") for a in LOGINS):
        print("      ERP logins already saved on this PC:")
        for a, what in LOGINS.items():
            print(f"        {cred[a]['user']}  for {what}")
        if input("      Keep them? (Y = keep, N = enter new ones): ").strip().lower() != "n":
            print("      kept.")
            return
    print("      The ERP download logs in to KolaoERP with your own ERP accounts.")
    print("      They are saved on this PC only (credentials.json) and never uploaded. Passwords do not show while typed.")
    print("      Leave a user ID empty to skip it (run Setup.bat again later to enter it).")
    for a, what in LOGINS.items():
        print()
        user = input(f"      ERP user ID for {what}: ").strip()
        if not user:
            old = (cred.get(a) or {}).get("user")
            print(f"      kept {old}." if old else "      skipped.")
            continue
        while True:
            password = getpass("      password: ")
            if password and password == getpass("      password again: "):
                break
            print("      the two passwords are not the same (or empty) - again please.")
        cred[a] = {"user": user, "password": password}
    cred["_note"] = "KolaoERP logins for erp_export.py. Stays on this PC (in .gitignore) - never upload it."
    CREDENTIALS.write_text(json.dumps(cred, indent=2, ensure_ascii=False), encoding="utf-8")
    missing = [what for a, what in LOGINS.items() if not (cred.get(a) or {}).get("user")]
    print()
    print("      saved." + (f" Not entered: {', '.join(missing)} - the ERP download cannot fetch that data yet." if missing else ""))


if __name__ == "__main__":
    if "--logins" in sys.argv:
        ask_logins()
        sys.exit(0)
    done = main()
    sys.exit(0 if done else 2 if done is None else 1)   # 2: cancelled, nothing downloaded (not a failure)
