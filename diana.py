"""Diana The KOLAO sale department reporter bot.

Runs the daily job (erp_export.py, then sales_report.py) and shows Diana, a little 8-bit anime office girl, at her
desk: she walks in and works while the job runs, looks shocked if a step fails and gives a thumbs up when it is complete.
The scripts' text still goes to the black window. Try the look without running anything:  diana.py --demo  or  --demo fail  (add a number to try that outfit, or auto, then an hour 0-23 and clear / cloudy / rain / storm / cold)
"""
import ctypes
import datetime as dt
import random
import subprocess
import sys
import threading
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
STEPS = [("Downloading ERP files...", "erp_export.py",
          "ERP download stopped - the report was NOT run. See the reason above."),
         ("Building the report...", "sales_report.py", "The report stopped with an error. See the reason above.")]

W, H = 75, 50     # scene size in 8-bit pixels (300 x 200 at 100% screen scale)
TICK = 120        # ms per animation frame
AUTO_CLOSE = 20   # seconds the COMPLETE scene stays up

COLORS = {"k": "#1b1b2f", "s": "#ffe0c7", "h": "#2e2a5e", "l": "#6f6bd6", "e": "#3d8bfd", "w": "#ffffff",
          "c": "#ff9eb5", "m": "#d9455f", "b": "#4f8ef7", "r": "#e63946", "p": "#3a3a55", "y": "#ffd23f",
          "d": "#8fd8ff", "o": "#ff5d8f", "q": "#3a3a75"}

# ---- Diana: a chibi anime girl, 16 pixels wide, big head, big shiny eyes, long hair --------------------------
# Outfit colours: b = top, w = collar / neckline, r = accent (tie, sash, print), q = skirt / trousers,
# g = shoes or hem, o = hair accessory
EYES = {"open": ("kkk", "wee", "eee"), "blink": ("sss", "kkk", "sss"), "shock": ("kkk", "wkw", "kkk"),
        "happy": ("sks", "ksk", "sss")}
MOUTH = {"open": "sssssmmsssss", "blink": "sssssmmsssss", "shock": "ssssskksssss", "happy": "ssssmmmmssss"}
_HAIR_LOW = [".hhhhlhhhhhhhhh.", ".hhhhhhhhhhhhhh.", "hhhsshhhsshhhshh"]
HAIRS = {"bow": ["....hhhhhhoo.oo.", "...hhhhhhhhooo..", "..hhhhhhhhhhhh.."] + _HAIR_LOW,
         "bun": ["......hhhh.o....", "...hhhhhhhoyo...", "..hhhhhhhhhohh.."] + _HAIR_LOW,   # bun with a flower
         "cap": ["....oooooooo....", "...ooooyoooooo..", "..hhhhhhhhhhoooo"] + _HAIR_LOW,
         "band": ["....hhhhhhhh....", "...hhhhhhhhhh...", "..oooooooooooo.."] + _HAIR_LOW,
         "beret": ["...oooooooo.....", "..oooooooooo.o..", "..hhhhhhhhhhhh.."] + _HAIR_LOW}
TOPS = {"sailor": ["hh.wwbbbbbbww.hh", "h.sbwwbrrbwwbs.h", "..sbbbbrrbbbbs..", "...bbbbbbbbbb..."],
        "lao": ["hh.rrbbbbbbbb.hh", "h.sbrrbbbbbbbs.h", "..sbbbrrbbbbbs..", "...bbbbbrrbbb..."],     # pha biang sash
        "hanbok": ["hh.bbwbbbbwbb.hh", "h.sbbbbrrbbbbs.h", "..sqqqqrrqqqqs..", "..qqqqqqrqqqqq.."],  # jeogori + ribbon
        "tee": ["hh.bbbbbbbbbb.hh", "h.sbbbrbbrbbbs.h", "..sbbbbrrbbbbs..", "...qqqqqqqqqq..."],
        "hoodie": ["hhbbbbbbbbbbbbhh", "h.sbbbwbbwbbbs.h", "..sbbbbbbbbbbs..", "...brrrrrrrrb..."],
        "dress": ["hh.bbbbbbbbbb.hh", "h.sbbbbbbbbbbs.h", "..srrrrrrrrrrs..", "...bbbbbbbbbb..."],
        "blazer": ["hh.bbbwwwwbbb.hh", "h.sbbbwrrwbbbs.h", "..sbbbbwwbbbbs..", "...bbbbbbbbbb..."],
        "track": ["hh.bbbbwwbbbb.hh", "h.sbbbbrrbbbbs.h", "..sbbbbrrbbbbs..", "...bbbbbbbbbb..."],
        "overalls": ["hh.bqbbbbbbqb.hh", "h.sbqbbbbbbqbs.h", "..sbqqqqqqqqbs..", "...qqqqqqqqqq..."],
        "raincoat": ["hhbbbbbbbbbbbbhh", "h.sbbbbrbbbbbs.h", "..sbbbbbbbbbbs..", "..bbbbbrbbbbbb.."],
        "sweater": ["hh.rrrrrrrrrr.hh", "h.sbrrbbbbbbbs.h", "..sbbrrbbbbbbs..", "...bbbbbbbbbb..."]}   # + scarf
BOTTOMS = {  # standing, walking step a, walking step b
    "skirt": (["..qqqqqqqqqqqq..", "....ss....ss....", "...kkk....kkk..."],
              ["..qqqqqqqqqqqq..", "...ss......ss...", "..kkk......kkk.."],
              ["..qqqqqqqqqqqq..", "......ssss......", ".....kkkkkk....."]),
    "flared": (["..qqqqqqqqqqqq..", "....ss....ss....", "...ggg....ggg..."],
               ["..qqqqqqqqqqqq..", "...ss......ss...", "..ggg......ggg.."],
               ["..qqqqqqqqqqqq..", "......ssss......", ".....gggggg....."]),
    "boots": (["..qqqqqqqqqqqq..", "....kk....kk....", "...kkk....kkk..."],
              ["..qqqqqqqqqqqq..", "...kk......kk...", "..kkk......kkk.."],
              ["..qqqqqqqqqqqq..", "......kkkk......", ".....kkkkkk....."]),
    "wellies": (["..qqqqqqqqqqqq..", "....gg....gg....", "...ggg....ggg..."],   # raincoat hem + rain boots
                ["..qqqqqqqqqqqq..", "...gg......gg...", "..ggg......ggg.."],
                ["..qqqqqqqqqqqq..", "......gggg......", ".....gggggg....."]),
    "pencil": (["...qqqqqqqqqq...", "....ss....ss....", "...kkk....kkk..."],
               ["...qqqqqqqqqq...", "...ss......ss...", "..kkk......kkk.."],
               ["...qqqqqqqqqq...", "......ssss......", ".....kkkkkk....."]),
    "jeans": (["....qqqqqqqq....", "....qqq..qqq....", "...ggg....ggg..."],
              ["....qqqqqqqq....", "...qqq....qqq...", "..ggg......ggg.."],
              ["....qqqqqqqq....", ".....qqqqqq.....", ".....gggggg....."]),
    "sport": (["....qqqqqqqq....", "....qrq..qrq....", "...ggg....ggg..."],
              ["....qqqqqqqq....", "...qrq....qrq...", "..ggg......ggg.."],
              ["....qqqqqqqq....", ".....qrqqrq.....", ".....gggggg....."]),
    "sinh": (["...qqqqqqqqqq...", "...gqgqgqgqgq...", "....kkk..kkk...."],
             ["...qqqqqqqqqq...", "..gqgqgqgqgqgq..", "..kkk......kkk.."],
             ["...qqqqqqqqqq...", "...gqgqgqgqgq...", ".....kkkkkk....."]),
    "chima": ([".qqqqqqqqqqqqqq.", "qqqqqqqqqqqqqqqq", "gggggggggggggggg"],
              ["..qqqqqqqqqqqqqq", ".qqqqqqqqqqqqqqq", ".ggggggggggggggg"],
              ["qqqqqqqqqqqqqq..", "qqqqqqqqqqqqqqq.", "ggggggggggggggg."])}
THUMB = ["s..", "s..", "sss", "sss"]  # raised fist, drawn beside the head
UMBRELLA = [".....uuuuuu.....", "...uuuuuuuuuu...", "..uuuuuuuuuuuu..", ".uuuuuuuuuuuuuu.", "u.u.u.u.u.u.u.u."]

# She wears the next outfit every day and starts over after the last one: (name, top, bottom, hair, colours)
OUTFITS = [
    ("navy sailor uniform", "sailor", "skirt", "bow", {"b": "#ffffff", "w": "#2e3a75", "r": "#e63946", "q": "#2e3a75"}),
    ("white tee & jeans", "tee", "jeans", "band", {"b": "#ffffff", "w": "#ffffff", "r": "#ff5d8f", "q": "#4a6fa5", "o": "#ff5d8f"}),
    ("pink Lao sinh", "lao", "sinh", "bun", {"b": "#ff9ecb", "w": "#ff9ecb", "r": "#ffd23f", "q": "#c2457a", "g": "#ffd23f", "o": "#ffffff"}),
    ("grey hoodie", "hoodie", "jeans", "cap", {"b": "#9aa0ad", "w": "#ffffff", "r": "#7d8390", "q": "#3a4a6b", "o": "#2e3a75"}),
    ("red dress", "dress", "flared", "bow", {"b": "#e63946", "w": "#e63946", "r": "#ffffff", "q": "#e63946", "g": "#1b1b2f", "o": "#ffffff"}),
    ("navy office suit", "blazer", "pencil", "band", {"b": "#2e3a75", "w": "#ffffff", "r": "#4f8ef7", "q": "#2e3a75", "o": "#2e3a75"}),
    ("green hanbok", "hanbok", "chima", "bow", {"b": "#7bd389", "w": "#ffffff", "r": "#e63946", "q": "#ff8fab", "g": "#ffd23f", "o": "#e63946"}),
    ("blue tracksuit", "track", "sport", "cap", {"b": "#4f8ef7", "w": "#ffffff", "r": "#ffffff", "q": "#4f8ef7", "o": "#ffffff"}),
    ("denim overalls", "overalls", "jeans", "beret", {"b": "#ffd166", "w": "#ffd166", "q": "#4a6fa5", "o": "#e63946"}),
    ("cream sweater", "sweater", "boots", "beret", {"b": "#f3e5c8", "w": "#f3e5c8", "r": "#c0392b", "q": "#7a5c3e", "o": "#c0392b"}),
    ("pink tee & jeans", "tee", "jeans", "bow", {"b": "#ff9ecb", "w": "#ff9ecb", "r": "#ffffff", "q": "#4a6fa5", "o": "#ffffff"}),
    ("summer sailor", "sailor", "skirt", "band", {"b": "#ffffff", "w": "#4f8ef7", "r": "#ff5d8f", "q": "#4f8ef7", "o": "#4f8ef7"}),
    ("lilac hoodie", "hoodie", "skirt", "bow", {"b": "#c7a8f0", "w": "#ffffff", "r": "#a98bd6", "q": "#5a3a8c", "o": "#ffffff"}),
    ("sky-blue dress", "dress", "flared", "band", {"b": "#8fd8ff", "w": "#8fd8ff", "r": "#ffffff", "q": "#8fd8ff", "g": "#ffffff", "o": "#ffffff"}),
    ("beige office suit", "blazer", "pencil", "bun", {"b": "#d9b99b", "w": "#ffffff", "r": "#7a5c3e", "q": "#d9b99b", "o": "#ff9eb5"}),
    ("blue Lao sinh", "lao", "sinh", "bun", {"b": "#4f8ef7", "w": "#4f8ef7", "r": "#ffd23f", "q": "#1f3a7a", "g": "#ffd23f", "o": "#ff5d8f"}),
    ("yellow tee & jeans", "tee", "jeans", "cap", {"b": "#ffd166", "w": "#ffd166", "r": "#f28c28", "q": "#4a6fa5", "o": "#f28c28"}),
    ("red tracksuit", "track", "sport", "band", {"b": "#e63946", "w": "#ffffff", "r": "#ffffff", "q": "#e63946", "o": "#ffffff"}),
    ("mint hoodie", "hoodie", "jeans", "beret", {"b": "#a8e6cf", "w": "#ffffff", "r": "#7cc9ad", "q": "#2f4858", "o": "#2f4858"}),
    ("purple hanbok", "hanbok", "chima", "bow", {"b": "#b38cf0", "w": "#ffffff", "r": "#ff5d8f", "q": "#4f8ef7", "g": "#ffffff", "o": "#ff5d8f"}),
    ("black dress", "dress", "flared", "bow", {"b": "#2b2b3a", "w": "#2b2b3a", "r": "#ffd23f", "q": "#2b2b3a", "g": "#e63946", "o": "#e63946"}),
    ("yellow overalls", "overalls", "jeans", "band", {"b": "#ffffff", "w": "#ffffff", "q": "#ffd166", "o": "#ffd166"}),
    ("red winter sweater", "sweater", "boots", "beret", {"b": "#e63946", "w": "#e63946", "r": "#ffffff", "q": "#2b2b3a", "o": "#ffffff"}),
    ("black tee & jeans", "tee", "jeans", "cap", {"b": "#2b2b3a", "w": "#2b2b3a", "r": "#ff5d8f", "q": "#4a6fa5", "o": "#2b2b3a"}),
    ("pink office suit", "blazer", "pencil", "bow", {"b": "#ff9ecb", "w": "#ffffff", "r": "#e63946", "q": "#ff9ecb", "o": "#ffffff"}),
    ("pink sailor uniform", "sailor", "skirt", "bow", {"b": "#ffffff", "w": "#ff7aa8", "r": "#ff7aa8", "q": "#ff7aa8", "o": "#ff7aa8"}),
    ("peach dress", "dress", "flared", "beret", {"b": "#ffb38a", "w": "#ffb38a", "r": "#ffffff", "q": "#ffb38a", "g": "#ffffff", "o": "#ffffff"}),
    ("gold Lao sinh", "lao", "sinh", "bun", {"b": "#ffffff", "w": "#ffffff", "r": "#e0a800", "q": "#b8860b", "g": "#e63946", "o": "#ffd23f"}),
    ("black tracksuit", "track", "sport", "cap", {"b": "#2b2b3a", "w": "#ffffff", "r": "#ff5d8f", "q": "#2b2b3a", "o": "#ff5d8f"}),
    ("green sweater", "sweater", "skirt", "band", {"b": "#7bd389", "w": "#7bd389", "r": "#ffd23f", "q": "#3a3a55", "o": "#ffd23f"}),
    ("red hanbok", "hanbok", "chima", "bow", {"b": "#ffd166", "w": "#ffffff", "r": "#4f8ef7", "q": "#e63946", "g": "#ffd23f", "o": "#e63946"}),
    ("pink hoodie", "hoodie", "skirt", "cap", {"b": "#ff9ecb", "w": "#ffffff", "r": "#ff7aa8", "q": "#3a3a55", "o": "#ffffff"}),
    ("yellow raincoat", "raincoat", "wellies", "cap", {"b": "#ffd23f", "w": "#ffd23f", "r": "#e0a800", "q": "#ffd23f", "g": "#e63946", "o": "#ffd23f"}),
    ("pink raincoat", "raincoat", "wellies", "cap", {"b": "#ff9ecb", "w": "#ff9ecb", "r": "#ffffff", "q": "#ff9ecb", "g": "#ffd23f", "o": "#ff9ecb"}),
    ("sky-blue raincoat", "raincoat", "wellies", "cap", {"b": "#8fd8ff", "w": "#8fd8ff", "r": "#ffffff", "q": "#8fd8ff", "g": "#ffd23f", "o": "#8fd8ff"}),
    ("white summer dress", "dress", "flared", "band", {"b": "#ffffff", "w": "#ffffff", "r": "#8fd8ff", "q": "#ffffff", "g": "#e8c9a6", "o": "#8fd8ff"}),
]


def posed(top, pose):
    """The top with her arms down (calm), both arms in the air (arms_up) or the right arm raised (thumb)."""
    if pose == "calm":
        return top
    r, sleeve = [list(row) for row in top], top[1][3]
    r[1][13] = r[2][13] = "."
    r[0][14] = r[1][14] = "s"
    r[1][13] = sleeve
    if pose == "arms_up":
        r[1][2] = r[2][2] = "."
        r[0][0] = r[1][0] = r[0][15] = r[1][15] = "s"
        r[0][14] = "."
        r[1][1] = r[1][2] = r[1][14] = sleeve
    return ["".join(row) for row in r]


def sprite(top="sailor", bottom="skirt", hair="bow", eyes="open", body="calm", legs="stand"):
    e = EYES[eyes]
    face = ["hhs" + e[i] + "ssss" + e[i] + "shh" for i in range(3)]
    step = {"stand": 0, "walk_a": 1, "walk_b": 2}[legs]
    return (HAIRS[hair] + face + ["hhsccssssssccshh", "hh" + MOUTH[eyes] + "hh", "hh.ssssssssss.hh", "hh..wwssssww..hh"]
            + posed(TOPS[top], body) + BOTTOMS[bottom][step])


# Weather outside her window: Vientiane, from Open-Meteo (free, no account). Only the city's position is sent.
WEATHER_URL = ("https://api.open-meteo.com/v1/forecast?latitude=17.97&longitude=102.61&current=weather_code,temperature_2m"
               "&timezone=Asia%2FBangkok")


def fetch_weather():
    """(weather, temperature): weather is 'storm', 'rain', 'cloudy' or 'clear'.
    (None, None) without internet: then the window shows a plain sky and any outfit goes."""
    import json
    import urllib.request
    try:
        now = json.load(urllib.request.urlopen(WEATHER_URL, timeout=3))["current"]
        code, temp = now["weather_code"], now["temperature_2m"]
    except Exception:
        return None, None
    if code >= 95:
        return "storm", temp
    if 51 <= code <= 67 or 80 <= code <= 82:
        return "rain", temp
    return ("cloudy" if code in (2, 3, 45, 48) or 71 <= code <= 77 else "clear"), temp


def time_of_day(hour):
    return "morning" if 5 <= hour < 8 else "day" if hour < 16 else "evening" if hour < 19 else "night"


# Her drink, 4 pixels wide, standing on the desk: (name, rows, colours, hot). Hot ones steam, cold ones drip.
DRINKS = [
    ("bubble tea", ["..k.", "..k.", "xxxx", "tttt", "tptt", "ptpt"], {"x": "#e8f6ff", "t": "#d9a066", "p": "#3b2314"}, False),
    ("iced coffee", ["..k.", "..k.", "cxcc", "cccx", "xccc", "cccc"], {"x": "#e8f6ff", "c": "#8a5a3c"}, False),
    ("Lao iced tea", ["..k.", "..k.", "xxxx", "tttt", "ttxt", "tttt"], {"x": "#fff4e0", "t": "#ff8c42"}, False),
    ("coconut", [".k..", ".k..", "xxgg", "gggg", "gggg", ".gg."], {"x": "#fff8e7", "g": "#6ab04c"}, False),
    ("strawberry milk", ["..k.", "..k.", "xxxx", "mmmm", "mmmm", "mmmm"], {"x": "#ffffff", "m": "#ffb3c6"}, False),
    ("coffee", ["ccc.", "uuuu", "uuuu", "uuu."], {"c": "#5a3825", "u": "#ff9eb5"}, True),
    ("hot cocoa", ["cxc.", "uuuu", "uuuu", "uuu."], {"c": "#7b4a2e", "x": "#ffffff", "u": "#ffffff"}, True),
    ("green tea", ["ggg.", "uuuu", "uuuu", "uuu."], {"g": "#9ccc65", "u": "#e8f6ff"}, True),
]


def drink_today(weather=None, temp=None, hour=None):
    """Something hot when it rains, when it is cool or early in the morning; something cold on a warm day."""
    hour = time.localtime().tm_hour if hour is None else hour
    hot = weather in ("rain", "storm") or (temp is not None and temp < COOL) or 5 <= hour < 8
    fits = [i for i, d in enumerate(DRINKS) if d[3] == hot]
    return fits[dt.date.today().toordinal() % len(fits)]


# Which tops suit the weather: rain, cool (under COOL degrees) or warm. Without a forecast: any outfit.
RAIN_TOPS = {"raincoat", "hoodie", "track"}
COOL_TOPS = {"sweater", "hoodie", "track", "blazer", "hanbok", "raincoat"}
WARM_SKIP = {"sweater", "hoodie", "raincoat"}
COOL = 22


def outfit_today(weather=None, temp=None):
    """A different outfit every day, picked from the ones that suit today's weather, in list order."""
    if weather in ("rain", "storm"):
        fits = [i for i, o in enumerate(OUTFITS) if o[1] in RAIN_TOPS]
    elif temp is not None and temp < COOL:
        fits = [i for i, o in enumerate(OUTFITS) if o[1] in COOL_TOPS]
    elif weather is not None:
        fits = [i for i, o in enumerate(OUTFITS) if o[1] not in WARM_SKIP]
    else:
        fits = list(range(len(OUTFITS)))
    return fits[dt.date.today().toordinal() % len(fits)]


FLOOR = 40
DESK_X, DESK_TOP, DESK_W = 35, 31, 39
SEAT_X = 42                  # where Diana works behind the desk
SIT_Y = DESK_TOP - 14        # top of her hair when she sits at the desk
STAND_Y = FLOOR - 20         # top of Diana's hair when she stands on the floor
DONE_X = 12                  # where she stands for the thumbs up


class Job(threading.Thread):
    """Runs the steps one after the other; .step is the caption, .result becomes "ok" or "fail" at the end."""

    def __init__(self, demo=None):
        super().__init__(daemon=True)
        self.demo, self.step, self.result, self.code = demo, STEPS[0][0], None, 0

    def run(self):
        for caption, script, failed in STEPS:
            self.step = caption
            if self.demo:
                time.sleep(4)
                code = 1 if self.demo == "fail" and script == STEPS[-1][1] else 0
            else:
                code = subprocess.run([sys.executable, "-X", "utf8", str(HERE / script)]).returncode
            if code:
                print(f"\n{failed}\n", flush=True)
                self.code, self.result = code, "fail"
                return
        self.result = "ok"


class Scene:
    def __init__(self, job, day, hour=None, weather=None, temp=None):
        import tkinter as tk
        try:  # draw at the screen's real resolution so the pixels stay sharp on scaled (125%, 150%) screens
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        name, top, bottom, hair, colors = OUTFITS[day]
        self.job, self.outfit, self.look = job, name, (top, bottom, hair)
        self.colors = {**COLORS, "g": "#f4f4f4", **colors}
        self.root = root = tk.Tk()
        self.PX = max(4, round(4 * root.winfo_fpixels("1i") / 96))  # one 8-bit pixel in screen pixels
        root.title("Diana")
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        self.cv = tk.Canvas(root, width=W * self.PX, height=H * self.PX, highlightthickness=0, bd=0, bg="#111111")
        self.cv.pack()
        self.cv.bind("<Button-1>", self.clicked)
        root.geometry(f"{W * self.PX}x{H * self.PX}+{self.corner()}")
        root.update_idletasks()
        self.no_focus()
        self.state, self.t, self.x, self.ended_at = "walk", 0, -16, None
        self.papers = []
        self.fixed_hour, self.weather = hour, weather
        self.drink = drink_today(weather, temp, self.hour)

    @property
    def hour(self):
        return time.localtime().tm_hour if self.fixed_hour is None else self.fixed_hour

    # ---- window ------------------------------------------------------------------------------------
    def corner(self):
        """Bottom right of the screen, just above the taskbar: clear of the ERP windows the download clicks on."""
        class Rect(ctypes.Structure):
            _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]
        r, w, h = Rect(), W * self.PX, H * self.PX
        if ctypes.windll.user32.SystemParametersInfoW(0x30, 0, ctypes.byref(r), 0):  # SPI_GETWORKAREA
            return f"{r.right - w - 12}+{r.bottom - h - 12}"
        return f"{self.root.winfo_screenwidth() - w - 12}+{self.root.winfo_screenheight() - h - 60}"

    def no_focus(self):
        """Never take the keyboard from the ERP or the black window, and stay off the taskbar."""
        u = ctypes.windll.user32
        hwnd = u.GetParent(self.root.winfo_id()) or self.root.winfo_id()
        u.SetWindowLongW(hwnd, -20, u.GetWindowLongW(hwnd, -20) | 0x08000000 | 0x80)  # NOACTIVATE | TOOLWINDOW

    def clicked(self, _):
        if self.state in ("done", "fail"):
            self.root.destroy()

    # ---- drawing -------------------------------------------------------------------------------------
    def px(self, x, y, color, w=1, h=1):
        p = self.PX
        self.cv.create_rectangle(x * p, y * p, (x + w) * p, (y + h) * p, fill=color, width=0)

    def draw(self, rows, x, y, extra=None):
        colors = {**self.colors, **(extra or {})}
        for dy, row in enumerate(rows):
            for dx, c in enumerate(row):
                if c != ".":
                    self.px(x + dx, y + dy, colors[c])

    def room(self):
        sky = time_of_day(self.hour)
        self.px(0, 0, "#e6d2c0" if sky == "night" else "#fbe8d3", W, FLOOR)   # warm wall, a bit dimmer at night
        self.px(0, FLOOR - 2, "#e8c9a6", W, 2)    # skirting
        self.px(4, 5, "#c98b5a", 16, 12)          # window: the sky follows the time of day
        self.window(sky)
        self.clouds_and_rain()
        self.px(11, 6, "#c98b5a", 1, 10)
        self.px(24, 7, "#7bc96f", 4, 3)           # little plant on a shelf
        self.px(25, 10, "#c9744b", 2, 2)
        self.px(22, 12, "#c98b5a", 8, 1)
        self.px(0, FLOOR, "#c08a5a", W, 2)        # wooden floor
        self.px(0, FLOOR + 2, "#111111", W, H - FLOOR - 2)

    def window(self, sky):
        if sky == "morning":
            self.px(5, 6, "#ffd6a5", 14, 5)
            self.px(5, 11, "#ffb4a2", 14, 5)
            self.px(15, 12, "#ffef9f", 3, 2)        # sun just coming up
            self.px(16, 11, "#ffef9f", 1, 1)
        elif sky == "day":
            self.px(5, 6, "#9fdcff", 14, 10)
            self.px(7, 9, "#ffffff", 5, 2)          # cloud
            self.px(8, 8, "#ffffff", 3, 1)
            self.px(15, 7, "#ffe066", 2, 2)         # sun
        elif sky == "evening":
            self.px(5, 6, "#c77dff", 14, 3)
            self.px(5, 9, "#ff8fa3", 14, 3)
            self.px(5, 12, "#ffb347", 14, 4)
            self.px(7, 12, "#ff6b35", 3, 2)         # sun going down
            self.px(8, 11, "#ff6b35", 1, 1)
        else:
            self.px(5, 6, "#1d2350", 14, 10)
            self.px(15, 7, "#fff3b0", 2, 3)         # crescent moon
            self.px(14, 8, "#fff3b0", 1, 1)
            for i, (sx, sy) in enumerate([(6, 8), (9, 12), (13, 13), (7, 14), (17, 12), (12, 7)]):
                if (self.t + i) % 5:                 # twinkling stars
                    self.px(sx, sy, "#ffffff")

    def umbrella(self, x, y):
        top = self.colors["r"] if self.colors["r"] != self.colors["b"] else "#ff5d8f"
        self.draw(UMBRELLA, x + 4, y - 7, {"u": top})
        self.px(x + 13, y - 2, COLORS["k"], 1, 15)     # handle down to her hand

    def clouds_and_rain(self):
        w = self.weather
        if w not in ("cloudy", "rain", "storm"):
            return
        if w == "storm" and self.t % 23 in (0, 2):    # lightning
            self.px(5, 6, "#f5f5ff", 14, 10)
        grey = "#6c7383" if w != "cloudy" else "#c9ced8"
        self.px(5, 6, grey, 14, 3)                     # clouds hiding the sun / moon
        self.px(6, 9, grey, 5, 1)
        self.px(13, 9, grey, 5, 1)
        self.px(14, 6, "#8a92a3" if w != "cloudy" else "#e3e6ec", 4, 2)
        if w != "cloudy":
            for i in range(7):                         # falling raindrops
                x = 5 + (i * 5 + 2) % 14
                y = 9 + (self.t * 2 + i * 3) % 7
                self.px(x, y, "#a7d8ff", 1, 2)

    def desk_back(self, screen):
        self.px(59, 19, "#4b4b63", 12, 11)        # monitor
        self.px(60, 20, "#0e2a3d", 10, 9)
        if screen == "work":
            for i in range(4):
                self.px(61, 21 + 2 * i, "#7cf5c4", random.randint(2, 8), 1)
        elif screen == "fail":
            self.px(60, 20, "#7a0000" if self.t % 4 < 2 else "#e63946", 10, 9)
            for i in range(5):                     # an X on the screen
                self.px(63 + i, 22 + i, "#ffffff")
                self.px(67 - i, 22 + i, "#ffffff")
        else:
            for dx, dy in [(2, 4), (3, 5), (4, 6), (5, 5), (6, 4), (7, 3), (8, 2)]:  # a tick
                self.px(60 + dx, 20 + dy, "#7cf5c4")
        self.px(64, 30, "#4b4b63", 3, 1)
        self.px(36, 27, "#ffffff", 5, 4)          # pile of papers
        self.px(37, 28, "#c7c7d9", 3, 1)
        self.px(37, 30, "#c7c7d9", 3, 1)
        self.drink_on_desk()

    def drink_on_desk(self):
        name, rows, colors, hot = DRINKS[self.drink]
        x, y = 71, DESK_TOP - len(rows)
        self.draw(rows, x, y, colors)
        if hot:                                    # steam curling up
            for i in range(3):
                self.px(x + (i + self.t // 2) % 2 + (1 if i == 1 else 0), y - 2 - 2 * i, "#ffffff")
        elif self.t % 8 < 4:                       # a cold drop running down the glass
            self.px(x + 3, y + 3 + (self.t // 2) % 2, "#e8f6ff")

    def desk_front(self):
        self.px(DESK_X, DESK_TOP, "#a8693a", DESK_W, 1)
        self.px(DESK_X, DESK_TOP + 1, "#8b5a2b", DESK_W, FLOOR - DESK_TOP - 1)
        self.px(DESK_X + 3, DESK_TOP + 3, "#6b4220", DESK_W - 6, 1)
        self.px(DESK_X + 3, DESK_TOP + 6, "#6b4220", DESK_W - 6, 1)

    def caption(self, text, color="#ffffff"):
        self.cv.create_text(W * self.PX // 2, (FLOOR + 2 + (H - FLOOR - 2) / 2) * self.PX, text=text, fill=color,
                            font=("Consolas", 10, "bold"))

    def bubble(self, text, x, y, w, color):
        self.px(x, y, COLORS["k"], w, 9)
        self.px(x + 1, y + 1, "#ffffff", w - 2, 7)
        self.px(x + 15, y + 9, COLORS["k"], 3, 1)   # tail pointing down at Diana
        self.px(x + 16, y + 10, COLORS["k"], 1, 1)
        self.cv.create_text((x + w / 2) * self.PX, (y + 4.5) * self.PX, text=text, fill=color,
                            font=("Consolas", 12, "bold"))

    def sparkle(self, x, y, color=COLORS["y"]):
        self.px(x, y - 1, color)
        self.px(x - 1, y, color, 3, 1)
        self.px(x, y + 1, color)

    def heart(self, x, y):
        for dx, dy in [(0, 0), (2, 0), (0, 1), (1, 1), (2, 1), (1, 2)]:
            self.px(x + dx, y + dy, "#ff5d8f")

    # ---- animation ---------------------------------------------------------------------------------
    def frame(self):
        self.t += 1
        self.cv.delete("all")
        self.room()
        result = self.job.result
        if self.state == "walk":
            self.x += 2
            if self.x >= SEAT_X:
                self.x, self.state = SEAT_X, "work"
        elif self.state == "work" and result:
            self.state, self.t, self.ended_at = ("done" if result == "ok" else "fail"), 0, time.time()
        elif self.state == "done" and self.x > DONE_X and self.t > 3:
            self.x -= 2

        bob = self.t % 2   # a little bounce in her step
        if self.state == "walk":
            self.desk_back("work")
            self.draw(sprite(*self.look, legs="walk_a" if bob else "walk_b"), self.x, STAND_Y - bob)
            self.desk_front()
            if self.weather in ("rain", "storm"):
                self.umbrella(self.x, STAND_Y - bob)
            elif self.t < 12:
                self.heart(self.x + 14, STAND_Y - 4 - self.t // 3)
            self.caption(f"Diana~ {self.outfit} today!")
        elif self.state == "work":
            self.desk_back("work")
            self.draw(sprite(*self.look, "blink" if self.t % 25 == 0 else "open"), SEAT_X, SIT_Y)
            self.typing()
            self.desk_front()
            self.caption(self.job.step)
        elif self.state == "fail":
            jump = (0, -2, -3, -2)[self.t % 4] if self.t < 12 else 0
            y = STAND_Y - 3 + jump
            self.desk_back("fail")
            self.draw(sprite(*self.look, "shock", "arms_up"), SEAT_X, y)
            self.desk_front()
            self.px(SEAT_X + 15, y + 6 + (self.t // 2) % 3, COLORS["d"], 1, 2)   # sweat drop
            if self.t % 6 < 4:                     # a blinking "!"
                self.px(SEAT_X + 7, y - 9, "#ff3b30", 2, 5)
                self.px(SEAT_X + 7, y - 3, "#ff3b30", 2, 1)
            self.caption("Eh?! Check the black window", "#ff8fa3")
        else:
            self.desk_back("done")
            self.desk_front()
            walking = self.x > DONE_X
            self.draw(sprite(*self.look, "happy", "calm" if walking else "thumb",
                             ("walk_a" if bob else "walk_b") if walking else "stand"), self.x, STAND_Y - (bob if walking else 0))
            if not walking:
                self.draw(THUMB, self.x + 14, STAND_Y + 9 - (self.t // 3) % 2)
                self.bubble("COMPLETE!", 2, STAND_Y - 12, 38, "#2e9d4f")
                for i, (sx, sy) in enumerate([(46, 6), (3, 31), (33, 22), (52, 3)]):
                    if (self.t + i) % 4 < 2:
                        self.sparkle(sx, sy)
                if (self.t // 4) % 2:
                    self.heart(self.x + 17, STAND_Y + 3)
            self.caption("All done!  (click to close)", "#9cff9c")
            if time.time() - self.ended_at > AUTO_CLOSE:
                self.root.destroy()
                return
        self.root.after(TICK, self.frame)

    def typing(self):
        left = self.t % 2 == 0
        self.px(SEAT_X + (3 if left else 4), DESK_TOP - 1, COLORS["s"], 2, 1)   # hands tapping the keyboard
        self.px(SEAT_X + (11 if left else 10), DESK_TOP - 1, COLORS["s"], 2, 1)
        if self.t % 10 == 0:                      # a finished page flies off the pile
            self.papers.append([37, 25])
        for p in self.papers:
            p[0] -= 1
            p[1] -= 1
            self.px(p[0], p[1], "#ffffff", 3, 2)
        self.papers = [p for p in self.papers if p[1] > 1]

    def run(self):
        self.frame()
        self.root.mainloop()


def main():
    args = sys.argv[1:]
    demo = (args[1] if len(args) > 1 else "ok") if args[:1] == ["--demo"] else None
    job = Job(demo)
    job.start()   # the job never waits for the forecast
    # --demo ok 5 21 rain: outfit no. 5 at 9 pm in the rain; "auto" instead of 5 picks the outfit for that weather
    hour = int(args[3]) if demo and len(args) > 3 else None
    if demo and len(args) > 4:
        weather, temp = args[4], {"cold": 18}.get(args[4], 30)
        weather = "clear" if weather == "cold" else weather
    else:
        weather, temp = fetch_weather()
    pick = args[2] if demo and len(args) > 2 else "auto"
    day = outfit_today(weather, temp) if pick == "auto" else int(pick) % len(OUTFITS)
    try:
        Scene(job, day, hour, weather, temp).run()
    except Exception as e:                      # no window possible: the job still runs, just without Diana
        print(f"(Diana could not show her window: {e})", flush=True)
    job.join()
    sys.exit(job.code)


if __name__ == "__main__":
    main()
