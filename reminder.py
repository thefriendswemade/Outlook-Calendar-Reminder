"""Pops up an always-on-top reminder 1 minute before every Outlook (classic) calendar event.

Stays on screen until dismissed. Runs hidden in the background and can start at login
(the packaged .exe offers this on first run). Shows a red warning if it stops working.
"""
import datetime as dt
import os
import sys
import tkinter as tk
import traceback
import winreg

import pythoncom
import win32api
import win32com.client
import win32event
import winerror

LEAD = dt.timedelta(minutes=1)         # how early to pop up
LATE_GRACE = dt.timedelta(minutes=10)  # still alert if we missed it (e.g. PC woke from sleep)
WARN_AFTER = dt.timedelta(minutes=6)   # continuous failure before warning (Outlook can take minutes to start at login)
POLL_MS = 15_000
SNOOZE_MIN = 5
FMT = "%m/%d/%Y %I:%M %p"
FROZEN = getattr(sys, "frozen", False)  # True when running as the packaged .exe
APP_NAME = "Outlook Calendar Reminder"
RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
QUIT_EVENT = r"Global\OutlookCalendarReminderQuit"
if FROZEN:
    _dir = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "OutlookCalendarReminder")
    os.makedirs(_dir, exist_ok=True)
else:
    _dir = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(_dir, "reminder.log")


def message_box(text, yes_no=False):
    """Native Windows message box; returns True for Yes/OK."""
    import ctypes
    flags = 0x00040000 | (0x24 if yes_no else 0x40)  # topmost | (yes/no + question) or (OK + info)
    return ctypes.windll.user32.MessageBoxW(0, text, APP_NAME, flags) in (1, 6)


def startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY) as k:
            winreg.QueryValueEx(k, APP_NAME)
            return True
    except OSError:
        return False


def set_startup(on):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
        if on:
            winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, f'"{sys.executable}" --startup')
        else:
            try:
                winreg.DeleteValue(k, APP_NAME)
            except OSError:
                pass


def log(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except OSError:
        pass


def style_window(w):
    """Windows 11 rounded corners + drop shadow + subtle border on the frameless popup (best effort)."""
    try:
        import ctypes
        from ctypes import wintypes
        u32, dwm = ctypes.windll.user32, ctypes.windll.dwmapi
        hwnd = u32.GetParent(w.winfo_id()) or w.winfo_id()
        dwm.DwmSetWindowAttribute(hwnd, 33, ctypes.byref(ctypes.c_int(2)), 4)       # rounded corners
        dwm.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(ctypes.c_int(0x3b3226)), 4)  # border (BGR)
        class M(ctypes.Structure):
            _fields_ = [(n, ctypes.c_int) for n in ("l", "r", "t", "b")]
        dwm.DwmExtendFrameIntoClientArea(hwnd, ctypes.byref(M(1, 1, 1, 1)))        # shadow
    except Exception as e:
        log(f"style_window: {e}")


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.fired = set()
        self.cards = []
        self.ns = None
        self.fail_since = None
        self.warn_win = None
        self.root.report_callback_exception = self.on_error
        self.quit_event = win32event.CreateEvent(None, True, False, QUIT_EVENT)
        self.root.after(500, self.poll)
        self.root.after(1000, self.check_quit)

    def check_quit(self):
        if win32event.WaitForSingleObject(self.quit_event, 0) == win32event.WAIT_OBJECT_0:
            self.root.destroy()
            return
        self.root.after(1000, self.check_quit)

    # ---- Outlook access ----
    def calendar(self):
        if self.ns is None:
            # attach to the running Outlook only; never launch it ourselves
            self.ns = win32com.client.GetActiveObject("Outlook.Application").GetNamespace("MAPI")
        return self.ns.GetDefaultFolder(9)  # olFolderCalendar

    def upcoming(self):
        now = dt.datetime.now()
        items = self.calendar().Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        flt = "[Start] >= '{}' AND [Start] <= '{}'".format(
            (now - LATE_GRACE).strftime(FMT), (now + LEAD + dt.timedelta(minutes=1)).strftime(FMT))
        for it in items.Restrict(flt):
            if it.AllDayEvent or it.MeetingStatus in (5, 7):  # canceled
                continue
            # pywin32 labels Outlook's local wall-clock time as UTC; just drop the label
            start = it.Start.replace(tzinfo=None)
            yield it, start

    def poll(self):
        try:
            now = dt.datetime.now()
            for it, start in self.upcoming():
                key = (it.EntryID, start)
                if key in self.fired or now < start - LEAD:
                    continue
                self.fired.add(key)
                self.popup(it.Subject or "(no title)", start, it.Location or "")
            self.fail_since = None
            if self.warn_win is not None and self.warn_win.winfo_exists():
                self.warn_win.destroy()
        except Exception as e:  # Outlook closed/busy: retry next poll, warn if it persists
            self.ns = None
            now = dt.datetime.now()
            if self.fail_since is None:
                self.fail_since = now
                log(f"poll error: {e}")
            elif now - self.fail_since >= WARN_AFTER and not (
                    self.warn_win is not None and self.warn_win.winfo_exists()):
                self.warn(f"Can't read Outlook (is classic Outlook running?).\n{e}")
        self.root.after(POLL_MS, self.poll)

    # ---- UI ----
    def on_error(self, exc, val, tb):
        log("".join(traceback.format_exception(exc, val, tb)))
        self.warn(f"Calendar reminders hit an error:\n{val}")

    def warn(self, text):
        if self.warn_win is not None and self.warn_win.winfo_exists():
            self.warn_win.destroy()
        w = self.warn_win = tk.Toplevel(self.root)
        w.title("Calendar reminders NOT working")
        w.attributes("-topmost", True)
        w.configure(bg="#7f1d1d", padx=18, pady=14)
        tk.Label(w, text="⚠ Calendar reminders NOT working", font=("Segoe UI", 13, "bold"),
                 fg="white", bg="#7f1d1d").pack(anchor="w")
        tk.Label(w, text=text + "\n\nYou will not get popups until this is fixed.", fg="#fecaca",
                 bg="#7f1d1d", font=("Segoe UI", 10), wraplength=380, justify="left").pack(anchor="w", pady=8)
        tk.Button(w, text="OK", width=10, command=w.destroy).pack(anchor="e")
        w.update_idletasks()
        w.geometry(f"+{w.winfo_screenwidth() - w.winfo_width() - 20}+40")

    def popup(self, subject, start, location):
        BG, FG, MUTED, ACCENT, AMBER = "#1b1f27", "#f3f4f6", "#9aa3b2", "#4f8cff", "#f59e0b"
        w = tk.Toplevel(self.root)
        w.overrideredirect(True)
        w.attributes("-topmost", True)
        w.attributes("-alpha", 0.0)
        w.configure(bg=BG)
        bar = tk.Frame(w, bg=ACCENT, width=5)
        bar.pack(side="left", fill="y")
        body = tk.Frame(w, bg=BG, padx=18, pady=14)
        body.pack(side="left", fill="both", expand=True)

        status = tk.Label(body, font=("Segoe UI Semibold", 8), fg=ACCENT, bg=BG)
        status.pack(anchor="w")
        tk.Label(body, text=subject, font=("Segoe UI Semibold", 15), fg=FG, bg=BG,
                 wraplength=330, justify="left").pack(anchor="w", pady=(2, 0))
        when = start.strftime("%I:%M %p").lstrip("0")
        tk.Label(body, text=when + (f"   ·   {location}" if location else ""), font=("Segoe UI", 10),
                 fg=MUTED, bg=BG, wraplength=330, justify="left").pack(anchor="w", pady=(2, 12))

        def button(parent, text, cmd, primary):
            base, hover = (ACCENT, "#6a9eff") if primary else ("#2a303b", "#363d4a")
            b = tk.Label(parent, text=text, font=("Segoe UI Semibold", 10), fg="white" if primary else FG,
                         bg=base, padx=16, pady=6, cursor="hand2")
            b.bind("<Enter>", lambda e: b.config(bg=hover))
            b.bind("<Leave>", lambda e: b.config(bg=base))
            b.bind("<Button-1>", lambda e: cmd())
            return b

        row = tk.Frame(body, bg=BG)
        row.pack(anchor="w")
        button(row, "Dismiss", lambda: self.close(w), True).pack(side="left")
        button(row, f"Snooze {SNOOZE_MIN} min",
               lambda: self.snooze(w, subject, start, location), False).pack(side="left", padx=(8, 0))

        def tick():
            if not w.winfo_exists():
                return
            secs = (start - dt.datetime.now()).total_seconds()
            if secs > 60:
                txt, col = f"STARTS IN {int(secs // 60) + 1} MIN", ACCENT
            elif secs > 0:
                txt, col = "STARTS IN LESS THAN A MINUTE", ACCENT
            elif secs > -60:
                txt, col = "STARTING NOW", AMBER
            else:
                txt, col = f"STARTED {int(-secs // 60)} MIN AGO", AMBER
            status.config(text=txt, fg=col)
            bar.config(bg=col)
            w.after(1000, tick)
        tick()

        # drag anywhere on the card
        drag = {}
        def press(e): drag["x"], drag["y"] = e.x_root - w.winfo_x(), e.y_root - w.winfo_y()
        def move(e): w.geometry(f"+{e.x_root - drag['x']}+{e.y_root - drag['y']}")
        for widget in (w, body, bar):
            widget.bind("<ButtonPress-1>", press)
            widget.bind("<B1-Motion>", move)

        w.update_idletasks()
        w._size = (max(380, w.winfo_reqwidth()), w.winfo_reqheight())
        self.cards.append(w)
        style_window(w)
        self.reflow()
        self.fade_in(w)
        w.bell()

    def reflow(self):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        y = sh - 64
        for c in self.cards:
            cw, ch = c._size
            y -= ch + 10
            c.geometry(f"{cw}x{ch}+{sw - cw - 20}+{y}")

    def fade_in(self, w, step=0):
        if not w.winfo_exists():
            return
        w.attributes("-alpha", min(0.97, step / 8))
        if step < 8:
            w.after(16, lambda: self.fade_in(w, step + 1))

    def close(self, w):
        if w in self.cards:
            self.cards.remove(w)
        w.destroy()
        self.reflow()

    def snooze(self, w, subject, start, location):
        self.close(w)
        self.root.after(SNOOZE_MIN * 60_000, lambda: self.popup(subject, start, location))


def already_running():
    """Packaged exe: launching it again offers to turn the reminders off."""
    if message_box("Calendar Reminder is running.\n\nTurn it off?", yes_no=True):
        win32event.SetEvent(win32event.CreateEvent(None, True, False, QUIT_EVENT))
        if startup_enabled() and message_box("Also stop it from starting when you log in?", yes_no=True):
            set_startup(False)
        message_box("Calendar Reminder is off. Run it again any time to turn it back on.")


def first_run_setup():
    """Packaged exe: on first launch, offer to start at login, and confirm it is running."""
    if not startup_enabled():
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\OutlookCalendarReminder") as k:
                try:
                    winreg.QueryValueEx(k, "asked")
                    return_after_ask = False
                except OSError:
                    winreg.SetValueEx(k, "asked", 0, winreg.REG_DWORD, 1)
                    return_after_ask = True
        except OSError:
            return_after_ask = False
        if return_after_ask and message_box(
                "Start Calendar Reminder automatically when you log in?\n\n"
                "(Recommended, so you never have to think about it.)", yes_no=True):
            set_startup(True)
    message_box("Calendar Reminder is running in the background.\n\n"
                "You'll get a popup 1 minute before each Outlook event "
                "(classic Outlook must be open).\n\nRun this program again to turn it off.")


if __name__ == "__main__":
    _mutex = win32event.CreateMutex(None, False, r"Global\OutlookCalendarReminder")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        if FROZEN and "--startup" not in sys.argv:
            already_running()
        sys.exit(0)  # already running
    if FROZEN and "--startup" not in sys.argv:
        first_run_setup()
    pythoncom.CoInitialize()
    app = App()
    if "--test" in sys.argv:
        app.popup("Test reminder", dt.datetime.now(), "Everything works")
    try:
        app.root.mainloop()
    except BaseException:
        log(traceback.format_exc())
        raise
