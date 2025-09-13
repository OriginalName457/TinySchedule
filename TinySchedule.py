import json
import threading
import time
import uuid
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import calendar as calmod
from datetime import datetime, timedelta
from pynput import mouse, keyboard

# Optional deps (handled gracefully if missing)
PIL_AVAILABLE = True
CV2_AVAILABLE = True
try:
    from PIL import Image, ImageTk, ImageSequence
except Exception:
    PIL_AVAILABLE = False
try:
    import cv2
except Exception:
    CV2_AVAILABLE = False


# ----------------------------
# Macro engine (records + plays back)
# ----------------------------
class Macro:
    def __init__(self):
        self.events = []
        self.recording = False
        self.paused = False
        self._last_ts = None
        self._mouse_listener = None
        self._kb_listener = None
        self._stop_flag = threading.Event()
        self._lock = threading.RLock()
        self._controllers = {"mouse": mouse.Controller(), "keyboard": keyboard.Controller()}
        self._control_keys = set()

    def start_recording(self):
        with self._lock:
            if self.recording:
                return
            self.events = []
            self._stop_flag.clear()
            self.recording = True
            self.paused = False
            self._last_ts = time.time()
            self._mouse_listener = mouse.Listener(on_move=self._on_move, on_click=self._on_click, on_scroll=self._on_scroll)
            self._kb_listener = keyboard.Listener(on_press=self._on_press, on_release=self._on_release)
            self._mouse_listener.start()
            self._kb_listener.start()

    def stop_recording(self):
        with self._lock:
            if not self.recording:
                return
            self.recording = False
            if self._mouse_listener:
                self._mouse_listener.stop(); self._mouse_listener = None
            if self._kb_listener:
                self._kb_listener.stop(); self._kb_listener = None

    def pause_toggle(self):
        with self._lock:
            if self.recording:
                self.paused = not self.paused
                self._last_ts = time.time()

    def _dt(self):
        now = time.time()
        dt = now - self._last_ts if self._last_ts else 0
        self._last_ts = now
        return dt

    def _rec(self, item):
        with self._lock:
            if not self.recording or self.paused:
                return
            self.events.append(item)

    def _on_move(self, x, y):
        self._rec({"t": self._dt(), "type": "move", "x": int(x), "y": int(y)})

    def _on_click(self, x, y, button, pressed):
        self._rec({"t": self._dt(), "type": "click", "x": int(x), "y": int(y), "button": button.name, "pressed": bool(pressed)})

    def _on_scroll(self, x, y, dx, dy):
        self._rec({"t": self._dt(), "type": "scroll", "x": int(x), "y": int(y), "dx": int(dx), "dy": int(dy)})

    def _on_press(self, key):
        if key in self._control_keys:
            return
        try:
            k = key.char; kind = "char"
        except AttributeError:
            k = getattr(key, "name", str(key)); kind = "special"
        self._rec({"t": self._dt(), "type": "key", "action": "press", "kind": kind, "key": k})

    def _on_release(self, key):
        if key in self._control_keys:
            return
        try:
            k = key.char; kind = "char"
        except AttributeError:
            k = getattr(key, "name", str(key)); kind = "special"
        self._rec({"t": self._dt(), "type": "key", "action": "release", "kind": kind, "key": k})

    def save(self, path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.events, f, ensure_ascii=False, separators=(",", ":"))

    def load(self, path):
        with open(path, "r", encoding="utf-8") as f:
            self.events = json.load(f)

    def abort_playback(self):
        self._stop_flag.set()

    def _first_xy(self):
        for ev in self.events:
            if ev.get("type") in ("move", "click", "scroll"):
                return int(ev.get("x", 0)), int(ev.get("y", 0))
        return None

    def play(self, speed=1.0, loops=1, suppress_hotkeys=None):
        if not self.events:
            return
        self._stop_flag.clear()
        if suppress_hotkeys is None:
            suppress_hotkeys = []
        first = self._first_xy()
        if first:
            self._controllers["mouse"].position = first
        for _ in range(max(1, int(loops))):
            if self._stop_flag.is_set():
                break
            for ev in self.events:
                if self._stop_flag.is_set():
                    break
                t = float(ev.get("t", 0)) / max(0.001, float(speed))
                if t > 0:
                    time.sleep(t)
                typ = ev.get("type")
                if typ == "move":
                    self._controllers["mouse"].position = (int(ev["x"]), int(ev["y"]))
                elif typ == "click":
                    btn = getattr(mouse.Button, ev.get("button", "left"))
                    self._controllers["mouse"].position = (int(ev["x"]), int(ev["y"]))
                    if ev.get("pressed", True):
                        self._controllers["mouse"].press(btn)
                    else:
                        self._controllers["mouse"].release(btn)
                elif typ == "scroll":
                    self._controllers["mouse"].position = (int(ev["x"]), int(ev["y"]))
                    self._controllers["mouse"].scroll(int(ev.get("dx", 0)), int(ev.get("dy", 0)))
                elif typ == "key":
                    k = ev.get("key")
                    if ev.get("kind") == "char":
                        keyobj = k
                    else:
                        keyobj = getattr(keyboard.Key, k, None)
                        if keyobj is None or keyobj in suppress_hotkeys:
                            continue
                    if ev.get("action") == "press":
                        self._controllers["keyboard"].press(keyobj)
                    else:
                        self._controllers["keyboard"].release(keyobj)


# ----------------------------
# Media panel (image/GIF/video loop)
# ----------------------------
class MediaPanel(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=(8, 8, 8, 8))
        self.configure(borderwidth=1)
        # UI
        toolbar = ttk.Frame(self)
        toolbar.pack(fill=tk.X)
        ttk.Button(toolbar, text="Load Image/GIF/Video", command=self._pick_media).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Clear", command=self.clear).pack(side=tk.LEFT, padx=(6,0))
        self._status = ttk.Label(toolbar, text="", foreground="#666")
        self._status.pack(side=tk.RIGHT)

        self._stage = tk.Label(self, bg="#0e0e10")
        self._stage.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        # State
        self._img_obj = None          # current PhotoImage
        self._gif_frames = []         # list[PhotoImage]
        self._gif_delays = []         # list[int ms]
        self._gif_idx = 0
        self._gif_job = None

        self._cap = None              # OpenCV VideoCapture
        self._video_job = None
        self._video_fps_delay = 33

        # Resize handling
        self.bind("<Configure>", lambda e: self._on_resize())

    def _pick_media(self):
        path = filedialog.askopenfilename(
            filetypes=[
                ("Images/GIF/Video", "*.png *.jpg *.jpeg *.gif *.bmp *.mp4 *.mov *.avi *.mkv *.webm"),
                ("All files", "*.*")
            ]
        )
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext in [".png", ".jpg", ".jpeg", ".bmp"]:
            self.load_image(path)
        elif ext == ".gif":
            self.load_gif(path)
        else:
            self.load_video(path)

    def clear(self):
        self._status.config(text="")
        self._stop_gif()
        self._stop_video()
        self._img_obj = None
        self._stage.config(image="", text="")
        self._stage.update_idletasks()

    # ---------- Image ----------
    def load_image(self, path):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing dependency",
                                 "Pillow is required for images.\nInstall: sudo apt install python3-pil")
            return
        self._stop_gif(); self._stop_video()
        try:
            img = Image.open(path).convert("RGBA")
            self._status.config(text=os.path.basename(path))
            self._display_pil(img)
        except Exception as e:
            messagebox.showerror("Image error", str(e))

    # ---------- GIF (animated) ----------
    def load_gif(self, path):
        if not PIL_AVAILABLE:
            messagebox.showerror("Missing dependency",
                                 "Pillow is required for GIFs.\nInstall: sudo apt install python3-pil")
            return
        self._stop_gif(); self._stop_video()
        try:
            im = Image.open(path)
            frames = []
            delays = []
            for frame in ImageSequence.Iterator(im):
                delay = frame.info.get("duration", 100)
                frames.append(frame.convert("RGBA"))
                delays.append(max(10, delay))
            self._gif_frames = frames
            self._gif_delays = delays
            self._gif_idx = 0
            self._status.config(text=os.path.basename(path))
            self._run_gif()
        except Exception as e:
            messagebox.showerror("GIF error", str(e))

    def _run_gif(self):
        if not self._gif_frames:
            return
        target_size = self._stage.winfo_width(), self._stage.winfo_height()
        if target_size == (1, 1):
            target_size = (640, 360)
        frame = self._fit_pil(self._gif_frames[self._gif_idx], target_size)
        self._img_obj = ImageTk.PhotoImage(frame)
        self._stage.config(image=self._img_obj)
        delay = self._gif_delays[self._gif_idx]
        self._gif_idx = (self._gif_idx + 1) % len(self._gif_frames)
        self._gif_job = self.after(delay, self._run_gif)

    def _stop_gif(self):
        if self._gif_job:
            try: self.after_cancel(self._gif_job)
            except Exception: pass
            self._gif_job = None
        self._gif_frames = []
        self._gif_delays = []
        self._gif_idx = 0

    # ---------- Video ----------
    def load_video(self, path):
        if not CV2_AVAILABLE or not PIL_AVAILABLE:
            messagebox.showerror(
                "Missing dependency",
                "Video playback needs OpenCV and Pillow.\nInstall: sudo apt install python3-opencv python3-pil"
            )
            return
        self._stop_gif(); self._stop_video()
        try:
            cap = cv2.VideoCapture(path)
            if not cap.isOpened():
                raise RuntimeError("Could not open video.")
            fps = cap.get(cv2.CAP_PROP_FPS)
            self._video_fps_delay = int(1000 / fps) if fps and fps > 0 else 33
            self._cap = cap
            self._status.config(text=os.path.basename(path))
            self._run_video()
        except Exception as e:
            messagebox.showerror("Video error", str(e))

    def _run_video(self):
        if self._cap is None:
            return
        ok, frame = self._cap.read()
        if not ok:
            # loop to start
            self._cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ok, frame = self._cap.read()
            if not ok:
                self._stop_video()
                return
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        pil = Image.fromarray(frame)
        target_size = self._stage.winfo_width(), self._stage.winfo_height()
        if target_size == (1, 1):
            target_size = (640, 360)
        pil = self._fit_pil(pil, target_size)
        self._img_obj = ImageTk.PhotoImage(pil)
        self._stage.config(image=self._img_obj)
        self._video_job = self.after(self._video_fps_delay, self._run_video)

    def _stop_video(self):
        if self._video_job:
            try: self.after_cancel(self._video_job)
            except Exception: pass
            self._video_job = None
        if self._cap is not None:
            try: self._cap.release()
            except Exception: pass
            self._cap = None

    # ---------- Helpers ----------
    def _on_resize(self):
        # re-render current media to fit
        if self._gif_frames:
            # force immediate next frame render at new size
            if self._gif_job:
                try: self.after_cancel(self._gif_job)
                except Exception: pass
                self._gif_job = None
            self._run_gif()
        elif self._cap is not None:
            # next tick will resize
            pass
        elif self._img_obj:
            # redisplay static image scaled
            # (we don't have original PIL stored, so reload path is required to do perfect resample)
            # Instead: do nothing; user can reload. Keeping it simple.
            pass

    def _fit_pil(self, img, target_size):
        tw, th = target_size
        if tw <= 0 or th <= 0:
            return img
        iw, ih = img.size
        scale = min(tw / iw, th / ih)
        nw, nh = max(1, int(iw * scale)), max(1, int(ih * scale))
        return img.resize((nw, nh), Image.LANCZOS)

    def destroy(self):
        self.clear()
        super().destroy()


# ----------------------------
# Scheduler (stores & triggers events)
# ----------------------------
class Scheduler:
    def __init__(self, macro_ref, persist_path):
        self._macro_ref = macro_ref
        self._persist_path = persist_path
        self._events = {}
        self._lock = threading.RLock()
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self.load()
        self._thread.start()

    def _now(self):
        return datetime.now()

    def add(self, title, when_dt, macro_path, speed=1.0, loops=1, uid=None):
        with self._lock:
            if uid is None:
                uid = str(uuid.uuid4())
            self._events[uid] = {
                "id": uid,
                "title": title,
                "when": when_dt.isoformat(),
                "macro": macro_path,
                "speed": float(speed),
                "loops": int(loops),
            }
            self.save()
            return uid

    def remove(self, uid):
        with self._lock:
            if uid in self._events:
                del self._events[uid]
                self.save()

    def list_all(self):
        with self._lock:
            arr = list(self._events.values())
        arr.sort(key=lambda e: e["when"])
        return arr

    def save(self):
        with self._lock:
            try:
                with open(self._persist_path, "w", encoding="utf-8") as f:
                    json.dump(self._events, f, ensure_ascii=False, indent=2)
            except Exception:
                pass

    def load(self):
        try:
            with open(self._persist_path, "r", encoding="utf-8") as f:
                self._events = json.load(f)
        except Exception:
            self._events = {}

    def _loop(self):
        while not self._stop.is_set():
            time.sleep(0.5)
            due = []
            with self._lock:
                now = self._now()
                for k, e in list(self._events.items()):
                    when = datetime.fromisoformat(e["when"])
                    if when <= now:
                        due.append(e)
                        del self._events[k]
                if due:
                    self.save()
            for e in due:
                try:
                    self._run_event(e)
                except Exception:
                    pass

    def _run_event(self, e):
        m = Macro()
        if os.path.isfile(e["macro"]):
            m.load(e["macro"])
            m.play(
                speed=e.get("speed", 1.0),
                loops=e.get("loops", 1),
                suppress_hotkeys=[keyboard.Key.f8, keyboard.Key.f7, keyboard.Key.f9, keyboard.Key.f10]
            )


# ----------------------------
# App UI (calendar kept from previous version)
# ----------------------------
class App:
    # Larger, easy-click calendar cells
    CELL_W = 160
    CELL_H = 120
    CELL_PAD = 6
    CELL_BG = "#f7f7f9"
    CELL_HILITE = "#dbeafe"
    BADGE_BG = "#eef2ff"
    BADGE_FG = "#3730a3"

    def __init__(self, root):
        self.root = root
        self.root.title("TinyTask-Py")

        self.macro = Macro()
        self.speed = tk.DoubleVar(value=1.0)
        self.loops = tk.IntVar(value=1)
        self.status = tk.StringVar(value="Ready")

        # Hotkeys
        self.hk = keyboard.GlobalHotKeys({
            "<f8>": self._hk_record_stop,
            "<f7>": self.on_pause,
            "<f9>": self.on_play,
            "<f10>": self.on_abort
        })
        self.hk.start()
        self.macro._control_keys = {keyboard.Key.f8, keyboard.Key.f7, keyboard.Key.f9, keyboard.Key.f10}

        # Notebook
        self._notebook = ttk.Notebook(self.root)
        self._macro_tab = ttk.Frame(self._notebook)
        self._sched_tab = ttk.Frame(self._notebook)
        self._notebook.add(self._macro_tab, text="Macro")
        self._notebook.add(self._sched_tab, text="Scheduler")
        self._notebook.pack(fill=tk.BOTH, expand=True)

        self._build_macro_ui(self._macro_tab)

        now = datetime.now()
        self._sched_state = {
            "cur_year": now.year,
            "cur_month": now.month,
            "selected_date": None,
            "cells": {}
        }
        self.scheduler = Scheduler(self.macro, os.path.join(os.path.expanduser("~"), ".tinytask_py_schedule.json"))
        self._build_sched_ui(self._sched_tab)

    # ---------- Macro tab (left controls + right media panel) ----------
    def _build_macro_ui(self, parent):
        root_frame = ttk.Frame(parent, padding=8)
        root_frame.pack(fill=tk.BOTH, expand=True)
        root_frame.columnconfigure(0, weight=0)  # controls
        root_frame.columnconfigure(1, weight=1)  # media

        controls = ttk.Frame(root_frame)
        controls.grid(row=0, column=0, sticky="nsw", padx=(0,8))
        media_wrap = ttk.Frame(root_frame)
        media_wrap.grid(row=0, column=1, sticky="nsew")

        row = 0
        ttk.Label(controls, text="Playback speed").grid(row=row, column=0, sticky="w")
        ttk.Scale(controls, from_=0.25, to=3.0, orient=tk.HORIZONTAL, variable=self.speed, length=180).grid(row=row, column=1, sticky="w")
        row += 1
        ttk.Label(controls, text="Loops").grid(row=row, column=0, sticky="w")
        tk.Spinbox(controls, from_=1, to=9999, textvariable=self.loops, width=8).grid(row=row, column=1, sticky="w")
        row += 1
        btns = ttk.Frame(controls)
        btns.grid(row=row, column=0, columnspan=2, pady=(8,4), sticky="w")
        ttk.Button(btns, text="Record (F8)", command=self.on_record).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Pause (F7)", command=self.on_pause).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Stop (F8)", command=self.on_stop).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Play (F9)", command=self.on_play).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Abort (F10)", command=self.on_abort).pack(side=tk.LEFT, padx=2)
        row += 1
        filebar = ttk.Frame(controls)
        filebar.grid(row=row, column=0, columnspan=2, pady=(4,8), sticky="w")
        ttk.Button(filebar, text="Save", command=self.on_save).pack(side=tk.LEFT, padx=2)
        ttk.Button(filebar, text="Load", command=self.on_load).pack(side=tk.LEFT, padx=2)
        row += 1
        ttk.Label(controls, text="Hotkeys: F8 start/stop record, F7 pause, F9 play, F10 abort").grid(row=row, column=0, columnspan=2, sticky="w", pady=(4,0))
        row += 1
        ttk.Label(controls, textvariable=self.status).grid(row=row, column=0, columnspan=2, sticky="w", pady=(8,0))

        media_wrap.rowconfigure(0, weight=1)
        media_wrap.columnconfigure(0, weight=1)
        self.media_panel = MediaPanel(media_wrap)
        self.media_panel.grid(row=0, column=0, sticky="nsew")

    # ---------- Scheduler tab (same function as your last build) ----------
    def _build_sched_ui(self, parent):
        top = ttk.Frame(parent, padding=10)
        top.pack(fill=tk.BOTH, expand=True)

        hdr = ttk.Frame(top)
        hdr.pack(fill=tk.X)
        ttk.Button(hdr, text="◀", width=3, command=self._prev_month).pack(side=tk.LEFT)
        self._month_label = ttk.Label(hdr, text="", font=("TkDefaultFont", 12, "bold"))
        self._month_label.pack(side=tk.LEFT, expand=True)
        ttk.Button(hdr, text="▶", width=3, command=self._next_month).pack(side=tk.LEFT)

        mid = ttk.Frame(top)
        mid.pack(fill=tk.BOTH, expand=True, pady=(8,0))

        cal_container = ttk.Frame(mid)
        cal_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)

        weekdays = ttk.Frame(cal_container)
        weekdays.pack(fill=tk.X)
        for i, d in enumerate(["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]):
            lbl = ttk.Label(weekdays, text=d, anchor="center")
            lbl.grid(row=0, column=i, padx=6, pady=(0,6), sticky="ew")
            weekdays.grid_columnconfigure(i, weight=0)

        self._calframe = ttk.Frame(cal_container)
        self._calframe.pack()

        right = ttk.Frame(mid)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(16,0))
        self._selected_label = ttk.Label(right, text="Select a date", font=("TkDefaultFont", 10, "bold"))
        self._selected_label.pack(anchor="w")
        self._events_list = tk.Listbox(right, height=16)
        self._events_list.pack(fill=tk.BOTH, expand=True)
        btnbar = ttk.Frame(right)
        btnbar.pack(fill=tk.X, pady=6)
        ttk.Button(btnbar, text="Add", command=self._add_event_dialog).pack(side=tk.LEFT, padx=4)
        ttk.Button(btnbar, text="Remove", command=self._remove_selected_event).pack(side=tk.LEFT, padx=4)

        self._render_calendar()
        self._refresh_events_list()

    def _bind_click_recursive(self, widget, on_click, on_dbl):
        widget.bind("<Button-1>", on_click)
        widget.bind("<Double-Button-1>", on_dbl)
        for ch in widget.winfo_children():
            self._bind_click_recursive(ch, on_click, on_dbl)

    def _render_calendar(self):
        for w in self._calframe.winfo_children():
            w.destroy()
        self._sched_state["cells"].clear()

        y = self._sched_state["cur_year"]; m = self._sched_state["cur_month"]
        self._month_label.config(text=f"{calmod.month_name[m]} {y}")

        raw_weeks = calmod.Calendar(firstweekday=0).monthdayscalendar(y, m)
        while raw_weeks and all(d == 0 for d in raw_weeks[-1]):
            raw_weeks.pop()

        for r, week in enumerate(raw_weeks):
            for c, day in enumerate(week):
                cell = tk.Frame(
                    self._calframe, bd=1, relief=tk.SOLID,
                    width=self.CELL_W, height=self.CELL_H, bg=self.CELL_BG, highlightthickness=0
                )
                cell.grid_propagate(False)
                cell.grid(row=r, column=c, padx=self.CELL_PAD, pady=self.CELL_PAD)
                self._sched_state["cells"][(r, c)] = {"frame": cell, "day": day}

                if day == 0:
                    continue

                def on_click(evt=None, dd=day):
                    self._select_date(y, m, dd)
                def on_dbl(evt=None, dd=day):
                    self._select_date(y, m, dd)
                    self._add_event_dialog(prefill_date=datetime(y, m, dd))

                day_lbl = tk.Label(cell, text=str(day), bg=self.CELL_BG, anchor="w")
                day_lbl.place(x=8, y=6)

                plus = ttk.Button(cell, text="+", width=2, command=lambda dd=day: self._quick_add(dd))
                plus.place(relx=1.0, x=-8, y=6, anchor="ne")

                count = self._count_events_for_date(y, m, day)
                if count:
                    badge = tk.Label(cell, text=f"{count} scheduled", bg=self.BADGE_BG, fg=self.BADGE_FG)
                    badge.place(x=8, y=self.CELL_H-24)

                self._bind_click_recursive(cell, on_click, on_dbl)

        for c in range(7):
            self._calframe.grid_columnconfigure(c, minsize=self.CELL_W + 2*self.CELL_PAD)
        for r in range(len(raw_weeks)):
            self._calframe.grid_rowconfigure(r, minsize=self.CELL_H + 2*self.CELL_PAD)

        sel = self._sched_state.get("selected_date")
        if sel and sel.year == y and sel.month == m:
            self._highlight_selected(sel.day)
        else:
            self._sched_state["selected_date"] = None
            self._selected_label.config(text="Select a date")

    def _select_date(self, year, month, day):
        self._sched_state["selected_date"] = datetime(year, month, day).date()
        self._selected_label.config(text=f"Events on {self._sched_state['selected_date'].strftime('%Y-%m-%d')}")
        self._highlight_selected(day)
        self._refresh_events_list()

    def _highlight_selected(self, day):
        for info in self._sched_state["cells"].values():
            info["frame"].configure(bg=self.CELL_BG)
        for (r, c), info in self._sched_state["cells"].items():
            if info["day"] == day:
                info["frame"].configure(bg=self.CELL_HILITE)
                break

    def _count_events_for_date(self, y, m, d):
        target = f"{y:04d}-{m:02d}-{d:02d}"
        return sum(1 for e in self.scheduler.list_all()
                   if datetime.fromisoformat(e["when"]).strftime("%Y-%m-%d") == target)

    def _prev_month(self):
        y = self._sched_state["cur_year"]; m = self._sched_state["cur_month"]
        dt = datetime(y, m, 15) - timedelta(days=31)
        self._sched_state["cur_year"], self._sched_state["cur_month"] = dt.year, dt.month
        self._render_calendar(); self._refresh_events_list()

    def _next_month(self):
        y = self._sched_state["cur_year"]; m = self._sched_state["cur_month"]
        dt = datetime(y, m, 15) + timedelta(days=31)
        self._sched_state["cur_year"], self._sched_state["cur_month"] = dt.year, dt.month
        self._render_calendar(); self._refresh_events_list()

    def _quick_add(self, day):
        y = self._sched_state["cur_year"]; m = self._sched_state["cur_month"]
        self._add_event_dialog(prefill_date=datetime(y, m, day))
        self._render_calendar(); self._refresh_events_list()

    def _add_event_dialog(self, prefill_date=None):
        win = tk.Toplevel(self.root)
        win.title("Schedule Macro")
        frm = ttk.Frame(win, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Title").grid(row=0, column=0, sticky="e"); title_var = tk.StringVar(value="")
        ttk.Entry(frm, textvariable=title_var, width=30).grid(row=0, column=1, sticky="w")
        ttk.Label(frm, text="Date (YYYY-MM-DD)").grid(row=1, column=0, sticky="e")
        date_var = tk.StringVar(value=(prefill_date or datetime.now()).strftime("%Y-%m-%d"))
        ttk.Entry(frm, textvariable=date_var, width=15).grid(row=1, column=1, sticky="w")
        ttk.Label(frm, text="Time (HH:MM)").grid(row=2, column=0, sticky="e")
        time_var = tk.StringVar(value=datetime.now().strftime("%H:%M"))
        ttk.Entry(frm, textvariable=time_var, width=10).grid(row=2, column=1, sticky="w")
        ttk.Label(frm, text="Macro file").grid(row=3, column=0, sticky="e")
        macro_var = tk.StringVar(value="")
        def pick_file():
            p = filedialog.askopenfilename(filetypes=[("Macro JSON", "*.json"), ("All files", "*.*")])
            if p:
                macro_var.set(p)
        ttk.Entry(frm, textvariable=macro_var, width=30).grid(row=3, column=1, sticky="w")
        ttk.Button(frm, text="…", command=pick_file).grid(row=3, column=2, sticky="w")
        ttk.Label(frm, text="Speed").grid(row=4, column=0, sticky="e"); spd_var = tk.DoubleVar(value=1.0)
        tk.Spinbox(frm, from_=0.25, to=3.0, increment=0.05, textvariable=spd_var, width=8).grid(row=4, column=1, sticky="w")
        ttk.Label(frm, text="Loops").grid(row=5, column=0, sticky="e"); loop_var = tk.IntVar(value=1)
        tk.Spinbox(frm, from_=1, to=9999, textvariable=loop_var, width=8).grid(row=5, column=1, sticky="w")
        msg = ttk.Label(frm, text="")
        msg.grid(row=6, column=0, columnspan=3, sticky="w", pady=(6,0))

        def ok():
            try:
                dt = datetime.fromisoformat(date_var.get() + " " + time_var.get())
                if not os.path.isfile(macro_var.get()):
                    raise ValueError("Macro file not found")
                uid = self.scheduler.add(title_var.get() or "Macro", dt, macro_var.get(), speed=spd_var.get(), loops=loop_var.get())
                win.destroy()
                if dt.year == self._sched_state["cur_year"] and dt.month == self._sched_state["cur_month"]:
                    self._select_date(dt.year, dt.month, dt.day)
                else:
                    self._refresh_events_list()
                messagebox.showinfo(
                    "Scheduled",
                    f"Event scheduled: {title_var.get()} @ {dt.strftime('%Y-%m-%d %H:%M')}\nID: {uid}"
                )
            except Exception as e:
                msg.config(text=str(e))

        ttk.Button(frm, text="Schedule", command=ok).grid(row=7, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=win.destroy).grid(row=7, column=1, pady=8)

    def _refresh_events_list(self):
        items = self.scheduler.list_all()
        self._events_list.delete(0, tk.END)
        sel = self._sched_state.get("selected_date")
        y = self._sched_state["cur_year"]; m = self._sched_state["cur_month"]
        if sel:
            for e in items:
                dt = datetime.fromisoformat(e["when"])
                if dt.date() == sel:
                    self._events_list.insert(tk.END, f"{e['id'][:8]} | {dt.strftime('%H:%M')} | {e['title']}")
        else:
            for e in items:
                dt = datetime.fromisoformat(e["when"])
                if dt.year == y and dt.month == m:
                    self._events_list.insert(tk.END, f"{e['id'][:8]} | {dt.strftime('%Y-%m-%d %H:%M')} | {e['title']}")

    # ---------- Macro actions ----------
    def _hk_record_stop(self):
        if self.macro.recording:
            self.on_stop()
        else:
            self.on_record()

    def on_record(self):
        self.status.set("Recording…")
        self.macro.start_recording()

    def on_pause(self):
        self.macro.pause_toggle()
        self.status.set("Paused" if self.macro.paused else "Recording…")

    def on_stop(self):
        self.macro.stop_recording()
        self.status.set(f"Recorded {len(self.macro.events)} events")

    def on_play(self):
        if not self.macro.events:
            messagebox.showinfo("TinyTask-Py", "No events recorded")
            return
        self.status.set("Playing…")
        th = threading.Thread(target=self._play_thread, daemon=True)
        th.start()

    def _play_thread(self):
        self.macro.play(
            speed=self.speed.get(),
            loops=self.loops.get(),
            suppress_hotkeys=[keyboard.Key.f8, keyboard.Key.f7, keyboard.Key.f9, keyboard.Key.f10]
        )
        self.status.set("Ready")

    def on_abort(self):
        self.macro.abort_playback()
        self.status.set("Aborted")

    def on_save(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Macro JSON", "*.json")])
        if path:
            try:
                self.macro.save(path)
                self.status.set("Saved")
            except Exception as e:
                messagebox.showerror("Save failed", str(e))

    def on_load(self):
        path = filedialog.askopenfilename(filetypes=[("Macro JSON", "*.json"), ("All files", "*.*")])
        if path:
            try:
                self.macro.load(path)
                self.status.set(f"Loaded {len(self.macro.events)} events")
            except Exception as e:
                messagebox.showerror("Load failed", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
