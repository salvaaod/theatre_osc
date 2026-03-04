#!/usr/bin/env python3

import json
import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
from pythonosc.udp_client import SimpleUDPClient


# ==============================
# X32 CONFIGURATION
# ==============================

DEFAULT_OSC_IP = "192.168.10.59"
DEFAULT_OSC_PORT = 10023
DEFAULT_ADDRESS_TEMPLATE = "/ch/{ch:02d}/mix/on"
DEFAULT_OSC_VALUE_FOR_ON = 1
DEFAULT_OSC_VALUE_FOR_OFF = 0


TRUTHY = {"YES", "Y", "TRUE", "T", "1", "ON"}
FALSY = {"NO", "N", "FALSE", "F", "0", "OFF", ""}
APP_SETTINGS_FILE = "theatre_settings.json"


# ==============================
# Data Helpers
# ==============================

def normalize_to_bool(value):
    if value is None:
        return False
    s = str(value).strip().upper()
    if s in TRUTHY:
        return True
    if s in FALSY:
        return False
    return False


def load_excel_file(path):
    df = pd.read_excel(path, index_col=0, dtype=str)
    df = df.fillna("")
    df.index = df.index.map(lambda x: str(x).strip())
    df.columns = [str(c).strip() for c in df.columns]
    return df


def dataframe_to_scene_dict(df):
    scenes = {}
    for scene in df.index:
        scenes[scene] = {
            actor: normalize_to_bool(df.loc[scene, actor])
            for actor in df.columns
        }
    return scenes


def build_channel_map(actors):
    return {actor: i + 1 for i, actor in enumerate(actors)}


def split_actor_text(actor, card_width, font_size):
    actor = str(actor).strip()
    if not actor:
        return ""

    chars_per_line = max(3, (card_width - 12) // max(font_size - 2, 1))
    if len(actor) <= chars_per_line:
        return actor

    words = actor.split()
    if len(words) > 1:
        line1 = ""
        line2_words = []
        for w in words:
            candidate = f"{line1} {w}".strip()
            if len(candidate) <= chars_per_line:
                line1 = candidate
            else:
                line2_words.append(w)
        if line1:
            line2 = " ".join(line2_words).strip()
            if len(line2) > chars_per_line:
                line2 = f"{line2[:chars_per_line - 1]}…"
            return f"{line1}\n{line2}" if line2 else line1

    line1 = actor[:chars_per_line].rstrip()
    line2 = actor[chars_per_line:].lstrip()
    if len(line2) > chars_per_line:
        line2 = f"{line2[:chars_per_line - 1]}…"
    return f"{line1}\n{line2}"


def find_startup_excel(base_dir, remembered_path):
    candidates = []

    if remembered_path:
        remembered_path = remembered_path.strip()
        if remembered_path:
            if os.path.isabs(remembered_path):
                candidates.append(remembered_path)
            else:
                candidates.append(os.path.join(base_dir, remembered_path))
                candidates.append(os.path.join(os.getcwd(), remembered_path))

    for scan_dir in [base_dir, os.getcwd()]:
        if not os.path.isdir(scan_dir):
            continue
        for name in os.listdir(scan_dir):
            low = name.lower()
            if low.endswith(".xlsx") or low.endswith(".xls"):
                candidates.append(os.path.join(scan_dir, name))

    unique_existing = []
    seen = set()
    for p in candidates:
        abs_path = os.path.abspath(p)
        if abs_path in seen:
            continue
        seen.add(abs_path)
        if os.path.isfile(abs_path):
            unique_existing.append(abs_path)

    if not unique_existing:
        return None

    unique_existing.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return unique_existing[0]


# ==============================
# OSC Sender
# ==============================

class X32Sender:
    def __init__(self, ip, port, address_template, actor_channel_map):
        self.client = SimpleUDPClient(ip, port)
        self.address_template = address_template
        self.actor_channel_map = actor_channel_map

    def send(self, actor, enabled):
        ch = self.actor_channel_map[actor]
        address = self.address_template.format(ch=ch)
        value = DEFAULT_OSC_VALUE_FOR_ON if enabled else DEFAULT_OSC_VALUE_FOR_OFF
        self.client.send_message(address, value)
        logging.info("OSC: %s %s", address, value)


# ==============================
# Application
# ==============================

class TheatreApp:
    def __init__(self, root):
        self.root = root
        self.root.title("X32 Theatre Mic Controller")
        self.root.configure(bg="black")

        logging.basicConfig(
            filename="show_log.txt",
            level=logging.INFO,
            format="%(asctime)s %(message)s"
        )

        self.settings_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            APP_SETTINGS_FILE
        )

        self.df = None
        self.scenes = {}
        self.scene_names = []
        self.actors = []
        self.channel_map = {}
        self.current_scene_index = 0
        self.last_live_state = None
        self.last_excel_path = ""

        # Static dimensions configured by controls only
        self.card_width = 140
        self.card_height = 90
        self.card_font_size = 12

        self.live_mode = tk.BooleanVar(value=False)
        self.card_frames = {}
        self.card_name_labels = {}

        self.build_ui()
        self.load_settings()
        self.try_load_startup_excel()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- UI ----------

    def build_ui(self):
        top = tk.Frame(self.root, bg="black")
        top.pack(fill="x", padx=10, pady=10)

        tk.Button(top, text="Load Excel", command=self.load_excel).pack(side="left")

        self.live_checkbox = tk.Checkbutton(
            top,
            text="LIVE MODE",
            variable=self.live_mode,
            fg="white",
            bg="black",
            selectcolor="black",
            font=("Helvetica", 14)
        )
        self.live_checkbox.pack(side="right")

        self.scene_label = tk.Label(
            self.root,
            text="No Scene",
            font=("Helvetica", 24, "bold"),
            fg="white",
            bg="black",
            width=24
        )
        self.scene_label.pack(pady=10)

        controls = tk.Frame(self.root, bg="black")
        controls.pack()

        tk.Button(controls, text="Previous", command=self.previous_scene).pack(side="left", padx=5)
        tk.Button(controls, text="Next", command=self.next_scene).pack(side="left", padx=5)
        tk.Button(controls, text="GO", command=self.apply_scene, width=10).pack(side="left", padx=10)

        tk.Button(controls, text="H-", command=self.decrease_card_width, width=4).pack(side="left", padx=(20, 5))
        tk.Button(controls, text="H+", command=self.increase_card_width, width=4).pack(side="left", padx=5)
        tk.Button(controls, text="V-", command=self.decrease_card_height, width=4).pack(side="left", padx=(10, 5))
        tk.Button(controls, text="V+", command=self.increase_card_height, width=4).pack(side="left", padx=5)
        tk.Button(controls, text="A-", command=self.decrease_font_size, width=4).pack(side="left", padx=(10, 5))
        tk.Button(controls, text="A+", command=self.increase_font_size, width=4).pack(side="left", padx=5)

        self.status_label = tk.Label(
            self.root,
            text="No Excel loaded",
            font=("Helvetica", 10),
            fg="#bbbbbb",
            bg="black",
            anchor="w"
        )
        self.status_label.pack(fill="x", padx=20, pady=(6, 0))

        cards_container = tk.Frame(self.root, bg="black")
        cards_container.pack(fill="both", expand=True, padx=20, pady=(20, 20))

        self.cards_canvas = tk.Canvas(cards_container, bg="black", highlightthickness=0)
        self.cards_canvas.pack(side="top", fill="both", expand=True)

        self.hbar = tk.Scrollbar(cards_container, orient="horizontal", command=self.cards_canvas.xview)
        self.hbar.pack(side="bottom", fill="x")
        self.cards_canvas.configure(xscrollcommand=self.hbar.set)

        self.cards_strip = tk.Frame(self.cards_canvas, bg="black")
        self.cards_window_id = self.cards_canvas.create_window((0, 0), window=self.cards_strip, anchor="nw")
        self.cards_strip.bind("<Configure>", self.on_cards_strip_configure)

        self.root.bind("<Left>", lambda _e: self.previous_scene())
        self.root.bind("<Right>", lambda _e: self.next_scene())
        self.root.bind("<space>", lambda _e: self.apply_scene())
        self.root.bind("<Return>", lambda _e: self.apply_scene())

    # ---------- Settings ----------

    def load_settings(self):
        if not os.path.exists(self.settings_path):
            return

        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                settings = json.load(f)

            width = int(settings.get("width", 1100))
            height = int(settings.get("height", 800))
            x = int(settings.get("x", 100))
            y = int(settings.get("y", 100))

            self.card_width = int(settings.get("card_width", self.card_width))
            self.card_height = int(settings.get("card_height", self.card_height))
            self.card_font_size = int(settings.get("card_font_size", self.card_font_size))
            self.last_excel_path = str(settings.get("last_excel_path", "")).strip()

            self.root.geometry(f"{width}x{height}+{x}+{y}")
        except Exception as exc:
            logging.warning("Could not load app settings: %s", exc)

    def save_settings(self):
        self.root.update_idletasks()

        settings = {
            "width": self.root.winfo_width(),
            "height": self.root.winfo_height(),
            "x": self.root.winfo_x(),
            "y": self.root.winfo_y(),
            "card_width": self.card_width,
            "card_height": self.card_height,
            "card_font_size": self.card_font_size,
            "last_excel_path": self.last_excel_path,
        }

        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2)
        except Exception as exc:
            logging.warning("Could not save app settings: %s", exc)

    # ---------- Excel ----------

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        self.load_excel_from_path(path)

    def load_excel_from_path(self, path, show_error_dialog=True):
        if not path:
            return False

        try:
            df = load_excel_file(path)
            if df.empty or len(df.columns) == 0:
                raise ValueError("Excel has no scenes/actors")

            scenes = dataframe_to_scene_dict(df)
            if not scenes:
                raise ValueError("Excel has no scene rows")

            self.df = df
            self.scenes = scenes
            self.scene_names = list(scenes.keys())
            self.actors = list(df.columns)
            self.channel_map = build_channel_map(self.actors)
            self.current_scene_index = 0
            self.last_live_state = None
            self.last_excel_path = os.path.abspath(path)

            self.build_cards()
            self.render_scene()
            self.status_label.config(text=f"Loaded Excel: {os.path.basename(path)}")
            self.save_settings()

            logging.info("Loaded Excel: %s", self.last_excel_path)
            return True

        except Exception as exc:
            msg = f"Could not load Excel: {path} | {exc}"
            self.status_label.config(text=msg)
            logging.warning(msg)
            if show_error_dialog:
                messagebox.showerror("Excel load error", str(exc))
            return False

    def try_load_startup_excel(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        startup_path = find_startup_excel(base_dir, self.last_excel_path)
        if not startup_path:
            self.status_label.config(text="No startup Excel found")
            return

        if not self.load_excel_from_path(startup_path, show_error_dialog=False):
            self.status_label.config(text=f"Startup Excel failed: {os.path.basename(startup_path)}")

    # ---------- Card Drawing ----------

    def clear_cards(self):
        for w in self.cards_strip.winfo_children():
            w.destroy()
        self.card_frames.clear()
        self.card_name_labels.clear()

    def build_cards(self):
        self.clear_cards()

        for idx, actor in enumerate(self.actors):
            frame = tk.Frame(self.cards_strip, bg="#111111", bd=2, relief="ridge")
            frame.grid(row=0, column=idx, padx=10, pady=10)

            label = tk.Label(
                frame,
                text=split_actor_text(actor, self.card_width, self.card_font_size),
                fg="white",
                bg="#990000",
                font=("Helvetica", self.card_font_size, "bold"),
                justify="center"
            )
            label.pack(fill="both", expand=True, padx=4, pady=4)

            self.card_frames[actor] = frame
            self.card_name_labels[actor] = label

        self.apply_static_card_geometry()

    def apply_static_card_geometry(self):
        if not self.actors:
            return

        self.card_width = max(40, int(self.card_width))
        self.card_height = max(30, int(self.card_height))
        self.card_font_size = max(6, int(self.card_font_size))

        for actor in self.actors:
            frame = self.card_frames[actor]
            frame.configure(width=self.card_width, height=self.card_height)
            frame.grid_propagate(False)

            label = self.card_name_labels[actor]
            label.configure(
                font=("Helvetica", self.card_font_size, "bold"),
                text=split_actor_text(actor, self.card_width, self.card_font_size),
                justify="center"
            )

        self.cards_canvas.update_idletasks()
        self.cards_canvas.configure(scrollregion=self.cards_canvas.bbox("all"))

    def on_cards_strip_configure(self, _event):
        self.cards_canvas.configure(scrollregion=self.cards_canvas.bbox("all"))

    # ---------- Controls ----------

    def increase_card_width(self):
        self.card_width += 10
        self.apply_static_card_geometry()
        self.save_settings()

    def decrease_card_width(self):
        self.card_width -= 10
        self.apply_static_card_geometry()
        self.save_settings()

    def increase_card_height(self):
        self.card_height += 10
        self.apply_static_card_geometry()
        self.save_settings()

    def decrease_card_height(self):
        self.card_height -= 10
        self.apply_static_card_geometry()
        self.save_settings()

    def increase_font_size(self):
        self.card_font_size += 1
        self.apply_static_card_geometry()
        self.save_settings()

    def decrease_font_size(self):
        self.card_font_size -= 1
        self.apply_static_card_geometry()
        self.save_settings()

    # ---------- Scene ----------

    def previous_scene(self):
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index - 1) % len(self.scene_names)
        self.render_scene()

    def next_scene(self):
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index + 1) % len(self.scene_names)
        self.render_scene()

    def render_scene(self):
        if not self.scene_names:
            return

        scene_name = self.scene_names[self.current_scene_index]
        self.scene_label.config(text=f"Scene: {scene_name}")
        scene_state = self.scenes[scene_name]

        for actor, enabled in scene_state.items():
            frame = self.card_frames.get(actor)
            label = self.card_name_labels.get(actor)
            if not frame or not label:
                continue

            if enabled:
                frame.config(bg="#008000")
                label.config(bg="#008000")
            else:
                frame.config(bg="#990000")
                label.config(bg="#990000")

    def apply_scene(self):
        if not self.scene_names:
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.scenes[scene_name]

        self.render_scene()

        if not self.live_mode.get():
            logging.info("Preview scene: %s", scene_name)
            return

        sender = X32Sender(
            DEFAULT_OSC_IP,
            DEFAULT_OSC_PORT,
            DEFAULT_ADDRESS_TEMPLATE,
            self.channel_map,
        )

        changes = 0
        if self.last_live_state is None:
            for actor, enabled in scene_state.items():
                sender.send(actor, enabled)
                changes += 1
        else:
            for actor, enabled in scene_state.items():
                if self.last_live_state.get(actor) != enabled:
                    sender.send(actor, enabled)
                    changes += 1

        self.last_live_state = dict(scene_state)
        logging.info("LIVE scene: %s | Changes: %d", scene_name, changes)

    def on_close(self):
        self.save_settings()
        self.root.destroy()


def main():
    root = tk.Tk()
    TheatreApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
