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

DEFAULT_OSC_IP = "192.168.10.59"   # Change to your X32 IP
DEFAULT_OSC_PORT = 10023            # X32 default OSC port

DEFAULT_ADDRESS_TEMPLATE = "/ch/{ch:02d}/mix/on"

# X32 logic:
# 1 = ON (audio passes)
# 0 = OFF (muted)
DEFAULT_OSC_VALUE_FOR_ON = 1
DEFAULT_OSC_VALUE_FOR_OFF = 0


TRUTHY = {"YES", "Y", "TRUE", "T", "1", "ON"}
FALSY = {"NO", "N", "FALSE", "F", "0", "OFF", ""}
APP_SETTINGS_FILE = "theatre_settings.json"


# ==============================
# Utility Functions
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
    # Maps Excel column order to X32 channels 1..N
    return {actor: i + 1 for i, actor in enumerate(actors)}


def find_startup_excel(script_dir, remembered_path):
    candidates = []

    if remembered_path:
        remembered_path = remembered_path.strip()
        if remembered_path:
            candidates.append(remembered_path)
            if not os.path.isabs(remembered_path):
                candidates.append(os.path.join(script_dir, remembered_path))

    for base_dir in [script_dir, os.getcwd()]:
        if not os.path.isdir(base_dir):
            continue
        for file_name in os.listdir(base_dir):
            lower_name = file_name.lower()
            if lower_name.endswith(".xlsx") or lower_name.endswith(".xls"):
                candidates.append(os.path.join(base_dir, file_name))

    existing = []
    seen = set()
    for candidate in candidates:
        normalized = os.path.abspath(candidate)
        if normalized in seen:
            continue
        seen.add(normalized)
        if os.path.isfile(normalized):
            existing.append(normalized)

    if not existing:
        return None

    if remembered_path:
        remembered_abs = os.path.abspath(remembered_path)
        if remembered_abs in existing:
            return remembered_abs

    existing.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return existing[0]


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
# Main Application
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

        self.df = None
        self.scenes = {}
        self.scene_names = []
        self.actors = []
        self.channel_map = {}
        self.current_scene_index = 0
        self.last_live_state = None

        self.live_mode = tk.BooleanVar(value=False)
        self.settings_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            APP_SETTINGS_FILE
        )
        self.card_frames = {}
        self.card_name_labels = {}

        self.card_font_size = 12
        self.card_width = 140
        self.card_height = 90

        self.last_excel_path = None

        self.build_ui()
        self.load_settings()
        self.try_load_startup_excel()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ==============================
    # UI
    # ==============================

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

        grid_container = tk.Frame(self.root, bg="black")
        grid_container.pack(fill="both", expand=True, padx=20, pady=(20, 20))

        self.grid_canvas = tk.Canvas(grid_container, bg="black", highlightthickness=0)
        self.grid_canvas.pack(side="top", fill="both", expand=True)

        self.x_scrollbar = tk.Scrollbar(grid_container, orient="horizontal", command=self.grid_canvas.xview)
        self.x_scrollbar.pack(side="bottom", fill="x")
        self.grid_canvas.configure(xscrollcommand=self.x_scrollbar.set)

        self.grid_frame = tk.Frame(self.grid_canvas, bg="black")
        self.grid_canvas_window = self.grid_canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")

        self.grid_frame.bind("<Configure>", self.on_grid_frame_configure)

        self.root.bind("<Left>", lambda e: self.previous_scene())
        self.root.bind("<Right>", lambda e: self.next_scene())
        self.root.bind("<space>", lambda e: self.apply_scene())
        self.root.bind("<Return>", lambda e: self.apply_scene())

    # ==============================
    # Excel Loading
    # ==============================

    def load_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return

        self.load_excel_from_path(path)

    def load_excel_from_path(self, path, show_error_dialog=True):
        if not path:
            return False

        try:
            self.df = load_excel_file(path)
            self.scenes = dataframe_to_scene_dict(self.df)

            if self.df.empty or len(self.df.columns) == 0:
                raise ValueError("Excel file has no scenes/actors to load")

        except Exception as e:
            error_message = f"Could not load Excel file: {path}\n{e}"
            self.status_label.config(text=error_message)
            logging.warning(error_message)
            if show_error_dialog:
                messagebox.showerror("Error", str(e))
            return False

        self.scene_names = list(self.scenes.keys())
        self.actors = list(self.df.columns)
        self.channel_map = build_channel_map(self.actors)

        self.current_scene_index = 0
        self.last_live_state = None
        self.last_excel_path = os.path.abspath(path)

        self.build_grid()
        self.render_scene()
        self.save_settings()

        loaded_message = f"Loaded Excel: {os.path.basename(path)}"
        self.status_label.config(text=loaded_message)
        logging.info("%s", loaded_message)
        return True

    def try_load_startup_excel(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        path = find_startup_excel(script_dir, self.last_excel_path)
        if not path:
            self.status_label.config(text="No startup Excel found")
            return

        loaded = self.load_excel_from_path(path, show_error_dialog=False)
        if not loaded:
            self.status_label.config(text=f"Startup Excel failed: {os.path.basename(path)}")

    # ==============================
    # Grid
    # ==============================

    def build_grid(self):

        for widget in self.grid_frame.winfo_children():
            widget.destroy()

        self.card_frames.clear()
        self.card_name_labels.clear()

        for i, actor in enumerate(self.actors):
            frame = tk.Frame(self.grid_frame, bg="#111111", bd=2, relief="ridge")
            frame.grid(row=0, column=i, padx=10, pady=10)

            name = tk.Label(
                frame,
                text=self.format_actor_name(actor),
                fg="white",
                bg="#990000",
                font=("Helvetica", self.card_font_size, "bold"),
                justify="center"
            )
            name.pack(fill="both", expand=True, padx=4, pady=4)

            self.card_frames[actor] = frame
            self.card_name_labels[actor] = name

        self.update_card_sizes()

    def on_grid_frame_configure(self, _event):
        self.grid_canvas.configure(scrollregion=self.grid_canvas.bbox("all"))

    def format_actor_name(self, actor):
        actor = str(actor).strip()
        if not actor:
            return ""

        max_chars_per_line = max(4, (self.card_width - 16) // max(self.card_font_size - 2, 1))
        if len(actor) <= max_chars_per_line:
            return actor

        words = actor.split()
        if len(words) > 1:
            lines = []
            current = ""
            for word in words:
                next_candidate = f"{current} {word}".strip()
                if len(next_candidate) <= max_chars_per_line:
                    current = next_candidate
                else:
                    if current:
                        lines.append(current)
                    current = word
                if len(lines) == 1:
                    break

            remaining_words = words[len(" ".join(lines + ([current] if current else [])).split()):]
            second_line_base = current if current else ""
            if remaining_words:
                second_line_base = f"{second_line_base} {' '.join(remaining_words)}".strip()

            if lines:
                line1 = lines[0]
                line2 = second_line_base
                if len(line2) > max_chars_per_line:
                    line2 = f"{line2[:max_chars_per_line - 1]}…"
                return f"{line1}\n{line2}"

        split_at = max_chars_per_line
        line1 = actor[:split_at].rstrip()
        line2 = actor[split_at:].lstrip()
        if len(line2) > max_chars_per_line:
            line2 = f"{line2[:max_chars_per_line - 1]}…"
        return f"{line1}\n{line2}"

    def update_card_sizes(self):
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
                text=self.format_actor_name(actor),
                justify="center"
            )

        self.grid_canvas.update_idletasks()
        self.grid_canvas.configure(scrollregion=self.grid_canvas.bbox("all"))

    def increase_font_size(self):
        self.card_font_size += 1
        self.update_card_sizes()
        self.save_settings()

    def decrease_font_size(self):
        self.card_font_size -= 1
        self.update_card_sizes()
        self.save_settings()

    def increase_card_width(self):
        self.card_width += 10
        self.update_card_sizes()
        self.save_settings()

    def decrease_card_width(self):
        self.card_width -= 10
        self.update_card_sizes()
        self.save_settings()

    def increase_card_height(self):
        self.card_height += 10
        self.update_card_sizes()
        self.save_settings()

    def decrease_card_height(self):
        self.card_height -= 10
        self.update_card_sizes()
        self.save_settings()

    def load_settings(self):
        if not os.path.exists(self.settings_path):
            return

        try:
            with open(self.settings_path, "r", encoding="utf-8") as settings_file:
                settings = json.load(settings_file)

            width = int(settings.get("width", 1100))
            height = int(settings.get("height", 800))
            x = int(settings.get("x", 100))
            y = int(settings.get("y", 100))

            self.card_width = int(settings.get("card_width", self.card_width))
            self.card_height = int(settings.get("card_height", self.card_height))
            self.card_font_size = int(settings.get("card_font_size", self.card_font_size))

            self.root.geometry(f"{width}x{height}+{x}+{y}")

            last_excel = settings.get("last_excel_path", "")
            if isinstance(last_excel, str) and last_excel.strip():
                self.last_excel_path = last_excel.strip()
        except Exception as e:
            logging.warning("Could not load app settings: %s", e)

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
            with open(self.settings_path, "w", encoding="utf-8") as settings_file:
                json.dump(settings, settings_file, indent=2)
        except Exception as e:
            logging.warning("Could not save app settings: %s", e)

    def on_close(self):
        self.save_settings()
        self.root.destroy()

    # ==============================
    # Scene Navigation
    # ==============================

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

        scene = self.scene_names[self.current_scene_index]
        self.scene_label.config(text=f"Scene: {scene}")

        for actor, enabled in self.scenes[scene].items():
            frame = self.card_frames[actor]
            label = self.card_name_labels[actor]
            if enabled:
                frame.config(bg="#008000")
                label.config(bg="#008000")
            else:
                frame.config(bg="#990000")
                label.config(bg="#990000")

    # ==============================
    # Apply Scene (Send OSC)
    # ==============================

    def apply_scene(self):

        if not self.scene_names:
            return

        scene = self.scene_names[self.current_scene_index]
        scene_state = self.scenes[scene]

        self.render_scene()

        if not self.live_mode.get():
            logging.info("Preview scene: %s", scene)
            return

        sender = X32Sender(
            DEFAULT_OSC_IP,
            DEFAULT_OSC_PORT,
            DEFAULT_ADDRESS_TEMPLATE,
            self.channel_map
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

        logging.info("LIVE scene: %s | Changes: %d", scene, changes)


# ==============================
# Main
# ==============================

def main():
    root = tk.Tk()
    app = TheatreApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
