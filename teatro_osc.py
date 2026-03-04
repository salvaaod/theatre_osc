#!/usr/bin/env python3

import os
import time
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from pythonosc.udp_client import SimpleUDPClient


# ==============================
# X32 CONFIGURATION
# ==============================

DEFAULT_OSC_IP = "192.168.10.59"   # Change to your X32 IP
DEFAULT_OSC_PORT = 10023         # X32 default OSC port

DEFAULT_ADDRESS_TEMPLATE = "/ch/{ch:02d}/mix/on"

# X32 logic:
# 1 = ON (audio passes)
# 0 = OFF (muted)
DEFAULT_OSC_VALUE_FOR_ON = 1
DEFAULT_OSC_VALUE_FOR_OFF = 0


TRUTHY = {"YES", "Y", "TRUE", "T", "1", "ON"}
FALSY = {"NO", "N", "FALSE", "F", "0", "OFF", ""}


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

        self.build_ui()

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
            bg="black"
        )
        self.scene_label.pack(pady=10)

        controls = tk.Frame(self.root, bg="black")
        controls.pack()

        tk.Button(controls, text="Previous", command=self.previous_scene).pack(side="left", padx=5)
        tk.Button(controls, text="Next", command=self.next_scene).pack(side="left", padx=5)
        tk.Button(controls, text="GO", command=self.apply_scene, width=10).pack(side="left", padx=10)

        self.grid_frame = tk.Frame(self.root, bg="black")
        self.grid_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.mic_labels = {}

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

        try:
            self.df = load_excel_file(path)
            self.scenes = dataframe_to_scene_dict(self.df)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.scene_names = list(self.scenes.keys())
        self.actors = list(self.df.columns)
        self.channel_map = build_channel_map(self.actors)

        self.current_scene_index = 0
        self.last_live_state = None

        self.build_grid()
        self.render_scene()

        logging.info("Loaded Excel: %s", os.path.basename(path))

    # ==============================
    # Grid
    # ==============================

    def build_grid(self):

        for widget in self.grid_frame.winfo_children():
            widget.destroy()

        self.mic_labels.clear()

        cols = 4
        for i, actor in enumerate(self.actors):
            row = i // cols
            col = i % cols

            frame = tk.Frame(self.grid_frame, bg="#111111", bd=2, relief="ridge")
            frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")

            name = tk.Label(frame, text=actor, fg="white", bg="#111111",
                            font=("Helvetica", 16, "bold"))
            name.pack(pady=5)

            state = tk.Label(frame, text="OFF", bg="red",
                             font=("Helvetica", 18, "bold"), width=8)
            state.pack(pady=5)

            self.mic_labels[actor] = state

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
            label = self.mic_labels[actor]
            if enabled:
                label.config(text="ON", bg="green")
            else:
                label.config(text="OFF", bg="red")

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
    root.geometry("1100x800")
    app = TheatreApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()


