#!/usr/bin/env python3

import json
import logging
import os
import sys

import pandas as pd
from pythonosc.udp_client import SimpleUDPClient
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMenuBar,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)


DEFAULT_OSC_IP = "192.168.10.59"
DEFAULT_OSC_PORT = 10023
DEFAULT_ADDRESS_TEMPLATE = "/ch/{ch:02d}/mix/on"
DEFAULT_OSC_VALUE_FOR_ON = 1
DEFAULT_OSC_VALUE_FOR_OFF = 0

TRUTHY = {"YES", "Y", "TRUE", "T", "1", "ON"}
FALSY = {"NO", "N", "FALSE", "F", "0", "OFF", ""}
APP_SETTINGS_FILE = "theatre_settings.json"

CARD_SPACING = 6
CARD_BORDER_WIDTH = 4
CARD_ON_COLOR = "#f0f0f0"
CARD_OFF_COLOR = "#b00020"
CARD_TEXT_ON_COLOR = "#111111"
CARD_TEXT_OFF_COLOR = "#ffffff"


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


def split_first_space(text):
    text = str(text).strip()
    if " " in text:
        return text.replace(" ", "\n", 1)
    return text


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
            if name.lower().endswith((".xlsx", ".xls")):
                candidates.append(os.path.join(scan_dir, name))

    existing = []
    seen = set()
    for path in candidates:
        abs_path = os.path.abspath(path)
        if abs_path in seen:
            continue
        seen.add(abs_path)
        if os.path.isfile(abs_path):
            existing.append(abs_path)

    if not existing:
        return None

    existing.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return existing[0]


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


class Card(QFrame):
    def __init__(self, text, size):
        super().__init__()
        self.raw_text = text
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)

        self.set_size(size)
        self.set_muted(True)

    def set_size(self, size):
        self.setFixedSize(size, size)
        self.label.setText(split_first_space(self.raw_text))

        font = QFont()
        font.setPointSize(10 if size <= 60 else 14)
        font.setBold(True)

        self.label.setFont(font)
        self.label.setMaximumWidth(size - 8)

    def set_muted(self, muted):
        bg = CARD_OFF_COLOR if muted else CARD_ON_COLOR
        text = CARD_TEXT_OFF_COLOR if muted else CARD_TEXT_ON_COLOR
        self.setStyleSheet(
            f"""
            QFrame {{
                border: {CARD_BORDER_WIDTH}px solid #555;
                border-radius: 4px;
                background: {bg};
            }}
            QLabel {{
                border: none;
                color: {text};
            }}
            """
        )


class TheatreApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("X32 Theatre Mic Controller")

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

        self.card_size = 80
        self.osc_ip = DEFAULT_OSC_IP
        self.osc_port = DEFAULT_OSC_PORT
        self.cards = []

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(10, 10, 10, 10)

        self.build_ui()
        self.load_settings()
        self.try_load_startup_excel()
        self.adjust_window()

    def build_ui(self):
        self.menu_bar = QMenuBar()
        self.main_layout.setMenuBar(self.menu_bar)

        file_menu = self.menu_bar.addMenu("File")
        load_action = QAction("Load Excel", self)
        load_action.triggered.connect(self.load_excel)
        file_menu.addAction(load_action)

        settings_menu = self.menu_bar.addMenu("Settings")
        size_menu = settings_menu.addMenu("Card Size")
        for s in [60, 80]:
            action = QAction(f"{s}px", self)
            action.triggered.connect(lambda checked=False, v=s: self.set_card_size(v))
            size_menu.addAction(action)

        osc_menu = settings_menu.addMenu("OSC")
        ip_action = QAction("Set IP", self)
        ip_action.triggered.connect(self.set_osc_ip)
        osc_menu.addAction(ip_action)

        port_action = QAction("Set Port", self)
        port_action.triggered.connect(self.set_osc_port)
        osc_menu.addAction(port_action)

        controls = QHBoxLayout()
        self.main_layout.addLayout(controls)

        button_size = (140, 56)

        self.prev_btn = QPushButton("Previous")
        self.prev_btn.setFixedSize(*button_size)
        self.prev_btn.clicked.connect(self.previous_scene)
        controls.addWidget(self.prev_btn)

        self.next_btn = QPushButton("Next")
        self.next_btn.setFixedSize(*button_size)
        self.next_btn.clicked.connect(self.next_scene)
        controls.addWidget(self.next_btn)

        self.take_btn = QPushButton("Take")
        self.take_btn.setFixedSize(*button_size)
        self.take_btn.clicked.connect(self.apply_scene)
        controls.addWidget(self.take_btn)

        controls.addStretch()

        self.scene_label = QLabel("No Scene")
        self.scene_label.setAlignment(Qt.AlignCenter)
        self.scene_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.main_layout.addWidget(self.scene_label)

        self.row_layout = QHBoxLayout()
        self.row_layout.setSpacing(CARD_SPACING)
        self.main_layout.addLayout(self.row_layout)

        self.status_label = QLabel("No Excel loaded")
        self.status_label.setAlignment(Qt.AlignLeft)
        self.main_layout.addWidget(self.status_label)

        QShortcut(QKeySequence(Qt.Key_Left), self, activated=self.previous_scene)
        QShortcut(QKeySequence(Qt.Key_Right), self, activated=self.next_scene)
        QShortcut(QKeySequence(Qt.Key_Space), self, activated=self.apply_scene)
        QShortcut(QKeySequence(Qt.Key_Return), self, activated=self.apply_scene)

    def load_settings(self):
        if not os.path.exists(self.settings_path):
            return
        try:
            with open(self.settings_path, "r", encoding="utf-8") as f:
                settings = json.load(f)

            self.card_size = int(settings.get("card_size", self.card_size))
            self.last_excel_path = str(settings.get("last_excel_path", "")).strip()
            self.osc_ip = str(settings.get("osc_ip", self.osc_ip)).strip() or self.osc_ip
            self.osc_port = int(settings.get("osc_port", self.osc_port))
        except Exception as exc:
            logging.warning("Could not load app settings: %s", exc)

    def save_settings(self):
        payload = {
            "card_size": self.card_size,
            "last_excel_path": self.last_excel_path,
            "osc_ip": self.osc_ip,
            "osc_port": self.osc_port,
        }
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2)
        except Exception as exc:
            logging.warning("Could not save app settings: %s", exc)

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Excel",
            "",
            "Excel files (*.xlsx *.xls)"
        )
        if path:
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

            self.rebuild_cards()
            self.draw_current_scene()
            self.status_label.setText(
                f"Loaded Excel: {os.path.basename(path)} | OSC {self.osc_ip}:{self.osc_port}"
            )
            self.save_settings()
            logging.info("Loaded Excel: %s", self.last_excel_path)
            return True

        except Exception as exc:
            message = f"Could not load Excel: {path} | {exc}"
            self.status_label.setText(message)
            logging.warning(message)
            if show_error_dialog:
                QMessageBox.critical(self, "Excel load error", str(exc))
            return False

    def try_load_startup_excel(self):
        base_dir = os.path.dirname(os.path.abspath(__file__))
        startup = find_startup_excel(base_dir, self.last_excel_path)
        if not startup:
            self.status_label.setText("No startup Excel found")
            return

        if not self.load_excel_from_path(startup, show_error_dialog=False):
            self.status_label.setText(f"Startup Excel failed: {os.path.basename(startup)}")

    def rebuild_cards(self):
        while self.row_layout.count():
            item = self.row_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        self.cards = []
        for actor in self.actors:
            card = Card(actor, self.card_size)
            self.cards.append(card)
            self.row_layout.addWidget(card)

    def draw_current_scene(self):
        if not self.scene_names:
            self.scene_label.setText("No Scene")
            self.adjust_window()
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.scenes[scene_name]
        self.scene_label.setText(f"Scene: {scene_name}")

        for actor, card in zip(self.actors, self.cards):
            card.set_size(self.card_size)
            enabled = bool(scene_state.get(actor, False))
            card.set_muted(not enabled)

        self.adjust_window()

    def set_card_size(self, size):
        self.card_size = int(size)
        self.draw_current_scene()
        self.save_settings()

    def set_osc_ip(self):
        value, ok = QInputDialog.getText(self, "Set OSC IP", "IP:", text=self.osc_ip)
        if not ok:
            return
        value = value.strip()
        if not value:
            return
        self.osc_ip = value
        self.status_label.setText(f"OSC target: {self.osc_ip}:{self.osc_port}")
        self.save_settings()

    def set_osc_port(self):
        value, ok = QInputDialog.getInt(
            self,
            "Set OSC Port",
            "Port:",
            value=self.osc_port,
            minValue=1,
            maxValue=65535,
        )
        if not ok:
            return
        self.osc_port = int(value)
        self.status_label.setText(f"OSC target: {self.osc_ip}:{self.osc_port}")
        self.save_settings()

    def adjust_window(self):
        self.adjustSize()
        self.setFixedSize(self.sizeHint())

    def previous_scene(self):
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index - 1) % len(self.scene_names)
        self.draw_current_scene()

    def next_scene(self):
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index + 1) % len(self.scene_names)
        self.draw_current_scene()

    def apply_scene(self):
        if not self.scene_names:
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.scenes[scene_name]

        sender = X32Sender(
            self.osc_ip,
            self.osc_port,
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
        logging.info("TAKE scene: %s | Changes: %d", scene_name, changes)
        self.status_label.setText(
            f"TAKE sent: {scene_name} | Changes: {changes} | OSC {self.osc_ip}:{self.osc_port}"
        )

    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    window = TheatreApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
