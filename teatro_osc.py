#!/usr/bin/env python3

import argparse
import json
import logging
import os
import re
import sys
import threading

import pandas as pd
from pythonosc.dispatcher import Dispatcher
from pythonosc.osc_message_builder import OscMessageBuilder
from pythonosc.osc_server import ThreadingOSCUDPServer
from pythonosc.udp_client import SimpleUDPClient
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QAction, QActionGroup, QFont, QFontMetrics, QKeySequence, QShortcut
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
DEFAULT_OSC_LISTEN_IP = "0.0.0.0"
DEFAULT_ADDRESS_TEMPLATE = "/ch/{ch:02d}/mix/on"
DEFAULT_OSC_VALUE_FOR_ON = 1
DEFAULT_OSC_VALUE_FOR_OFF = 0

TRUTHY = {"YES", "Y", "TRUE", "T", "1", "ON"}
FALSY = {"NO", "N", "FALSE", "F", "0", "OFF", ""}
APP_SETTINGS_FILE = "theatre_settings.json"

CARD_SPACING = 6
CARD_BORDER_WIDTH = 2
CARD_ON_COLOR = "#f0f0f0"
CARD_OFF_COLOR = "#b00020"
CARD_TEXT_ON_COLOR = "#111111"
CARD_TEXT_OFF_COLOR = "#ffffff"
CARD_MISMATCH_BORDER_COLOR = "#f4d03f"

BASE_CARD_SIZE = 80
BASE_CONTROL_BUTTON_WIDTH = 112
BASE_CONTROL_BUTTON_HEIGHT = 56
BASE_CONTROL_BUTTON_FONT_SIZE = 14
XREMOTE_KEEPALIVE_INTERVAL_MS = 9000

def setup_logging(debug=False):
    level = logging.DEBUG if debug else logging.INFO
    root = logging.getLogger()
    root.setLevel(level)

    for handler in list(root.handlers):
        root.removeHandler(handler)

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_handler = logging.FileHandler("show_log.txt", encoding="utf-8")
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    if debug:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        root.addHandler(console_handler)

    logging.info("Logging initialised | debug=%s", debug)


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

    def query_channel(self, actor):
        ch = self.actor_channel_map[actor]
        address = self.address_template.format(ch=ch)
        self.client.send_message(address, [])
        logging.info("OSC query: %s", address)


class Card(QFrame):
    clicked = Signal(str)

    def __init__(self, text, size):
        super().__init__()
        self.actor = text
        self.raw_text = text
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.setContentsMargins(0, 0, 0, 0)
        self.setLayout(layout)

        self.set_size(size)
        self.set_muted(True)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.actor)
        super().mousePressEvent(event)

    def set_size(self, size):
        self.setFixedSize(size, size)
        self.label.setText(split_first_space(self.raw_text))

        font = QFont()
        font.setPointSize(10 if size <= 60 else 14)
        font.setBold(True)

        self.label.setFont(font)
        self.label.setMaximumWidth(size - 8)

    def set_muted(self, muted, mismatch=False):
        bg = CARD_OFF_COLOR if muted else CARD_ON_COLOR
        text = CARD_TEXT_OFF_COLOR if muted else CARD_TEXT_ON_COLOR
        border = CARD_MISMATCH_BORDER_COLOR if mismatch else "#555"
        self.setStyleSheet(
            f"""
            QFrame {{
                border: {CARD_BORDER_WIDTH}px solid {border};
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
    channel_status_received = Signal(int, bool)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("X32 Theatre Mic Controller")

        self.settings_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            APP_SETTINGS_FILE
        )

        self.df = None
        self.scenes = {}
        self.scene_names = []
        self.scene_override = None
        self.scene_override_name = ""
        self.actors = []
        self.channel_map = {}
        self.current_scene_index = 0
        self.last_live_state = None
        self.current_live_state = {}
        self.mismatch_actors = set()
        self.manual_override_actors = set()
        self.last_excel_path = ""
        self.saved_window_position = None
        self.always_visible = False
        self.osc_listener = None
        self.osc_listener_thread = None

        self.card_size = 80
        self.osc_ip = DEFAULT_OSC_IP
        self.osc_port = DEFAULT_OSC_PORT
        self.cards = []
        self.control_buttons = []
        self.pending_take = False
        self.card_edit_unlocked = False
        self.bulk_toggle_scene_name = ""
        self.bulk_toggle_target = None
        self.bulk_toggle_snapshot = None
        self.force_full_send_next_take = True
        self.take_blink_on = False
        self.take_blink_timer = QTimer(self)
        self.take_blink_timer.setInterval(450)
        self.take_blink_timer.timeout.connect(self.toggle_take_blink)

        self.xremote_timer = QTimer(self)
        self.xremote_timer.setInterval(XREMOTE_KEEPALIVE_INTERVAL_MS)
        self.xremote_timer.timeout.connect(self.send_xremote_keepalive)

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(10, 10, 10, 10)

        self.build_ui()
        self.channel_status_received.connect(self.handle_channel_status_update)
        self.load_settings()
        self.start_osc_listener()
        self.update_control_button_sizes()
        self.try_load_startup_excel()
        self.adjust_window()
        self.apply_saved_window_position()

    def build_ui(self):
        self.menu_bar = QMenuBar()
        self.main_layout.setMenuBar(self.menu_bar)

        file_menu = self.menu_bar.addMenu("File")
        self.load_excel_action = QAction("Load Excel", self)
        self.load_excel_action.triggered.connect(self.load_excel)
        file_menu.addAction(self.load_excel_action)

        self.always_visible_menu = self.menu_bar.addMenu("Always Visible")
        self.always_visible_group = QActionGroup(self)
        self.always_visible_group.setExclusive(True)

        self.always_visible_on_action = QAction("On", self)
        self.always_visible_on_action.setCheckable(True)
        self.always_visible_on_action.triggered.connect(
            lambda checked=False: self.set_always_visible(True)
        )
        self.always_visible_group.addAction(self.always_visible_on_action)
        self.always_visible_menu.addAction(self.always_visible_on_action)

        self.always_visible_off_action = QAction("Off", self)
        self.always_visible_off_action.setCheckable(True)
        self.always_visible_off_action.triggered.connect(
            lambda checked=False: self.set_always_visible(False)
        )
        self.always_visible_group.addAction(self.always_visible_off_action)
        self.always_visible_menu.addAction(self.always_visible_off_action)
        self.sync_always_visible_menu_actions()

        settings_menu = self.menu_bar.addMenu("Settings")
        size_menu = settings_menu.addMenu("Card Size")
        self.card_size_group = QActionGroup(self)
        self.card_size_group.setExclusive(True)
        self.card_size_actions = {}
        for s in [60, 80]:
            action = QAction(f"{s}px", self)
            action.setCheckable(True)
            action.triggered.connect(lambda checked=False, v=s: self.set_card_size(v))
            self.card_size_group.addAction(action)
            self.card_size_actions[s] = action
            size_menu.addAction(action)

        osc_menu = settings_menu.addMenu("OSC")
        self.ip_action = QAction("Set IP", self)
        self.ip_action.triggered.connect(self.set_osc_ip)
        osc_menu.addAction(self.ip_action)

        self.port_action = QAction("Set Port", self)
        self.port_action.triggered.connect(self.set_osc_port)
        osc_menu.addAction(self.port_action)

        read_action = QAction("Read Channels", self)
        read_action.triggered.connect(self.read_channels_from_mixer)
        osc_menu.addAction(read_action)

        self.controls_layout = QHBoxLayout()
        self.main_layout.addLayout(self.controls_layout)

        button_style = "QPushButton { border: 2px solid #555; border-radius: 4px; }"
        all_off_button_style = "QPushButton { border: 2px solid #c00000; border-radius: 4px; }"

        self.prev_btn = QPushButton("Previous")
        self.prev_btn.setStyleSheet(button_style)
        self.prev_btn.clicked.connect(self.previous_scene)
        self.controls_layout.addWidget(self.prev_btn)
        self.control_buttons.append(self.prev_btn)

        self.next_btn = QPushButton("Next")
        self.next_btn.setStyleSheet(button_style)
        self.next_btn.clicked.connect(self.next_scene)
        self.controls_layout.addWidget(self.next_btn)
        self.control_buttons.append(self.next_btn)

        self.take_btn = QPushButton("Take")
        self.take_btn.clicked.connect(self.apply_scene)
        self.controls_layout.addWidget(self.take_btn)
        self.control_buttons.append(self.take_btn)

        self.controls_layout.addStretch()

        self.scene_label = QLabel("No Scene")
        self.scene_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.scene_label.setStyleSheet("font-size: 20px; font-weight: bold;")
        self.configure_scene_label_width()
        self.controls_layout.addWidget(self.scene_label)

        self.controls_layout.addStretch()

        self.all_on_btn = QPushButton("ALL ON")
        self.all_on_btn.setStyleSheet(button_style)
        self.all_on_btn.clicked.connect(lambda: self.set_all_for_current_scene(True))
        self.controls_layout.addWidget(self.all_on_btn)
        self.control_buttons.append(self.all_on_btn)

        self.all_off_btn = QPushButton("ALL OFF")
        self.all_off_btn.setStyleSheet(all_off_button_style)
        self.all_off_btn.clicked.connect(lambda: self.set_all_for_current_scene(False))
        self.controls_layout.addWidget(self.all_off_btn)
        self.control_buttons.append(self.all_off_btn)

        self.controls_cards_gap = QWidget()
        self.controls_cards_gap.setFixedHeight(0)
        self.main_layout.addWidget(self.controls_cards_gap)

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

        self.update_control_button_sizes()
        self.update_take_button_style()
        self.update_menu_state()

    def update_menu_state(self):
        if hasattr(self, "card_size_actions"):
            for size, action in self.card_size_actions.items():
                action.blockSignals(True)
                action.setChecked(size == int(self.card_size))
                action.blockSignals(False)

        if hasattr(self, "ip_action"):
            self.ip_action.setText(f"Set IP ({self.osc_ip})")

        if hasattr(self, "port_action"):
            self.port_action.setText(f"Set Port ({self.osc_port})")

        if hasattr(self, "load_excel_action"):
            excel_name = os.path.basename(self.last_excel_path) if self.last_excel_path else "none"
            self.load_excel_action.setText(f"Load Excel ({excel_name})")

    def configure_scene_label_width(self):
        placeholder = "SCENE: " + ("W" * 20)
        metrics = QFontMetrics(self.scene_label.font())
        width = metrics.horizontalAdvance(placeholder) + 20
        self.scene_label.setFixedWidth(width)

    def update_control_button_sizes(self):
        scale = max(0.5, self.card_size / BASE_CARD_SIZE)
        button_width = int(round(BASE_CONTROL_BUTTON_WIDTH * scale))
        button_height = int(round(BASE_CONTROL_BUTTON_HEIGHT * scale))
        font_size = max(9, int(round(BASE_CONTROL_BUTTON_FONT_SIZE * scale)))

        for button in self.control_buttons:
            button.setFixedSize(button_width, button_height)
            font = button.font()
            font.setPointSize(font_size)
            button.setFont(font)

        self.update_controls_cards_gap()

    def update_controls_cards_gap(self):
        scale = max(0.5, self.card_size / BASE_CARD_SIZE)
        gap = max(4, int(round(12 * (scale ** 2))))
        self.controls_cards_gap.setFixedHeight(gap)

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
            self.always_visible = bool(settings.get("always_visible", self.always_visible))
            window_x = settings.get("window_x")
            window_y = settings.get("window_y")
            if window_x is not None and window_y is not None:
                self.saved_window_position = (int(window_x), int(window_y))

            self.sync_always_visible_menu_actions()
            self.set_always_visible(self.always_visible, save=False)
            self.update_menu_state()
        except Exception as exc:
            logging.warning("Could not load app settings: %s", exc)

    def save_settings(self):
        payload = {
            "card_size": self.card_size,
            "last_excel_path": self.last_excel_path,
            "osc_ip": self.osc_ip,
            "osc_port": self.osc_port,
            "always_visible": self.always_visible,
            "window_x": int(self.pos().x()),
            "window_y": int(self.pos().y()),
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
            self.scene_override = None
            self.scene_override_name = ""
            self.actors = list(df.columns)
            self.channel_map = build_channel_map(self.actors)
            self.current_scene_index = 0
            self.last_live_state = None
            self.current_live_state = {}
            self.force_full_send_next_take = True
            self.mismatch_actors.clear()
            self.manual_override_actors.clear()
            self.last_excel_path = os.path.abspath(path)

            self.rebuild_cards()
            self.draw_current_scene()
            self.set_take_pending(True)
            self.status_label.setText(
                f"Loaded Excel: {os.path.basename(path)} | OSC {self.osc_ip}:{self.osc_port}"
            )
            self.update_menu_state()
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
            card.clicked.connect(self.toggle_actor_for_current_scene)
            self.cards.append(card)
            self.row_layout.addWidget(card)

    def card_enabled_state_for_display(self, actor, expected_enabled):
        if actor in self.manual_override_actors:
            return bool(expected_enabled)
        if actor in self.mismatch_actors and actor in self.current_live_state:
            return bool(self.current_live_state[actor])
        return bool(expected_enabled)

    def refresh_cards_from_scene(self, scene_state):
        for actor, card in zip(self.actors, self.cards):
            card.set_size(self.card_size)
            expected_enabled = bool(scene_state.get(actor, False))
            display_enabled = self.card_enabled_state_for_display(actor, expected_enabled)
            card.set_muted(not display_enabled, actor in self.mismatch_actors)

    def draw_current_scene(self):
        if not self.scene_names:
            self.scene_label.setText("No Scene")
            self.adjust_window()
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.get_scene_state(scene_name)
        if scene_state is None:
            self.scene_label.setText("No Scene")
            return
        self.scene_label.setText(f"SCENE: {scene_name}")

        self.refresh_cards_from_scene(scene_state)

        self.adjust_window()

    def set_card_size(self, size):
        self.card_size = int(size)
        self.update_control_button_sizes()
        self.draw_current_scene()
        self.update_menu_state()
        self.save_settings()

    def set_osc_ip(self):
        value, ok = QInputDialog.getText(self, "Set OSC IP", "IP:", text=self.osc_ip)
        if not ok:
            return
        value = value.strip()
        if not value:
            return
        self.osc_ip = value
        self.send_xremote_keepalive()
        self.status_label.setText(f"OSC target: {self.osc_ip}:{self.osc_port}")
        self.update_menu_state()
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
        self.restart_osc_listener()
        self.status_label.setText(f"OSC target/readback: {self.osc_ip}:{self.osc_port}")
        self.update_menu_state()
        self.save_settings()

    def sync_always_visible_menu_actions(self):
        self.always_visible_on_action.blockSignals(True)
        self.always_visible_off_action.blockSignals(True)
        self.always_visible_on_action.setChecked(self.always_visible)
        self.always_visible_off_action.setChecked(not self.always_visible)
        self.always_visible_on_action.blockSignals(False)
        self.always_visible_off_action.blockSignals(False)

    def set_always_visible(self, enabled, save=True):
        enabled = bool(enabled)
        self.always_visible = enabled
        self.sync_always_visible_menu_actions()

        window_handle = self.windowHandle()
        if window_handle is not None:
            window_handle.setFlag(Qt.WindowStaysOnTopHint, enabled)
        else:
            was_visible = self.isVisible()
            self.setWindowFlag(Qt.WindowStaysOnTopHint, enabled)
            if was_visible:
                self.show()

        if save:
            self.save_settings()

    def adjust_window(self):
        self.adjustSize()
        self.setFixedSize(self.sizeHint())

    def apply_saved_window_position(self):
        if self.saved_window_position is None:
            return
        x, y = self.saved_window_position
        self.move(x, y)

    def clear_bulk_toggle_state(self):
        self.bulk_toggle_scene_name = ""
        self.bulk_toggle_target = None
        self.bulk_toggle_snapshot = None

    def previous_scene(self):
        if self.bulk_toggle_interaction_locked():
            return
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index - 1) % len(self.scene_names)
        self.draw_current_scene()
        self.card_edit_unlocked = False
        self.clear_bulk_toggle_state()
        self.manual_override_actors.clear()
        self.set_take_pending(True)

    def next_scene(self):
        if self.bulk_toggle_interaction_locked():
            return
        if not self.scene_names:
            return
        self.current_scene_index = (self.current_scene_index + 1) % len(self.scene_names)
        self.draw_current_scene()
        self.card_edit_unlocked = False
        self.clear_bulk_toggle_state()
        self.manual_override_actors.clear()
        self.set_take_pending(True)

    def get_scene_state(self, scene_name):
        base_state = self.scenes.get(scene_name)
        if base_state is None:
            return None

        if self.scene_override_name != scene_name or self.scene_override is None:
            self.scene_override_name = scene_name
            self.scene_override = dict(base_state)

        return self.scene_override

    def current_scene_state(self):
        if not self.scene_names:
            return None
        scene_name = self.scene_names[self.current_scene_index]
        return self.get_scene_state(scene_name)

    def live_reference_state(self, scene_state):
        reference = {}
        for actor in scene_state:
            if actor in self.current_live_state:
                reference[actor] = bool(self.current_live_state[actor])
            elif self.last_live_state is not None:
                reference[actor] = bool(self.last_live_state.get(actor, scene_state[actor]))
            else:
                reference[actor] = bool(scene_state[actor])
        return reference

    def has_pending_changes(self):
        scene_state = self.current_scene_state()
        if scene_state is None:
            return False
        if self.last_live_state is None and not self.current_live_state:
            return True

        reference = self.live_reference_state(scene_state)
        return any(reference.get(actor) != bool(enabled) for actor, enabled in scene_state.items())

    def toggle_actor_for_current_scene(self, actor):
        if self.bulk_toggle_interaction_locked():
            return
        if not self.card_edit_unlocked:
            return

        scene_state = self.current_scene_state()
        if scene_state is None or actor not in scene_state:
            return

        base_enabled = bool(self.current_live_state.get(actor, scene_state[actor]))
        scene_state[actor] = not base_enabled
        self.manual_override_actors.add(actor)
        self.refresh_cards_from_scene(scene_state)

        had_bulk_toggle = self.bulk_toggle_snapshot is not None
        self.clear_bulk_toggle_state()
        self.set_take_pending(self.has_pending_changes() or had_bulk_toggle)

    def set_all_for_current_scene(self, enabled):
        if not self.card_edit_unlocked:
            return

        if not self.scene_names:
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.get_scene_state(scene_name)
        if scene_state is None:
            return

        value = bool(enabled)
        has_active_bulk_toggle = (
            self.bulk_toggle_snapshot is not None
            and self.bulk_toggle_scene_name == scene_name
        )
        is_repeat_cancel = has_active_bulk_toggle and self.bulk_toggle_target == value
        is_opposite_while_pending = has_active_bulk_toggle and self.bulk_toggle_target != value

        if is_opposite_while_pending:
            active_label = "ALL ON" if self.bulk_toggle_target else "ALL OFF"
            self.status_label.setText(
                f"{active_label} pending: press TAKE to apply or press {active_label} again to cancel"
            )
            return

        if is_repeat_cancel:
            scene_state.clear()
            scene_state.update(self.bulk_toggle_snapshot)
            self.clear_bulk_toggle_state()
            self.apply_bulk_toggle_interaction_lock()
        else:
            self.bulk_toggle_scene_name = scene_name
            self.bulk_toggle_target = value
            self.bulk_toggle_snapshot = dict(scene_state)
            for actor in scene_state:
                scene_state[actor] = value

        self.manual_override_actors.update(scene_state.keys())
        self.refresh_cards_from_scene(scene_state)

        pending = self.has_pending_changes() or (
            self.bulk_toggle_snapshot is not None
            and self.bulk_toggle_scene_name == scene_name
        )
        self.set_take_pending(pending)

    def bulk_toggle_interaction_locked(self):
        return (
            self.pending_take
            and self.bulk_toggle_snapshot is not None
            and self.bulk_toggle_scene_name == self.scene_names[self.current_scene_index]
        ) if self.scene_names else False

    def apply_bulk_toggle_interaction_lock(self):
        locked = self.bulk_toggle_interaction_locked()
        allowed_target = self.bulk_toggle_target if locked else None

        self.prev_btn.setEnabled(not locked)
        self.next_btn.setEnabled(not locked)
        self.all_on_btn.setEnabled((not locked) or allowed_target is True)
        self.all_off_btn.setEnabled((not locked) or allowed_target is False)

        for card in self.cards:
            card.setEnabled(not locked)

    def toggle_take_blink(self):
        self.take_blink_on = not self.take_blink_on
        self.update_take_button_style()

    def set_take_pending(self, pending):
        self.pending_take = bool(pending)
        if self.pending_take:
            self.take_blink_on = True
            if not self.take_blink_timer.isActive():
                self.take_blink_timer.start()
        else:
            self.take_blink_timer.stop()
            self.take_blink_on = False
        self.update_take_button_style()
        self.apply_bulk_toggle_interaction_lock()

    def update_take_button_style(self):
        border = "2px solid #555"
        if self.pending_take:
            background = "#ffd84d" if self.take_blink_on else "#b38f00"
            text_color = "#111111"
        else:
            background = "#2e7d32"
            text_color = "#ffffff"

        self.take_btn.setStyleSheet(
            f"QPushButton {{ border: {border}; border-radius: 4px; background: {background}; color: {text_color}; }}"
        )

    def apply_scene(self):
        if not self.scene_names:
            return

        scene_name = self.scene_names[self.current_scene_index]
        scene_state = self.get_scene_state(scene_name)
        if scene_state is None:
            return

        sender = X32Sender(
            self.osc_ip,
            self.osc_port,
            DEFAULT_ADDRESS_TEMPLATE,
            self.channel_map,
        )

        force_full_send = self.force_full_send_next_take or (
            self.bulk_toggle_snapshot is not None
            and self.bulk_toggle_scene_name == scene_name
        )

        changes = 0
        reference = self.live_reference_state(scene_state)
        for actor, enabled in scene_state.items():
            if force_full_send or reference.get(actor) != bool(enabled):
                sender.send(actor, enabled)
                changes += 1

        self.force_full_send_next_take = False
        self.last_live_state = dict(scene_state)
        self.current_live_state = dict(scene_state)
        self.mismatch_actors.clear()
        self.manual_override_actors.clear()
        self.card_edit_unlocked = True
        self.clear_bulk_toggle_state()
        self.set_take_pending(False)
        logging.info("TAKE scene: %s | Changes: %d", scene_name, changes)
        self.status_label.setText(
            f"TAKE sent: {scene_name} | Changes: {changes} | OSC {self.osc_ip}:{self.osc_port}"
        )

    def start_osc_listener(self):
        dispatcher = Dispatcher()
        dispatcher.set_default_handler(self.on_unhandled_osc)
        dispatcher.map("/ch/*/mix/on", self.on_channel_status_osc)

        try:
            self.osc_listener = ThreadingOSCUDPServer(
                (DEFAULT_OSC_LISTEN_IP, self.osc_port),
                dispatcher,
            )
        except OSError as exc:
            logging.warning(
                "Could not start OSC listener on %s:%s (%s)",
                DEFAULT_OSC_LISTEN_IP,
                self.osc_port,
                exc,
            )
            return

        self.osc_listener_thread = threading.Thread(
            target=self.osc_listener.serve_forever,
            daemon=True,
        )
        self.osc_listener_thread.start()
        logging.info("OSC listener started on %s:%s", DEFAULT_OSC_LISTEN_IP, self.osc_port)
        self.xremote_timer.start()
        self.send_xremote_keepalive()

    def stop_osc_listener(self):
        self.xremote_timer.stop()
        if self.osc_listener is None:
            return
        self.osc_listener.shutdown()
        self.osc_listener.server_close()
        self.osc_listener = None
        self.osc_listener_thread = None

    def restart_osc_listener(self):
        self.stop_osc_listener()
        self.start_osc_listener()

    def on_channel_status_osc(self, address, *args):
        logging.debug("OSC RX matched: address=%s args=%s", address, args)
        match = re.match(r"^/ch/(\d{1,2})/mix/on$", str(address))
        if not match or not args:
            logging.debug("OSC RX ignored (invalid format or empty args): address=%s args=%s", address, args)
            return

        try:
            channel = int(match.group(1))
            value = float(args[0])
        except (TypeError, ValueError):
            return

        enabled = value >= 0.5
        logging.debug("OSC parsed channel update: channel=%s enabled=%s raw_value=%s", channel, enabled, value)
        self.channel_status_received.emit(channel, enabled)

    def handle_channel_status_update(self, channel, enabled):
        logging.debug("UI handling channel update: channel=%s enabled=%s", channel, enabled)
        actor = next((name for name, ch in self.channel_map.items() if ch == channel), None)
        if actor is None:
            logging.debug("No actor mapped for channel=%s", channel)
            return

        self.current_live_state[actor] = enabled
        if self.last_live_state is None:
            self.last_live_state = {}
        self.last_live_state[actor] = bool(enabled)
        scene_state = self.current_scene_state()
        if scene_state is None:
            return

        expected = bool(scene_state.get(actor, False))
        if expected != enabled:
            self.mismatch_actors.add(actor)
        else:
            self.mismatch_actors.discard(actor)
            self.manual_override_actors.discard(actor)

        logging.debug(
            "Comparison result: actor=%s expected=%s live=%s mismatch_count=%s",
            actor,
            expected,
            enabled,
            len(self.mismatch_actors),
        )

        self.draw_current_scene()
        has_bulk_pending = (
            self.bulk_toggle_snapshot is not None
            and self.bulk_toggle_scene_name == self.scene_names[self.current_scene_index]
        ) if self.scene_names else False
        self.set_take_pending(self.has_pending_changes() or has_bulk_pending)

        if self.mismatch_actors:
            self.status_label.setText(
                "External mixer change detected: press TAKE to restore current scene (yellow border marks mismatches)"
            )

    def send_from_listener_socket(self, address, args=()):
        if self.osc_listener is None:
            logging.warning("OSC listener not running, cannot send: %s", address)
            return False

        try:
            builder = OscMessageBuilder(address=address)
            for arg in args:
                builder.add_arg(arg)
            message = builder.build()
            self.osc_listener.socket.sendto(message.dgram, (self.osc_ip, self.osc_port))
            logging.debug(
                "OSC datagram sent via listener socket | address=%s args=%s local=%s:%s remote=%s:%s bytes=%s",
                address,
                args,
                DEFAULT_OSC_LISTEN_IP,
                self.osc_port,
                self.osc_ip,
                self.osc_port,
                len(message.dgram),
            )
            return True
        except OSError as exc:
            logging.warning("OSC send failed for %s (%s)", address, exc)
            return False

    def send_xremote_keepalive(self):
        if self.send_from_listener_socket("/xremote"):
            logging.debug("OSC /xremote keepalive sent")

    def send_query_from_listener_socket(self, address):
        if self.send_from_listener_socket(address):
            logging.info("OSC query (listener socket): %s", address)
            return True
        return False

    def read_channels_from_mixer(self):
        if not self.actors:
            return

        if self.osc_listener is None:
            self.status_label.setText("OSC listener is not active; cannot request channel states")
            logging.warning("Read Channels requested but listener is not active")
            return

        sent = 0
        for actor in self.actors:
            ch = self.channel_map.get(actor)
            if ch is None:
                continue
            address = DEFAULT_ADDRESS_TEMPLATE.format(ch=ch)
            if self.send_query_from_listener_socket(address):
                sent += 1
                logging.debug("OSC query requested for actor=%s channel=%s", actor, ch)

        self.status_label.setText(
            f"Requested channel states from mixer ({sent}) | listening on UDP {self.osc_port}"
        )

    def on_unhandled_osc(self, address, *args):
        logging.debug("OSC RX unhandled: address=%s args=%s", address, args)

    def closeEvent(self, event):
        self.stop_osc_listener()
        self.save_settings()
        super().closeEvent(event)


def main():
    parser = argparse.ArgumentParser(description="X32 Theatre Mic Controller")
    parser.add_argument("--debug", action="store_true", help="Enable verbose OSC/debug logging")
    args, qt_args = parser.parse_known_args()

    setup_logging(debug=args.debug)

    app = QApplication([sys.argv[0], *qt_args])
    window = TheatreApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
