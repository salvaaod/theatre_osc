# theatre_osc

**OSC Theater Mic Controller** is a desktop cueing tool for running theater microphone scenes from an Excel cue sheet over OSC.
It is designed for mixers that follow Behringer/Midas-style OSC channel addressing and has been tested with:

- Midas **M32**
- Behringer **X32**
- Behringer **XR18 / AR18**

Default OSC target ports by mixer family:

- **M32 / X32:** `10023`
- **XR18 / AR18:** `10024`

When running the mixer emulator or another OSC tool on the same computer, the app avoids UDP bind conflicts by sending to the configured mixer port and listening on `target_port + 10` for replies and push updates, wrapping back into the valid `1..65535` range when needed.

## Features

### Cue and scene workflow

- Load an Excel cue sheet where each row is a scene and each actor column maps to a mixer channel.
- Navigate through scenes with **Previous** and **Next**.
- Use **Take** to send the current scene to the mixer.
- Use **Clear** to discard pending scene edits and return to the last live/taken state.
- Use **ALL ON** or **ALL OFF** to stage a full-scene mute/unmute change before pressing **Take**.
- After a scene has been taken once, click cards to make per-actor manual adjustments for the current scene preview.
- Sends only changed channels by default, while forcing a full resend after startup or bulk scene operations.

### Live mixer awareness

- Shows a centered **CONNECTED / DISCONNECTED** badge in the menu bar.
- Sends `/info` probes on a timer to detect whether the mixer is reachable.
- Sends `/xremote` keepalives so compatible mixers continue pushing updates.
- Can request current channel mute states with **Settings → OSC → Read Channels**.
- Highlights mismatches between the loaded cue sheet and live mixer state with a blinking yellow border.
- Syncs external mixer-side changes back into the current in-app scene preview so the UI stays aligned with what is live.

### Channel utilities

- **Send Channel Names** pushes actor names to mixer channel labels using `/ch/XX/config/name`.
- Channel names are trimmed to the mixer-safe 12-character limit before being sent.

### Desktop UI and persistence

- Uses a Qt / `PySide6` card-based desktop UI.
- Window auto-sizes to fit the cards and remains fixed-size.
- Card size can be changed from **Settings → Card Size** (`60px` or `80px`).
- **Always Visible** menu toggles an always-on-top window mode.
- Saves settings such as Excel path, window position, OSC target, card size, send delay, and always-on-top preference.
- On startup, only the remembered Excel file is auto-loaded; the app does not scan the working directory for spreadsheets.

## Excel cue sheet format

The workbook's active sheet is used.

- **Column A:** scene names.
- **Columns B+**: actor / mic names.
- Actor headers must be non-blank and unique.
- Blank rows are ignored.
- Scene names are required for any non-empty row.

Cell values are normalized to ON/OFF booleans using these values:

- Truthy: `YES`, `Y`, `TRUE`, `T`, `1`, `ON`
- Falsy: `NO`, `N`, `FALSE`, `F`, `0`, `OFF`, empty cell
- Any other value is treated as `OFF`

Actors are mapped left-to-right to mixer channels starting at channel 1.

## Installation

### Requirements

- Python **3.9+**
- `PySide6`
- `openpyxl`
- `python-osc`

Install dependencies with pip:

```bash
python3 -m pip install PySide6 openpyxl python-osc
```

## Usage

Start the app:

```bash
python3 theatre_osc.py
```

Enable verbose logging for OSC traffic and connection debugging:

```bash
python3 theatre_osc.py --debug
```

### Typical workflow

1. Launch the app.
2. Open **File → Load Excel** and select your cue sheet.
3. Confirm the OSC IP/port in **Settings → OSC**.
4. Optionally use **Read Channels** to pull the current mixer state into the UI.
5. Step through scenes with **Previous** / **Next**.
6. Press **Take** to apply the pending scene changes.
7. If needed, use **Send Channel Names** once to copy actor names to the mixer.

### Keyboard shortcuts

- **Left Arrow:** previous scene
- **Right Arrow:** next scene
- **Space:** Take
- **Enter / Return:** Take

## Settings reference

### File

- **Load Excel**: choose the workbook used for cue playback.

### Always Visible

- **On / Off**: toggle always-on-top window behavior.

### Settings → Card Size

- **60px**
- **80px**

### Settings → OSC

- **Set IP**: change the OSC target IP.
- **Set Port**: change the OSC target port; the local listener port is automatically recalculated as `port + 10`.
- **Set Send Delay (ms)**: add a delay between outbound OSC messages (`0..50 ms`) to reduce dropped commands on some mixers or networks.
- **Read Channels**: query channel mute/on states from the mixer.
- **Send Channel Names**: write actor names to mixer channel labels.

## Files created by the app

- `theatre_settings.json`: saved UI and OSC settings, written next to `theatre_osc.py`.
- `show_log.txt`: application log file.

## Notes and behavior details

- Card background colors reflect state: neutral cards are active/on, red cards are muted/off.
- The **Take** button blinks when there are staged changes waiting to be sent.
- Bulk operations lock navigation until you either press **Take** or cancel the same bulk operation.
- If the remembered startup Excel file is missing or invalid, the app starts without loading a workbook and shows `No startup Excel found` or a startup failure message.
- OSC replies and update subscriptions depend on the listener being able to bind its local UDP port.
