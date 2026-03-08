# theatre_osc

**OSC Theater Mic Controller** is a desktop tool for running theater microphone cues from an Excel cue sheet over OSC.
It is compatible with digital mixers that follow the same OSC channel naming/address conventions as Behringer/Midas consoles.
It has been tested with:

- Midas **M32**
- Behringer **X32**
- Behringer **XR18/AR18**

Default OSC ports by mixer family:

- **M32 / X32:** `10023`
- **XR18 / AR18:** `10024`

## What changed

- The GUI now uses a Qt (`PySide6`) card-based layout (no Tkinter canvas drawing).
- Window size auto-fits the cards and is fixed (not user-resizable).
- Active channels are shown with neutral cards; muted channels are red.
- Scene controls are `Previous`, `Next`, and `Take` (`Take` sends OSC changes).
- Card size and OSC IP/port are configurable from the Settings menu.

## Excel format

- First column: scene names (used as row index).
- Remaining columns: actor/mic names.
- Cell values are interpreted as ON/OFF states. Supported truthy values: `YES, Y, TRUE, T, 1, ON`; falsy values: `NO, N, FALSE, F, 0, OFF, (empty)`.

## Run

```bash
python3 teatro_osc.py
```

## Dependencies

- Python 3.9+ (recommended)
- `PySide6`
- `pandas`
- `python-osc`

Install with pip:

```bash
python3 -m pip install PySide6 pandas python-osc
```

## Notes

- Application/window state is stored in `theatre_settings.json` next to `teatro_osc.py`.
- If the remembered Excel path is not available anymore, startup continues normally without loading a file.
