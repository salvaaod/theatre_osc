# theatre_osc

Desktop controller for theatre microphone cues on an X32 console using OSC and an Excel cue sheet.

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

- `PySide6`
- `pandas`
- `python-osc`

## Notes

- Application/window state is stored in `theatre_settings.json` next to `teatro_osc.py`.
- If the remembered Excel path is not available anymore, startup continues normally without loading a file.
