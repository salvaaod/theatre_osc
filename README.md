# theatre_osc

Desktop controller for theatre microphone cues on an X32 console using OSC and an Excel cue sheet.

## What changed

- Mic cards now resize to better fit the available horizontal space.
- The settings file is now named `theatre_settings.json`.
- The app now saves the path of the last Excel file loaded.
- On startup, the app tries to auto-load the last Excel file if it still exists.

## Excel format

- First column: scene names (used as row index).
- Remaining columns: actor/mic names.
- Cell values are interpreted as ON/OFF states. Supported truthy values: `YES, Y, TRUE, T, 1, ON`; falsy values: `NO, N, FALSE, F, 0, OFF, (empty)`.

## Run

```bash
python3 teatro_osc.py
```

## Notes

- Application/window state is stored in `theatre_settings.json` next to `teatro_osc.py`.
- If the remembered Excel path is not available anymore, startup continues normally without loading a file.
