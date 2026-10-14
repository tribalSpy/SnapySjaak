# Uniware recorder

Maps the Uniware program (the old keyboard-driven ERP: menus, F-keys,
Shift/Ctrl/Alt combinations) by watching you use it. Separate from the
SnappySjaak app; runs on your own PC, nothing is sent anywhere.

While you work in Uniware as usual, it records for every key or key
combination you press:

    screen before  --key(s)-->  screen after

and reads each screen with Windows' own text recognition (built into
Windows 10/11). The result is a map of all menus and screens you visited:
what's on them, which item was selected, and which keys lead where.

This is step 1 (mapping). Step 2 -- recording one action and replaying it
automatically, with "wait until this screen shows X" checks -- builds on the
same screen reading.

## Install (once)

1. Python 3.12 (python.org; tick "Add python.exe to PATH"). The Windows text
   recognition package needs Python 3.9 - 3.12.
2. Double-click **install.bat**.

## Record

1. Double-click **run_recorder.bat** (a window with the instructions opens).
2. Click in the Uniware window, then press **Ctrl+Alt+R**: from now on the
   recorder watches that window.
3. Work in Uniware as usual: walk through the menus you use, open screens,
   press the F-keys you normally use. Every step shows up in the recorder
   window, e.g.
   `[Hoofdmenu] --Enter on "Fust administratie"--> [Fust administratie]`.
4. **Ctrl+Alt+P** pauses / resumes (e.g. while typing a password -- typed
   text on screens with "wachtwoord"/"password" is never stored anyway).
   **Ctrl+Alt+S** reads the current screen without a key press.
5. **Ctrl+Alt+Q** stops and saves.

Record as often as you like: every session adds to the same map.

## See the map

Double-click **make_report.bat**: it opens `data\map.html` with every screen
(screenshot + text), the keys that lead from it to other screens, and a
diagram of the whole map.

## After an update of this tool

Double-click **rebuild_map.bat**: re-reads every screenshot recorded so far
with the current screen recognition (old map kept as
`data\map.before-rebuild.json`), then make_report.bat again.

## How screens are named

Uniware puts a number on every screen: menus show **Prg. NNN** top right
(000 Hoofdmenu, 260 Inkoop, 400 Voorraad...), forms and lists show
**Title [NN]** in their top bar (Muteren Labels [38]). Those numbers are
the screen names in the map. The terminal program's own title, menu bar
("Bestand Edit Beeld...") and status bar are ignored, and the window is
captured on its own -- another window on top of it doesn't get in.

## Files

- `data\map.json` -- the map (all sessions together)
- `data\menu_overview.md` -- every menu with its options and the route to it
- `data\screens\` -- one screenshot per screen
- `data\sessions\` -- step-by-step log per recording session
- `config.json` (optional, copy from `config.example.json`) -- which screens
  hide typed text, how long to wait for a screen to finish drawing.

Only keys pressed **in the Uniware window** are recorded, nothing from
other programs.
