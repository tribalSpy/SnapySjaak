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

## Actions dashboard (record once, start with a button)

Double-click **run_dashboard.bat**: it opens http://127.0.0.1:8765 (this PC
only).

**Record an action**
1. Put Uniware on the screen where the action should start (e.g. the main
   menu) -- every run must start there too.
2. In the dashboard: type a name (e.g. "UKdocs invoice data") and click
   **Record**. Uniware comes to the front.
3. Do the action by hand. On each screen whose data you want, press
   **Ctrl+Alt+C** (capture this screen).
4. Click **Stop and save** (or Ctrl+Alt+Q).

**Edit** (per step): tick *Capture screen* / *All pages* (Page Down until the
list ends), or give typed text a **field name** (e.g. `invoice`) -- then the
dashboard asks that value every time you start the action.

**Slow screens / loading**: while recording, just wait until a screen has
finished loading before pressing the next key -- the screen you see at that
moment is what the step leads to, and how long it took is remembered. When
running, after every key it waits until Uniware shows that screen (Prg /
[number]) and has stopped changing, up to **Max wait** (default 4x the time
it took while recording, at least 30 s; editable per step). If data appears
later on the same screen (e.g. after typing a tour), put a word that only
shows once the data is there in **Wait until screen shows** (e.g. a column
heading or "Totaal"). Waited too long -> it stops, with a screenshot in
`outputs\`; it never keeps typing into a wrong screen.

**Start**: fill in the fields, click Start and don't touch keyboard or mouse
until it says Finished. **Stop** or **Ctrl+Alt+Q** stops at any time.
Captured screens are saved in `outputs\` as text, CSV (with positions, for
tables) and JSON.

**Ctrl+P (data opens in a new window)**: a key that opens another program's
window is marked by itself while recording ("opens the dump in a new
window", also a tick box under Edit). When running, it waits for that
window, finds the file it shows (from the program's command line, or the
file name in its title), copies it with the dump file name to the folder you
give, closes the window and brings Uniware back.

**Enter in a list without arrowing** = Enter on whatever row is selected
(e.g. the first order) -- never looked for by its text. Arrowing first and
then Enter = "find row with" (e.g. tour 508); empty that box under Edit to
go back to "the selected row". In lists the selected row is the teal row;
in menus it's the blue item.

**Uniware's own dump file** (Edit, per action): Uniware saves its export
itself with a generic name (`dump.filipe.2026-10-14_16_57_30.csv`). Fill in
the folder it saves to (e.g. the OneDrive folder) and the name to give it
(`{date}`, `{time}`, `{action}` and fields allowed, e.g. `tour-508_{date}`).
After the last step the run waits for the *new* dump file, waits until it
is completely written, then renames it (and moves it, if you give another
folder). No new file within the max wait = the run fails and nothing is
renamed. Record the action including the keys that make Uniware export.

Without that setting, a **Dump folder** gets the screens captured with
Ctrl+Alt+C (`.csv` + `.txt`), only after a run that *finished*.

**Menus are routes by name**: Uniware's menu levels sit side by side (Right =
one level deeper, Left = back; the level you're in is the rightmost column
with a blue item). While recording, arrow through the menu as usual; when
you press Enter (or another key) on a menu, all those arrows become one step
like `menu Expeditie > Inklaring bedrijf, then Enter`. When running, it reads
the screen and steers there from wherever the menu is: Left back to the
first level that's wrong, Up/Down by counting rows from the blue item, Right
into the next level -- checking after every move. So an action that starts
with a menu route can start from any menu position; if Uniware is inside a
program, it first presses the **Back key** (Esc; changeable under Edit) until
a menu shows (never on a menu itself). An item that isn't there stops the run.

**Lists whose order changes** (e.g. the tours after F1): while recording,
just arrow down to the row (e.g. 508) and press Enter. The action then
remembers *the row with 508*, not "3x Down": when running it reads the list,
finds the row with 508 (whole number -- not 1508 or 50,80), counts how far
it is from the highlighted row, presses Down/Up that often (scrolling if it
is off screen) and checks 508 is highlighted before Enter. Not in the list
today -> it stops with a screenshot. Under Edit, **find row with** shows /
changes the value (a field like `{tour}` also works).

**Change without recording again** (Edit): change the **key** of a step
(e.g. Left instead of Right), **remove** a step, or **Copy** a whole action
(e.g. tour 515 from tour 508) and change its "find row with" and file name.

**Read a value from the screen into the file name** (Edit, "Read column"):
after that step, the run reads a table column on screen -- e.g. `invoiceRes`
into field `invoiceres` -- and the dump file name can use it:
`data_tour-508_{invoiceres}_{date}`. It reads twice from two screenshots and
stops if they differ (a misread number must never end up as an invoice
number); look-alike letters in numbers are corrected (483S236 -> 4835236).

**List loop -- go through a list and act per row** (Edit, "Add list loop
after step N"): the run goes through the list row by row, reads the screen
for each row and presses one key when it shows one of the match words,
another when not. Example, tour 508 = flowers: when the screen shows
"Accessoire, Planten" press **Delete** (out of the tour), keep "Bloemen",
otherwise **Down**. And in tour 515: "Accessoire" -> **Enter** (add).
- A row whose type can't be read clearly (none of the words) is skipped,
  never deleted, and listed in the log.
- Stops at the end of the list, after "at most N rows", or when a row is
  still there after the key (Uniware asking a confirmation?).
- Several keys in a row, separated by commas -- e.g. `Delete, Right, Enter`
  when Uniware asks a confirmation after Delete (the screen settles between
  the keys, so the question is there before it's answered).
- A new loop starts as a **trial run**: it only logs what it would do --
  check the log, then untick "trial run".

**Two dumps, one file** (e.g. tour 508 with two companies, 4EADMI and LTD):
record both Ctrl+P's in one action and tick "join all dumps of a run into one
file" under Dump file. Each dump is kept in `outputs/dump-parts/`; at the end
of the run they become one file (one header, all rows, itemNumber numbered
on), named then -- so `data_tour-508_{invoiceres}` takes the invoiceRes read
on the LTD screen even though that comes after the first dump. Read
invoiceRes into `invoiceres` only on the screen whose number counts.

**Day lists -- information from one action to the next** (e.g. accessoires:
SPICER -> tour 516, other -> 515). Three actions, each ending on the main
menu; the list in between is shown under "Day lists" in the dashboard:
1. **Collect** -- tour 515 "In te delen", a list loop with "collect into day
   list" = `accessoires vandaag`: every row's order number is written down
   (the list is emptied first), only Down is pressed.
2. **Check** -- "Add for each order" (day list `accessoires vandaag`, the
   next N steps), then those steps: Ctrl+F, type `{order}`, Enter, F11, a
   **check** step (a line shows `SPICER` -> `516`, lines without -> `515`),
   back. An order with both kinds gets a warning "multiple types of
   accessoires" and no result -- set it by hand in the Day lists block.
   "Only orders without a result yet" lets a stopped run continue.
3. **Assign** -- a list loop with "act on day list" `accessoires vandaag`,
   orders marked `516` -> Enter (in tour 516's list), otherwise Down. Done
   orders get a time in the "Done" column.

**Tours (UKdocs)**: record one action per tour (508 flowers, 515
accessories, 1000 BQ, ...) -- some tours need an extra key, so each has its
own recording -- each with its own dump file name.

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
