# VB6 Modernizr

A Visual Basic 6 IDE add-in, written in VB6 itself, that adds a few
modern conveniences to the classic IDE. No external dependencies:
everything uses the VB6 runtime, the extensibility library that ships
with VB6, Win32, and `VBScript.RegExp` (part of Windows) for regex.

## Features

- **MDI window tabs** — a browser-style tab bar above the code/designer
  area, drawn in the classic VB6 3D look. Left click activates a
  window, middle click closes it, right click offers Close / Close
  Others / Close All, and the `▼` button on the right lists every
  window when the tabs overflow. A small colored square marks the
  window type (blue = code, green = designer). Toggle it from
  *Modernizr → Show/Hide Tabs*.
- **Find / Replace** (*Modernizr → Find / Replace...*) — search the
  current module, the selection, all open modules, or the whole active
  project. Match case / whole word / **regular expressions** (with
  `$1`-style group references in replacements). **Highlight All**
  outlines every match in the editor and draws orange tick marks on
  the code window's vertical scrollbar at each match position.
- **Find in Files** (*Modernizr → Find in Files...*) — scans the whole
  codebase on disk, including designer sections of `.frm`/`.ctl`/`.dsr`
  files and project files, under the active project's folder or any
  folder you pick. Double-click a result to jump straight to that line
  in the IDE (files outside the loaded project group open in Notepad).
- **Mouse wheel scrolling** — the stock VB6 editor ignores the wheel;
  the add-in makes code windows scroll (Shift+wheel scrolls
  horizontally). No need for Microsoft's separate MouseWheel fix.

## Building

1. Open `VB6Modernizr.vbp` in VB6.
2. If the references show as MISSING (paths differ per machine), open
   *Project → References* and re-check **Microsoft Visual Basic 6.0
   Extensibility** and **Microsoft Add-In Designer**.
3. *File → Make VB6Modernizr.dll*. Compiling registers the COM DLL and
   the add-in designer writes the add-in registration automatically.
   On Windows Vista+ run VB6 **as administrator** for this step so
   registration succeeds.

## Installing / loading

1. Restart VB6.
2. *Add-Ins → Add-In Manager…* — select **VB6 Modernizr**, check
   *Loaded/Unloaded* and *Load on Startup*.
3. If it does not appear in the list, add this line to
   `C:\Windows\vbaddin.ini` under `[Add-Ins32]` (create the section if
   needed) and restart VB6:

   ```ini
   [Add-Ins32]
   VB6Modernizr.Connect=1
   ```

To uninstall: unload it in the Add-In Manager, then
`regsvr32 /u VB6Modernizr.dll`.

## Development tips

- Don't run the add-in from a second IDE instance *and* keep a compiled
  copy loaded at the same time — unload the compiled one first.
- The tab bar and highlight painting rely on window subclassing. All
  subclass procs are guarded with `On Error Resume Next` and are
  removed on disconnect, but if you're hacking on those parts, test in
  a second VB6 instance (F5 from a first instance) so a crash doesn't
  take your editor down.

## Known limitations

- **Highlight geometry is approximate.** The VB6 editor has no
  highlight API, so match boxes are computed from
  `CodePane.TopLine`/`CountOfVisibleLines` plus the editor font read
  from the registry. Horizontal editor scrolling can't be detected, so
  boxes assume the view starts at column 1. Real tab characters in
  source lines also shift box positions.
- Replace operations work line-by-line; regex patterns cannot span
  multiple lines.
- The tab bar reserves its strip by adjusting the MDI client during
  the IDE's own layout passes; extremely unusual window managers or
  other add-ins that also move the MDI client could conflict.

## Ideas / roadmap

- Ctrl+Tab MRU window switcher (most-recently-used order, like modern
  editors)
- Bookmarks panel with named bookmarks that survive IDE restarts
- Procedure/TODO browser (jump list built from comments and procs)
- Code snippets with placeholders
- Auto-backup: zip the project folder every N minutes
- Duplicate-line / move-line-up/down keyboard shortcuts
- Indentation guides drawn in the code pane (same overlay technique as
  Highlight All)
- Changed-line markers in the margin when the project is under git
