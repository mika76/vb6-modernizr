# VB6 Modernizr — Roadmap

Feature plan, ordered by value-for-effort. Effort: S (< half a day),
M (a day-ish), L (multiple days). Check items off as they land.

## Done (v1.0)

- [x] MDI window tabs (VB6-style, overflow dropdown, middle-click close,
      context menu, forced-maximize, type glyphs)
- [x] Docked find/replace bar (Ctrl+F, Enter/Shift+Enter, F3/Shift+F3,
      Esc; case/word/regex; scopes: module/selection/open/project)
- [x] Highlight All: outline boxes in the editor + scrollbar marks,
      Procedure View aware
- [x] Find in Files (disk scan incl. designer sections, jump to line,
      header-offset compensation)
- [x] Mouse wheel scrolling in code panes (Shift+wheel = horizontal)
- [x] Discoverability: every command in the Modernizr menu with its
      shortcut in the caption; cheat-sheet window on Ctrl+Shift+/

## Phase 1 — editing quality of life (all ride the existing keyboard hook)

- [x] **Ctrl+Tab MRU window switcher** (M) — most-recently-used order,
      popup list while Ctrl held (frmSwitcher + modMRU); Ctrl+Shift+Tab
      cycles backwards, Esc cancels, releasing Ctrl commits.
- [x] **Duplicate line** Ctrl+D (S) — modEditOps.
- [x] **Move line up/down** Alt+Up / Alt+Down (S) — modEditOps.
- [x] **Delete line** Ctrl+Shift+K (S) — modEditOps.
- [x] **Comment / uncomment selection** Ctrl+/ (S) — modEditOps;
      note: bound to VK_OEM_2, i.e. the `/` key on US-style layouts.
- [x] **Find all references** Shift+F12 (M) — whole-word project
      search on the word under the cursor, results in frmRefs
      (double-click / Enter jumps).

## Phase 2 — navigation

- [ ] **Highlight word under cursor** (S) — Ctrl+W (or automatic on
      double-click): run Highlight All on the current word.
- [ ] **Go to definition** F12 (M) — regex scan of project components
      for `Sub|Function|Property|Const|Type|Enum <word>`, jump to hit.
      (Shift+F12 references: done, see Phase 1.)
- [ ] **Persistent bookmarks** (M) — F2 toggle, stored per-project in
      an `.ini` beside the `.vbp`; second color in the scrollbar marks.
- [ ] **TODO / procedure browser** (M) — dockable list (Find in Files
      window pattern) of procs and `' TODO:` comments, click to jump.

## Phase 3 — git integration

All via shelling out to `git.exe` (must be on PATH; no linked
dependency). Needs one shared piece first:

- [ ] **Git plumbing** (M) — run git with captured stdout without
      blocking the IDE (hidden process + temp-file redirect, polled by
      the existing refresh timer); cache per-file results, invalidate
      on save. Detect repo root from the project folder; degrade to
      "no git" silently.
- [ ] **Branch + dirty state in the tab bar** (S) — right-aligned
      `branch-name *` label from `git status --porcelain -b`.
- [ ] **Changed-file dots on tabs** (S) — color the tab glyph when the
      component's file is modified vs HEAD.
- [ ] **Changed-line markers** (L) — diff `CodeModule` contents against
      `git show HEAD:file` (header offset compensated, same math as
      Find in Files); paint margin/scrollbar marks via the overlay
      painter. VS-style green/blue bars.
- [ ] **Changes window** (M) — list of modified files (reuse results
      window pattern), double-click to open; optional simple commit
      (message box -> `git add` + `git commit`).
- [ ] **Blame current line** (M) — `git blame -L n,n --porcelain`,
      shown in the find bar status area or a tooltip.

Repo hygiene: `.gitattributes` forcing CRLF for VB6 sources and
marking `*.frx` etc. binary — done.

## Phase 4 — tab bar extras

- [ ] Dirty marker `*` on tabs via `VBComponent.IsDirty` (S)
- [ ] Drag to reorder tabs (M)
- [ ] "Copy full path" on the tab context menu (S)

## Phase 5 — bigger swings

- [ ] **Auto-backup** (M) — timer-driven zip/copy of the project folder
      every N minutes into a `.backups` dir with rotation.
- [ ] **Indentation guides** (L) — vertical dotted lines per indent
      level in the WM_PAINT overlay (geometry is calibrated now, but
      needs to repaint cleanly while typing).

## Not planned (tried/considered and rejected)

- Line numbers in the margin — the overlay repaints too visibly for
  something permanently on screen.
- Dark mode — the IDE hard-codes colors in too many places.
- Custom intellisense — VB6's own completion is too deeply wired in;
  fighting it risks IDE crashes.

## Known limitations / debt

- Highlight geometry is estimated (registry font + margin heuristic);
  no horizontal-scroll detection, real tab characters shift boxes.
- Highlights go stale on edit until re-run (acceptable; VS clears on
  edit too).
- Regex is single-line only (VBScript.RegExp over per-line text).
