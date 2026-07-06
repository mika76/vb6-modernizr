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

- [x] **Highlight word under cursor** (S) — Ctrl+F3 (Ctrl+W is native
      Edit Watch); runs Highlight All on the current word, project-wide.
- [x] **Go to definition** F12 (M) — regex scan for definition-shaped
      lines (procs/props/consts/types/enums/events/declares/vars);
      one hit jumps, several open the results window. Native Shift+F2
      remains untouched as a fallback.
- [x] **Persistent bookmarks** (M) — Ctrl+F2 toggle, F2 next (F2
      overrides Object Browser only while focus is in a code window;
      Shift+F2 stays native). Stored in "<project>.vbp.bookmarks"
      beside the .vbp with a line-text snapshot so bookmarks re-find
      their line after edits. Blue margin squares + scrollbar marks.
- [x] **TODO / procedure browser** (M) — Ctrl+Shift+O (frmBrowser):
      procedures or TODO comments of the active project with live
      filter; double-click / Enter jumps.

## Phase 3 — git integration

All via shelling out to `git.exe` (must be on PATH; no linked
dependency). Needs one shared piece first:

- [x] **Git plumbing** (M) — modGit: repo root found by walking up
      from the project folder (no process); async one-job-at-a-time
      runner (hidden `cmd /c git ... > tempfile`, completion polled by
      the tab-bar timer via GetExitCodeProcess); status every ~5 s;
      degrades silently without git/repo.
- [x] **Branch + dirty state in the tab bar** (S) — gray right-aligned
      `branch *` label.
- [x] **Changed-file dots on tabs** (S) — orange dot by the type glyph
      when any of the component's files is modified.
- [x] **Changed-line markers** (L) — implemented via `git diff -U0`
      hunks (git does the diffing) mapped to module lines with the
      header offset; margin bars green=added, blue=modified,
      red=deletion-below; also colored on the scrollbar. Reflects the
      SAVED file vs HEAD (unsaved editor changes not included).
- [x] **Changes window** (M) — Ctrl+Shift+G: staged/unstaged lists
      from the porcelain XY codes; stage/unstage selected (multi-
      select) or all (`git add` / `git reset -q HEAD`); Commit commits
      the staged set (`git commit -F tempfile`); double-click opens
      the file.
- [x] **Log with graph** (M) — Ctrl+Shift+L: `git log --graph` with a
      tab-separated pretty format in a monospace list (git draws the
      graph), optional --all, click a commit for `git show --stat`
      details.
- [x] **Blame current line** (M) — Ctrl+Shift+B: author/date/summary
      message box via `git blame -L n,n --porcelain`.

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
