# VB6 Modernizr — Plan

All five phases of the original roadmap shipped (v1.1): tabs with
drag-reorder/dirty/git markers, docked find/replace bar with regex +
highlight-all, find in files, navigation (definition, references,
bookmarks, code browser, MRU switcher), editing shortcuts, full git
integration (status, staging, commit, log graph, blame, line markers),
mouse wheel scrolling, indentation guides, and auto-backup.
See README.md for the full feature reference and known limitations.

## Backlog (open, not yet scheduled)

- [ ] Persist custom tab order across IDE restarts (store the key
      list in the project sidecar file, like bookmarks).
- [ ] Calibrate overlay geometry against the actual editor metrics
      instead of the empirical one-cell margin shift, if further
      off-by-a-bit reports come in.

## Conventions when adding a feature

- Sub in the right module → case in `modActions.DoAction` → `AddBtn`
  line in `Connect.Dsr` → row in `frmShortcuts` → note in README.
- Key handling goes in `modWheel.HandleKeyDown` (WM_SYSKEYDOWN for
  Alt combos), gated on `FocusInCodePane`.
- Paint overlays go through `modHighlight` (hooks are shared).
- After editing any `.frm/.bas/.cls/.dsr/.vbp` outside the IDE,
  re-normalize to CRLF/ANSI or VB6 won't load it.
