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
- [ ] Mouse-wheel scrolling inside ucList and the git log graph
      (needs a subclass hook; keyboard + scrollbar work today).

## Conventions when adding a feature

- Sub in the right module → case in `modActions.DoAction` → `AddBtn`
  line in `Connect.Dsr` → row in `frmShortcuts` → note in README.
- Key handling goes in `modWheel.HandleKeyDown` (WM_SYSKEYDOWN for
  Alt combos), gated on `FocusInCodePane`.
- Paint overlays go through `modHighlight` (hooks are shared).
- Shared colors and the shell-icon cache live in `modTheme`
  (`THEME_*` constants, `IconForFile`/`IconForComponent`); lists that
  need icons use the custom-drawn `ucList` control, not `VB.ListBox`
  (stock ListBoxes cannot be owner-drawn after creation).
- After editing any `.frm/.bas/.cls/.dsr/.vbp` outside the IDE,
  re-normalize to CRLF/ANSI or VB6 won't load it.
