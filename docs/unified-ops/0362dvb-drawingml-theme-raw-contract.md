# 0362dvb — DrawingML Theme-family raw mutation limits

- Date: 2026-10-24.
- Parent source checkpoint: `11728db894181ffbcdbcea3ae0008d51fdd2d8af`
  (`f036e87f21439a7d69f63adcf107a241d2cf4e65`).
- Source commit: `20c7dd4d4bd80c603eadae89626309d979f1f138`
  (`42d98afbb8634c0188b3cdc3140417fca0b772ca`).
- Status: source repair and focused validation complete; broader DrawingML raw
  codec ingress/output families remain separately scoped work.

## Contract repair

The public Theme-family mutation surface now has one required policy type,
`litchi_drawingml::theme::family::part::Limits`, carrying the caller-selected
`max_xml_bytes` bound. The four `Snapshot` methods and four free functions use
the canonical names and require that policy:

`replace_family`, `add_family`, `add_family_with_uri`, and `remove_family`.

The public limit-less twins and the four `_with_limit` spellings were removed;
the private implementation helpers retain the bounded splice logic. XLSB
Theme lifecycle, source rewrite, and writer consumers pass the retained XLSB
byte policy through `part::Limits`, and all five DrawingML Theme integration
tests were migrated in this same source commit. The glossary profile bans
`replace_family_with_limit`, `add_family_with_limit`,
`add_family_with_uri_limit`, and `remove_family_with_limit` at the Theme part
module.

The action scalar validator now classifies an embedded NUL as an invalid
lexical value and reports only a true over-budget scalar as the typed
`LimitKind::ScalarBytes` error. The regression test covers both cases.

This record does not claim closure for the remaining DrawingML chart, text,
diagram, model3d, SVG, InkML, or other raw codec policy lanes.

## Validation

The commands used the pinned workspace lock
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`
(139,591 bytes), one capped build job, offline locked dependencies, debug
information disabled, and the external target
`/home/zhuhe/code/litchi-worktrees/.coord/theme-family/target` (816,108,921
bytes after completion). Both commands exited 0:

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/theme-family/tmp cargo test -p litchi-drawingml --all-features --offline --locked --test ink_action_edit --test ink_action_profile --test theme_family_foreign_uri --test theme_family_mce --test theme_family_ownership --test theme_family_xml_grammar --test theme_part --no-fail-fast --target-dir /home/zhuhe/code/litchi-worktrees/.coord/theme-family/target
```

23 action-edit tests, 21 action-profile tests, 1 foreign-URI test, 4 MCE
tests, 7 ownership tests, 13 XML-grammar tests, and 22 complete-theme tests
passed. The retained log is
`drawingml-theme-limits-rerun3.log` (9,115 bytes,
SHA-256 `22073f1494be02da4501e6f0cb9023ee23d3742c672e8ad600591ccbb8b57b1e`).

```text
env -u LITCHI_CAPPED -u LITCHI_BLESS PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH LITCHI_MEM_MAX=20G CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_PROFILE_RELEASE_DEBUG=0 CARGO_PROFILE_BENCH_DEBUG=0 TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/theme-family/tmp cargo test -p litchi-xlsb --all-features --offline --locked --lib theme --no-fail-fast --target-dir /home/zhuhe/code/litchi-worktrees/.coord/theme-family/target
```

The XLSB theme-filtered library suite passed 12 tests (767 filtered). Its
retained log is `xlsb-theme-limits-rerun.log` (1,708 bytes,
SHA-256 `d6ea9893c4270fdafc1dde480fc32478c178277429725d7073243af7a829a550`).

`cargo fmt --all -- --check`, `git diff --check`, the staged large-evidence
guard, the staged binary guard, and the JSON parse of `tools/api_glossary.json`
also passed. No binary or large evidence file is committed.
