# 0362ul89t — PPTX metadata and transition consumer return

Date: 2026-10-25. This is a source-only compiler repair record; it makes no
runtime, performance, qualification, or acceptance claim. The retained
keep-going failure remains open until the composed candidate is checked again.

## Retained actual F5 failure

The isolated capture is retained outside Git at
`(deleted-coord-artifact)`.
It uses source commit
`15d8d53a0cf6d27f350b866463b4235bf26b08d9` (tree
`a619b5d0c36cd5e4a8b088bea1d0a806d13b45d9`) and the approved finite F5
driver. The formatting row passed, then
`cargo check --keep-going --offline --locked -p litchi-pptx --all-targets --all-features`
exited 101 after 72.379 seconds. The retained compiler output contains 130
diagnostics. The failed target reached 226,458,369 bytes. No later gate row is
inferred to pass.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 100,739 | `a31129804840cddda84361dc4201852856d55afc25de7f0d49cc7d5868295084` |
| `diagnostics-index.json` | 18,586 | `c25bee0ed0c486cd54e2d132c928509bfb08484991437136c53c403104b5d3be` |
| `rows.json` | 31,646 | `91083c05298999731b7f1169b433bfb12dc3642d5bbd17f613b59c9c32b60aeb` |
| `result.json` | 33,327 | `2bbc5513db14610cef06bebbeed636c8388c0cb2051d02570e4719fa774de266` |

## Source repair

Commit `254d43fbd6e15da6edc036c9065e99df3109f2fb`, based directly on the
retained `15d8d53a0cf6d27f350b866463b4235bf26b08d9` failure, changes only:

- Changes and Revision codec tests now use their already imported
  `SourceLimits` name instead of an unnecessarily qualified path. Selected
  source limits and package behavior are unchanged.
- Guides depth tests call the bounded free `rewrite_source` codec helper
  directly. The helper still receives the selected `Limits`; no removed
  associated API or unbounded fallback is restored.
- Three transition tests now pass `Limits::DEFAULT` to `read`. All production
  transition reads and writes already take explicit caller-selected limits;
  this only repairs test consumers and preserves their wire/fidelity checks.

The external source patch is retained at
`(deleted-coord-artifact)/metadata-consumers-89t-254d43fbd6.patch`
(4,689 bytes, SHA-256
`7749fc3a2bc5b0842b6c2e1faad92529a4f0bb0259fe838fc825aeb4f540d9fe`).
Assigned diagnostic and call-site scans, direct source parsing with rustfmt,
`git diff --check`, repository size guards, and reverse patch validation
passed. Cargo, rustdoc, and runtime tests were not run.

The next composed candidate still requires a fresh isolated F5 check. This
record does not reclassify the retained failure or claim acceptance.
