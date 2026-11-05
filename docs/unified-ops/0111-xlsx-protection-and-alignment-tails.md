# 0111 — XLSX protection and remaining alignment attributes

Cell and range writes now accept `alignment.relativeIndent`,
`alignment.justifyLastLine`, `protection.locked`, `protection.hidden` and
`quotePrefix`. OfficeCLI aliases `locked` and `formulahidden` are admitted.
Relative indentation uses the full signed 32-bit domain, matching OfficeCLI
and the local Open XML SDK schema. Null removes an authored attribute.

Native `ProtectionChange` and worksheet protection/quote-prefix staging join
font and alignment requests in the shared stylesheet lowering pass. Disjoint
protection fields can join; overlapping fields or quote-prefix assignments
conflict. Complete style replacement, reset and removal clear deferred changes.
The source-preserving codec retains opaque markup, quote delimiters and original
resources, sets each apply flag independently, and detects semantic no-ops.

The canonical native `Alignment::justify_last_line` field now uses `Option<bool>`
to retain explicit false separately from absence. This intentional alpha API
change also makes stylesheet serialization preserve that authored value.
Unified readback retains explicit zero relative indentation and false last-line
justification, matching OfficeCLI's extended alignment projection. Existing
quote-prefix false omission remains unchanged. XLSB explicitly reports the two
new alignment read properties as unsupported.

Validation:

- All 982 combined DOCX/XLSX/XLSB/encryption tests and doctests passed offline.
- Both encrypted and plain capability runs passed seven checks each. The OfficeCLI delta was
  regenerated and checked, and all 14 generator tests passed.
- Independent review found no actionable blockers in admission, transaction
  composition, native rewriting or read projection.
- Initial native failures were fixture assertions: self-closing versus expanded
  font markup and original style records retained alongside rewritten clones.
  The corrected checks verify original and cloned opaque content separately
  through exact counts and native readback.

The rewritten base remains an ancestor; no rebase replay is needed against
that current base. The full goal and performance acceptance remain open.

Final native checks passed all 56 stylesheet codec, 39 cell-edit and 73 semantic
style tests. These include explicit false/absence round trips, atomic mixed
formatting writes, replay/inverse, join conflicts, source preservation and exact
output limits with long namespace prefixes.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib raw::styles_edit::tests
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::cells
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib style::
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,encryption
```

Initial Clippy identified redundant cloning of the `Copy` protection request
and a manual optional-value mapping. These expressions now use direct copies
and `Option::copied`.

Final warning-denied all-target Clippy passed after also simplifying a
single-element test loop. All 50 facade cell-write tests passed again after
these final edits. Scoped formatting and whitespace checks passed. The final
OfficeCLI delta check is current with both capability snapshots.

The changed-file artifact audit found no oversized evidence or compiled
binaries. Both mandatory staged-file checks printed nothing. Staging failed:
Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock` because the
shared metadata is read-only. No commit was created; verified changes remain
in the working tree.

The borrowed Clippy fixture, generated root Cargo.lock, Python cache, temporary
logs and approximately 20 GiB of build outputs were removed. No temporary
worktree was created. The full goal remains active.
