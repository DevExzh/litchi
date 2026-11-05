# 0362ugk — PPTX lazy edit entry

This source-only repair moves only the primary PPTX edit entry point to the
edit-only lazy state required by the 0362 final sweep. `Package::edit` and
`edit_with_options` retain the cheap OPC graph, limits, owner policy, and the
package's payload-digest memo sink; they no longer capture the opened
presentation, scan the slide catalog and notes, build names, or fingerprint
the complete package at construction. Dropping an untouched edit therefore
does no presentation capture. The first source getter, semantic getter,
mutation, or commit materializes the existing validated capture path and
returns its ordinary typed error. A failed capture restores the deferred
state, so retries do not observe a partial edit.

`Package::opened_presentation*` remains the explicit eager, fallible read
projection. `Snapshot::edit` now also keeps its package graph, slide catalog,
and name index behind shared ownership: cloning the edit root is O(1), and the
first mutating operation obtains the private OPC graph with copy-on-write.
The edit's source, slides, and retained-MCE getters are fallible where
materialization is required; every direct consumer (including masters,
layouts, themes, SmartArt, notes, pictures, text, and hyperlinks) resolves
that state before reading it. Releasing retained MCE before materializing is
remembered and applied to the captured snapshot. The existing digest memo is
projected back to the package's own payload allocations, preserving the memo
fast path for later edits and replay. Signed admission, no-op publication,
patch/replay behavior, source provenance, owner policy, and output limits stay
on the existing capture/commit paths.

The constructor/drop regression uses the test fingerprint log to prove that
entry and drop do not capture, while the first `slides` use does. The unified
read regression likewise constructs the edit successfully and keeps its
malformed-tail refusal at that first fallible `slides` operation.

## Validation receipt

The final validation ran on 2026-10-24 from
`/home/zhuhe/code/litchi-worktrees/uo2-pptx-lazy-entry` with the committed
source below. There was no source or test overlay. Every Cargo command sourced
`(deleted-coord-artifact)`, which pins
`CARGO_TARGET_DIR` to
`(deleted-coord-artifact)`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`, release-like debug settings,
offline mode, the memory-capped Cargo shim, and the assigned `TMPDIR`.
The only ignored input overlay was the generated `Cargo.lock` required by this
checkout: `cargo generate-lockfile --offline`, 139,592 bytes, SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`; it was
not committed. All commands used `--offline --locked` with that exact lock.

Accepted commands and their exact source commits were:

- `c10d6c65398dfe75f35a0691a795210f65993392`:
  `cargo test --offline --locked -p litchi-pptx --all-features --all-targets`;
  exit 0, 1,837 passed, 0 failed and one existing ignored golden-transcript
  helper. Raw log
  `pptx-all-features-all-targets-c10d6c6539.log`, 201,479 bytes, SHA-256
  `165d07553e25c09dcb159e4bc3b0c18dbbee70b831a678fc2bf63959fd3837c1`.
  Assigned target bytes were 2,378,275,297 at start and 2,378,227,693 at end.
- `c10d6c65398dfe75f35a0691a795210f65993392`:
  `cargo clippy --offline --locked -p litchi-pptx --all-features --all-targets -- -D warnings`;
  exit 0. Raw log `pptx-clippy-all-features-all-targets-c10d6c6539.log`, 60
  bytes, SHA-256
  `d24957620a3002bd023dc3a6dcb3600988bbf75aeef94579c398e8bb5b6aa545`.
- `c10d6c65398dfe75f35a0691a795210f65993392`:
  `cargo test --offline --locked -p litchi-ops --all-features --test pptx_unified_read -- --nocapture`;
  exit 0, 11 passed and 0 failed. Raw log
  `ops-pptx-unified-read-c10d6c6539.log`, 12,497 bytes, SHA-256
  `9eea561b681429746063440b694e4162ef600af6146c70b58b636b5196987189`.
  Assigned target bytes were 2,378,227,692 at start and 4,887,534,491 at end.
- `fa63bdd22e4e4c650f81d1216f5e219b81bfd2b2` (the comment-only successor to
  the validated `c10d6c6539` source):
  `cargo test --offline --locked -p litchi-pptx --all-features --doc`;
  exit 0, 48 passed, 0 failed and two existing ignored documentation examples.
  Raw log
  `pptx-doc-all-features-fa63bdd22e.log`, 5,950 bytes, SHA-256
  `1ed3418ac42785f3f60f2741a1d3dfa3522c8568c830d08d98ec872fc61e79ea`.
  Assigned target bytes were 4,887,534,491 at start and 4,887,534,493 at end.

The raw logs and their small receipts are retained outside the checkout at
`(deleted-coord-artifact)`.
The final assigned target size was 4,887,534,493 bytes, below the 20 GB cap.
The earlier intermediate logs in that directory are diagnostic history, not
acceptance evidence.

One focused diagnostic was accidentally launched without sourcing the assigned
environment: from this worktree, the command was
`cargo test --offline --locked -p litchi-pptx --all-features --test source_primary_text_box caller_output_limits_precede_managed_memory_admission -- --nocapture`
at source `15b148707646010693fd74d66a4245aae5166000`. It therefore used the
default `target/` directory and the ordinary `unoptimized + debuginfo` profile,
not the memory-capped shim or assigned target; this run is not acceptance
evidence. It created 2,903,268,714 bytes at
`/home/zhuhe/code/litchi-worktrees/uo2-pptx-lazy-entry/target`. That generated
cache was removed with the scoped command
`cargo clean --target-dir /home/zhuhe/code/litchi-worktrees/uo2-pptx-lazy-entry/target`
(`Removed 4525 files, 2.9GiB total`). No tracked files were removed. No
further target directories are to be deleted; all accepted checks above used
the assigned 20 GB-capped target.
