# Evidence for change 0759: the spec-gap branch merge

This packet backs [record 0759](../../0759-spec-gap-branch-merge.md).

| Item | Value |
|---|---|
| Ours | `e6cca92db2` (`feat/office-format-completeness`) |
| Theirs | `a67a38abf2` (`feat/spec-gap-implementation`, committed tip only) |
| Merge base | `f0ab67b55d` |
| Merge commit | `f592ecc1b0`, tree `c3c076acd0` |
| Pure auto-merge tree, for comparison | `b609e44091` (`git merge-tree --write-tree e6cca92db2 a67a38abf2`) |

## Contents

- **[conflicts.txt](conflicts.txt).** The 65 trial-merge conflicts (62
  content, 2 add/add, 1 modify/delete), with who resolved each and how.
- **[adapted-files.txt](adapted-files.txt).** The 74 other files whose merged
  content differs from the pure auto-merge: semantic adaptations, adapted
  tests and harness/index updates.
- **[resolution-notes.md](resolution-notes.md).** Per-file notes from the
  coordinator and the four subagents.
- **[gates/](gates/).** The 0675 integration runner on the merge commit, one
  `<gate>.log` and `results-<gate>.json` per gate, each run on its own. The
  runner is [scripts/run-integration.py](scripts/run-integration.py), with
  `CARGO_BUILD_JOBS=16`. Each log's header gives `HEAD`, the index `TREE` and
  the unstaged and untracked counts, all 0.
- **[extra/](extra/).** Checks beyond the runner, on tree `c3c076acd0`:
  - the ODF/RTF/crypto/VBA/XLDM tests;
  - the facade with ODF, RTF and encryption features;
  - `check --workspace --all-targets --keep-going`;
  - `clippy --workspace --lib`;
  - `litchi-docx` clippy per feature;
  - the harness targets after the failing lib target (bins, tests,
    doctests);
  - the two pre-existing iWork failures reproduced on the merge.

  The one untracked file some headers show is this record, still being
  written.
- **[baseline/](baseline/).** The proofs that failures predate the merge. They
  come from detached checkouts of each tip, each with its own build
  directory:
  - `tip-head`: `check-workspace.log` (iWork examples) and
    `clippy-workspace-lib.log` (facade `unit_arg`);
  - `tip-inc`: `non-iwork.log`, and `harness.log` plus `allocator.log`, which
    fail `--locked` on the tip's stale lock. Also `clippy-workspace-lib.log`;
  - on both tips: `harness-two-tests.log`, `iwa-examples.log` and
    `numbers-table-relocation.log`.

  For `tip-inc/harness-two-tests.log` only, the tip's stale
  `tools/perf-baseline/Cargo.lock` was refreshed offline and then restored.
- **[inventory/](inventory/README.md).** A name-by-name comparison of the test
  lists, with a reason for every test that is absent.
- **[scripts/](scripts/).** The runner as used (it adds the tree and dirtiness
  header), plus:
  - `run-gates.sh`, which runs one gate per directory;
  - `list-tests.sh`, which produces the test inventory;
  - `regen_coverage_v2.py`, which rebinds the CRUD coverage index v2 to the
    merged registry;
  - `check_sides.py`, which finds lines either side added that are absent
    from a merged file.

- **[doc-fresh-writer-repin/](doc-fresh-writer-repin/receipt.json).** The
  post-merge re-pin of the three DOC fresh-writer corpora:
  - the one-sample preflight report and catalog (gzipped);
  - its receipt (revision, binary hash, command);
  - the promotion script, which verifies that only the DOC corpora changed,
    and its summary of old and new identities.

- **[postmerge-gates/](postmerge-gates/results.json).** The final 16-gate
  run after the post-merge fixes and the review follow-up, on `9dda621226`
  (tree `68d48794dd`, 0 unstaged, 0 untracked): one `<gate>.log` per gate,
  `results.json`, the runner's output and the runner copy (0675's, with
  `CARGO_BUILD_JOBS=16` and the header lines above). Every gate passed.
  [postmerge-cleanup.json](postmerge-cleanup.json) lists what was removed
  afterwards.
- **[remaining-absence-probes.txt](remaining-absence-probes.txt).** Part
  lookups that may still read a decode failure as absence. The review
  follow-up found them and left them unchanged. Each needs its own review.

[cleanup.json](cleanup.json) lists what was removed after the gates ran: the
two tip worktrees, every build directory and the scratch directory. The
integration worktree is kept.

## Gate summary at `f592ecc1b0`

| Gate | Exit | Result |
|---|---:|---|
| fmt | 0 | |
| check | 0 | |
| clippy | 0 | `--lib`, `-D warnings` |
| tests | 0 | 13,221 passed, 0 failed, 78 ignored (223 s) |
| facade | 0 | 382 passed, 7 ignored |
| rustdoc | 0 | `-D warnings` |
| harness | 101 | 552 passed, 2 failed, 1 ignored ([log](gates/harness.log)); both failures pre-existing on the incoming tip |
| allocator | 0 | 5 passed |
| facade-polyglot | 0 | 104 passed |
| claims | 0 | |
| claims-structural | 0 | |
| gate-tests | 0 | |
| report | 0 | |
| coverage | 0 | |
| non-iwork | 1 | pre-existing: `litchi-xldm` is not registered |
| boundaries | 0 | |

## Lock files

- **Root `Cargo.lock`.** It is gitignored and not committed. The merge used our
  lock plus an offline, workspace-only resolution; no registry version changed.
- **`tools/perf-baseline/Cargo.lock`.** It is tracked and part of the merge
  commit. It was refreshed with `cargo update --offline --workspace`, which
  added `same-file` 1.0.6 and `winapi-util` 0.1.11, the versions in the root
  lock.
