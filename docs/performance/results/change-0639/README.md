# Evidence: change 0639, two gate-list gaps and two dead paths

Change record:
[`0639-gate-list-gaps-and-two-dead-paths.md`](../../0639-gate-list-gaps-and-two-dead-paths.md).

Disposition: retained, gate and correctness work. `performance_claim: none`.
**No library operation is timed anywhere in this packet**: every entry is a
pass/fail verdict, a transcript, a byte census or an item count. There is
therefore no A/A floor to report — no latency, allocation or RSS figure of any
litchi operation is stated in the record or here. Two wall times do appear, both
of a *test suite* and both only to justify one gate flag: the `harness-tests`
mode's 1,869.73 s of test execution when it still passed `--test-threads=1`
against 581.14 s without it, for the same 531 passing tests.  They are not a
performance result and nothing else depends on them.

## Contents

| Path | What it is |
| --- | --- |
| `demo/demo-before.txt` | `office_crud_demo` run in an empty directory at the base `c7326f680`. Exit **1**, dying at the PPTX UPDATE step with `UnsafeEdit { operation: "presentation_mut", … }` — change 0607's *Defect observed* reproduced verbatim. |
| `demo/demo-after.txt` | The same binary built from this branch, same empty-directory protocol. Exit **0**, all six output files, `Closing slide now has 4 shapes` read back from the saved package. |
| `demo/member-delta.txt` | Member-by-member sha256 census of `demo_presentation.pptx` against `demo_presentation_updated.pptx`: 43 members before and after, order identical, **42 byte-identical, 1 changed** (`ppt/slides/slide3.xml`, 1,584 → 2,154 bytes), 0 added, 0 removed, both edited strings present in the one changed part. |
| `demo/member-delta.py` | The census script, so the table can be regenerated from any pair of packages. |
| `api/public-items.txt` | The rustdoc item-page inventory of `litchi` built with `--no-default-features --features docx,pptx,xls,xlsx`, one documented public path per line, `LC_ALL=C` sorted: **8,086 pages**. The base and the branch listings are byte-identical, so one copy is retained. |
| `api/public-items.diff` | The two legs' line counts and sha256s and the empty `diff` between them. Deleting `refine_workbook_format` removed no documented item, because `sheet::workbook_types` is a private module. |
| `counts/dead-helper.txt` | The workspace-wide `git grep` for `refine_workbook_format` at the base, the module's visibility, `git diff --numstat` for the file, and the `read_to_end` inventory of `crates/litchi/src` before and after. |
| `gates/new-modes-after.txt` | The two new gate modes run end to end on this branch through `tools/non_iwork_gate.py`, with their per-target `test result` lines and exit statuses. |
| `gates/new-modes-before.txt` | The identical argv run against the untouched base tree (the base has no gate modes), so the branch's totals read as a delta and any pre-existing failure is attributed. |
| `gates/harness-tests-serialized.txt`, `gates/harness-tests-serialized.json` | The first `harness-tests` run, when the mode still passed `--test-threads=1`: the same 531 passing tests over the same 18 targets, 1,869.73 s of test execution against 581.14 s unserialized. Retained because it is the measurement behind the mode's flag choice. |
| `gates/facade-format-tests.json`, `gates/harness-tests.json` | The gate's own `--record-file` execution reports for the two new modes, the same schema every other mode writes. |
| `gates.txt` | The tail of every gate run before the commit, with exit statuses, the pre-existing warnings identified as such, and the one failing gate reproduced on the untouched before checkout. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge. |

## Provenance

Base commit `c7326f68065edf6f2198ca3cb39c38c48cf00ed9`
(`feat/office-format-completeness`); branch
`perf/0639-gate-list-gaps-and-two-dead-paths`. Host: AMD EPYC 9R45, 32 cores,
123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0, cargo 1.95.0.

All work happened in one detached `git worktree` on disk at
`/home/zhuhe/code/litchi-worktrees/0639` with its own `target/`. The shared
working copy at `/home/zhuhe/code/litchi` was never built in or modified. The
worktree and its build directory were deleted once these transcripts were
copied here; see the cleanup note in `log-sections.md`. The host carried eight
concurrent agents throughout, which affects wall time only: no figure in this
packet is a time.

**No measured binary exists**, so no binary sha256 is listed (the two rustdoc
inventories carry their sha256s in `api/public-items.diff`). The example is
built with `cargo build --locked --offline --example office_crud_demo
--features docx,pptx,xlsx,ooxml-common` in the `dev` profile; it is a
determinism-friendly demonstration — no clock, no PRNG, no network — so one run
per leg is the whole evidence. The before and after legs use the *same* build
command and differ only in the checkout.

`Cargo.lock` is gitignored in this repository (`.gitignore` line 85, `*.lock`).
A fresh worktree therefore has none, and `--locked` fails outright until the
repository's own lockfile is copied in:

```sh
cp /home/zhuhe/code/litchi/Cargo.lock <worktree>/Cargo.lock
cp /home/zhuhe/code/litchi/tools/perf-baseline/Cargo.lock <worktree>/tools/perf-baseline/Cargo.lock
```

The two new gate modes deliberately pass no `--locked`, matching every other
mode of `tools/non_iwork_gate.py`, so CI's lockfile-free checkout works.

## Reproducing the demo legs

```sh
git worktree add --detach /path/base c7326f680
cp /home/zhuhe/code/litchi/Cargo.lock /path/base/Cargo.lock
CARGO_TARGET_DIR=/path/target-base cargo build --manifest-path /path/base/Cargo.toml \
  --locked --offline --example office_crud_demo --features docx,pptx,xlsx,ooxml-common
mkdir /path/run && cd /path/run && /path/target-base/debug/examples/office_crud_demo
```

The base binary exits 1 after `Updating PowerPoint presentation...` and leaves
`demo_presentation_updated.pptx` unwritten. The branch binary, built the same
way, exits 0 and writes it.

## Reproducing the member census

```sh
python3 docs/performance/results/change-0639/demo/member-delta.py \
  demo_presentation.pptx demo_presentation_updated.pptx
```

## The one failing gate is pre-existing

`python3 tools/check_example_targets.py` exits 1 with four cross-package
duplicate example target names, all between `litchi-keynote`, `litchi-numbers`
and `litchi-pages`. It was reproduced with the identical four findings and the
identical exit status on the untouched shared before checkout at
`/home/zhuhe/code/litchi-worktrees/before-c7326f680`, so it predates this
branch, and every crate it names is an iWork crate the brief excludes. The
example this change edits, `office_crud_demo`, is not among them and its target
name is unchanged. The gate's own unit tests pass.

## What this packet does not establish

- No latency, throughput, allocation, peak-RSS, cold-cache, physical-device,
  range-source, concurrency or cross-platform result, and no claim of any kind.
- Not that the two new modes' suites are *sufficient*, only that they are now
  reachable from a gate CI executes. Neither suite gained a test here.
- Not that CI runs on the performance program's own branch. Both workflows in
  this repository trigger against `main` only; this change does not widen that.
- Not that `office_crud_demo` stays green. It is an example, nothing runs it in
  CI, and no runner was added. The evidence is two manual runs.
- Not that the member census generalizes. It is one authored three-slide deck
  edited through two transaction operations, not a corpus sweep.
