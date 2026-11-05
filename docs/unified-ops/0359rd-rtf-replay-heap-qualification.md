# 0359rd — RTF replay heap attribution and rejected qualification

- Date: 2026-10-22.
- Status: native qualification open; heap attribution is diagnostic and the
  candidate remains unaccepted.
- Candidate source commit: `68f976eb00`.
- Probe source: `crates/litchi-rtf/examples/refusal_contract_bench.rs`, SHA-256
  `9d9e91c639e47ccc94f154464f6f97a0dc0002a3b2fb37251f61b17bf71cdb22`.

This record compares the original 0359b pre-API binary with the fixed empty
underline refusal candidate. It also compares the current pre-68 binary with
the candidate to locate the change. The heap captures use the same external
executable path and the same eight-route probe. All captures returned valid
heaptrack summaries, but the host had active compiler processes throughout;
they are retained as rejected qualification receipts. Allocation counts are
deterministic across both A-B-B-A pairs and are useful for attribution. No wall
time result is accepted here.

## Binaries and method

| Label | Binary SHA-256 | Role |
| --- | --- | --- |
| original | `4693bd12232d92a61ef2f437de2d623e345a106288d5d0a5c5638cf27677420a` | pre-API 0359b baseline |
| current | `b9fccd840476aac1beaa35cb5e14d5410f6272493dab33e361d83fb4d9aed8aa` | current diagnostic before 68 |
| candidate | `126db1632fe150eae900782a98906933dec9aa3945542a0e1a99d22ae4758157` | 68 fast-path candidate |

The original-to-candidate heap order was original, candidate, candidate,
original. The diagnostic order was current, candidate, candidate, current.
Each run used the fixed path under the external qualification directory,
`taskset -c 7`, a systemd user scope with `MemoryMax=20G`, and the existing
34-sample probe (three warmups and the median of 31 retained samples). The
heaptrack wrapper SHA-256 is
`72423f5068dcfd8f02d395668fc8d08c1e3b7e98861c5785eb741e3e5c5a27ec`.

The native runner was configured for five short A-B-B-A blocks, with each
run using the same path, `argv[0]`, CPU, environment and memory cap. Its first
baseline run never reached a clean window within 180 seconds: the load began
at 118.60/121.25/126.42 and ended at 140.82/126.07/127.04, with cargo,
rustc and clippy processes present. The receipt is therefore rejected and no
native timing median or ratio is reported.

## Heap result

| Route set | Allocations | Temporary | Peak heap | Leaked |
| --- | ---: | ---: | ---: | ---: |
| original, run 1 | 177,971 | 1,258 | 3.57 MiB | 544 B |
| original, run 2 | 177,971 | 1,258 | 3.57 MiB | 544 B |
| candidate, run 1 | 188,171 | 1,254 | 3.57 MiB | 544 B |
| candidate, run 2 | 188,171 | 1,254 | 3.57 MiB | 544 B |
| current, run 1 | 188,171 | 1,254 | 3.57 MiB | 544 B |
| current, run 2 | 188,171 | 1,254 | 3.57 MiB | 544 B |

The candidate has 10,200 more allocation calls than the original, while the
current and candidate are identical. This places the regression before the
68 empty-boundary fast path. Heaptrack's candidate-only allocation edge is
`litchi_rtf::api::facade::Package::apply`; the current implementation builds
the replay `Commit` with `self.changes.to_vec()`. The required `Commit` return
and its patch/diagnostics triad explain why this work was added, but the
allocation count is still an open performance regression and requires a
source-level repair. Peak heap, leak total and all eight route dispositions
remain unchanged in this experiment.

The empty underline refusal still reports the typed
`RefusalKind::UnsupportedContent` with its fixed content-free reason. The
protected-document refusal still reports `RefusalKind::Protected`, and the
successful parse, edit, publication, replay and tail-append routes preserve
their existing outputs. The heap experiment does not relax the API, error or
redaction contract.

## Qualification boundary

The original-to-candidate allocation delta is deterministic but was measured
under host contamination, so it is not a native qualification result. The
current-to-candidate equality is diagnostic only. The owner must rerun the
five-block native A-B-B-A sequence and paired heaptrack captures during a
clean window, then verify the repaired replay path against the original
baseline, current binary and candidate with output, error disposition and
allocation checks.

External receipts remain under
`(deleted-coord-artifact)`:

- `heap-qual-20261022/` — original/candidate A-B-B-A captures;
- `heap-diag-current-20261022/` — current/candidate A-B-B-A captures;
- `native-qual-20261022/block-01-01-baseline.json` — rejected clean-window
  receipt.

The compressed heap captures are all below 1 MiB and remain outside the
repository. The small JSON receipt in `results/0359rd/qualification-summary.json`
records their hashes and the rejection boundary.
