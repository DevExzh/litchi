# Evidence: change 0619, attribution of the standing XLS lifecycle assertion

Change record:
[`0619-harness-xls-lifecycle-assertion.md`](../../0619-harness-xls-lifecycle-assertion.md).

Disposition: retained, harness-only correction. `performance_claim: none`.
**No file under `crates/` was modified.** Nothing in this packet is a paired
timing measurement; every number is a deterministic count or a transcript.

## Contents

| Path | What it is |
| --- | --- |
| `bisect/<commit>.txt` | The transcript of `cargo test --release --locked ... xls_source_backed_lifecycle_selectors_are_matched_and_local` on one detached checkout, filtered to the running/panic/result/exit lines. Eight legs: `6b13261e5` (last good), `c1d2caf85` (first bad, change 0565), `2391a3462`, `93a610ded`, `08d968f8e`, `c1503db2b` (change 0595), `f8cf7d2a1`, `1e4198321` (head). |
| `trace/source-backed-open-read-trace.txt` | Every `InstrumentedSource::read_at` of one `XlsSourceBackedOpen` at `1e4198321` — offset, requested length, satisfied count, and the overlap in bytes against each of `cfb_structural`, `workbook_global`, `selected_worksheet`, `unselected_worksheets` and `opaque_payload` — preceded by the classification range census and followed by the snapshot totals. |
| `trace-probe.patch` | The throwaway probe that produced the trace: an `AtomicBool`-gated trace line in `InstrumentedSource::read_at` and one `#[test]` that builds the corpus and layout, opens one source-backed workbook and dumps the snapshot. Applied to `tools/perf-baseline/src/lib.rs` at `1e4198321`, reverted before the correction was written, and **not** part of the committed change. |
| `gates.txt` | The tail of each of the four gates, with its exit status. |
| `decision.json` | The `litchi-perf-change-decision` record. |
| `log-sections.md` | The four log paragraphs for the coordinator to merge. |

## Provenance

Base commit `1e41983213dc378c13774ed7038c51faf231977f`
(`feat/office-format-completeness`); branch
`perf/0619-harness-xls-lifecycle-assertion`. Host: AMD EPYC 9R45, 32 cores,
123 GiB, Linux 7.0.0-1012-aws; rustc 1.95.0.

Each bisect leg was a separate detached `git worktree` on disk with its own
external `CARGO_TARGET_DIR` under `/home/zhuhe/code/litchi-worktrees/targets/`,
built `--release --locked`, so no two legs shared a build. The worktrees and
target directories were deleted once their transcripts were copied here; see
the cleanup note in `log-sections.md`. The host carried eight concurrent agents
throughout, which affects wall time only: every figure in this packet is a
deterministic count or a pass/fail, and no timing is reported.

`tools/perf-baseline` is a separate Cargo project with path dependencies on the
workspace crates, which is why its `Cargo.toml` is named explicitly in every
command and why a workspace-level `cargo test` never covered this test.

## Reproducing the decisive pair

```sh
git worktree add --detach /path/good 6b13261e5
git worktree add --detach /path/bad  c1d2caf85
for leg in good bad; do
  CARGO_TARGET_DIR=/path/target-$leg cargo test \
    --manifest-path /path/$leg/tools/perf-baseline/Cargo.toml --release --locked \
    xls_source_backed_lifecycle_selectors_are_matched_and_local
done
```

`good` passes; `bad` fails with `left: [false] right: [true]`.

## What this packet does not establish

- No latency, allocation, peak-RSS, cold-cache, physical-device, range-source or
  cross-platform result, and no claim of any kind.
- Not that nothing else changed at the bisect commits: each leg ran one test.
- Not that the 93-byte over-read is the worst case. That breadth belongs to
  change 0565's 104-fixture survey, which measured at most 3,949 bytes.
