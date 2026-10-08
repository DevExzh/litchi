# 0970 bounded diagnostic evidence packet

**Status: sealed bounded diagnostic packet.** This record preserves the authorized 0970
historical-baseline diagnostic capture and its custody metadata without copying
raw packets, stdout/stderr, sidecars, or compiled executables into the
repository. It contains no candidate qualification, adoption decision, causal
claim, or source change.

The packet retains the frozen 0969 baseline-only A/A analyses and read-only host
observations in `analysis/`, `host/`, and `design/`. Those reports are
exploratory and descriptive. They do not establish a host cause or a new
performance result. The current design is
[`design/diagnostic-design-v3.md`](design/diagnostic-design-v3.md), with the
corrected source-variability review at
[`design/source-variability-review-v2.md`](design/source-variability-review-v2.md). The design document itself authorizes no source edit, Cargo command, build, benchmark, capture, or qualification rerun; the separately sealed root GO v4 receipt is the authorization bound to the completed historical-baseline diagnostic capture.

The executed 0970 toolchain and plan are retained under
[`tools/`](tools/) and [`plan/plan-0970-v6.json`](plan/plan-0970-v6.json):

- `launch_baseline_v2.py` is the launcher.
- `protocol_0970_v3.py` is the packet/protocol implementation.
- `run_0970_v3.py` is the capture runner.
- `reduce_0970_v3.py` is the first reducer invoked after capture; its pre-output failure is retained for history.
- `reduce_0970_v6.py` is the corrected reducer used for the authoritative descriptive output.
- `prepare_plan_0970_v3.py` is the plan generator.

The bounded receipts record the final helper seal v6, diagnostic/combined seal
v3, parser correction v2, additive timestamp correction v1, independent
prelaunch review v3, root GO v4, quiescence v1, and terminal v1. Reducer
correction receipts v1, v2, and v3 preserve the repair history. The historical
review v2 is retained under `review/historical/`; review v3 supersedes it for
prelaunch admission. The helper seal v5 and intermediate reducer artifacts
remain external as superseded or historical artifacts.

The sealed independent raw audit is
[`review/raw-independent-audit-receipt-0970-v1.json`](review/raw-independent-audit-receipt-0970-v1.json),
with its detailed report, scalar cross-check, and replay helper beside it. Its
immutable receipt records `pass_with_tooling_pending` because the shared v3
reducer had not yet been repaired. The fresh v6 reducer replay and independent
approval now resolve that tooling finding: raw custody, terminal completion,
affinity masks, exact correctness gates/oracles, sidecar parseability, and the
v6 reduction all pass. The scalar and v6 outputs remain descriptive telemetry;
they do not qualify a candidate or establish causality.

The authorized native run completed 96 of 96 packets with zero retries. Only
six top-level capture metadata files are copied under `capture/`: the affinity
control records, capture manifest, capture stage, and quiescence record. The
complete external capture inventory is
`/home/zhuhe/litchi-performance-0970/analysis/raw-inventory-diagnostic-v3-v1.json`
(210,822 bytes, SHA-256
`d52b0eea9d7db551233317d59acf4a2ea857ee9630d3c23be83539649831d74d`). It
contains one absolute and relative path, byte count, and SHA-256 for each of
582 capture files. The 576 packet/raw/sidecar files total 95,242,613 bytes,
with raw-set digest
`ec510a06c331809d77b8c25ec3f681e20e2a4641da430f54d20884fcf6c2b7b4`; the
complete 582-file capture set totals 95,265,873 bytes, with capture-set digest
`2b0b8d30b2983e0de6ff11914c57a13317854088c39d861e0f2be2a832667346`. The
bounded repository descriptor is
[`provenance/raw-inventory-descriptor-diagnostic-v3-v1.json`](provenance/raw-inventory-descriptor-diagnostic-v3-v1.json),
and the custody summary is
[`analysis/raw-capture-summary-diagnostic-v3-v1.json`](analysis/raw-capture-summary-diagnostic-v3-v1.json).

The exact authorized runner command was:

```text
env LC_ALL=C LANG=C PYTHONDONTWRITEBYTECODE=1 /usr/bin/python3.14 /home/zhuhe/litchi-performance-0970/tools/run_0970_v3.py --plan /home/zhuhe/litchi-performance-0970/plans/plan-0970-v6.json --run --go /home/zhuhe/litchi-performance-0970/receipts/root-go-0970-v4.json --output-root /home/zhuhe/litchi-performance-0970/capture/diagnostic-v3
```

A fresh output root must be selected for any replay. The packaging work did not
start a workload or build and did not rerun the capture.

The first offline reduction attempt is preserved in
[`receipts/reducer-first-attempt-v1.json`](receipts/reducer-first-attempt-v1.json).
It returned code 2 with `KeyError: 'pair'` before writing output; it did not
rerun the native capture. The repair history is retained in
[`receipts/reducer-correction-0970-v1.json`](receipts/reducer-correction-0970-v1.json),
[`receipts/reducer-correction-0970-v2.json`](receipts/reducer-correction-0970-v2.json),
and [`receipts/reducer-correction-0970-v3.json`](receipts/reducer-correction-0970-v3.json).
The authoritative output is
[`analysis/affinity-diagnostic-v6-final.json`](analysis/affinity-diagnostic-v6-final.json),
produced by `tools/reduce_0970_v6.py` over all 96 retained packets and
independently replayed byte-identically. The approval is
[`review/reducer-v6-independent-approval-0970-v1.json`](review/reducer-v6-independent-approval-0970-v1.json).
It leaves decision and qualification null and makes no host-cause, source-owner,
optimization, or adoption claim. The v6 reducer SHA-256 is
`1cc1245b9132c80ff57172e146e5628b148f2d108ed108c4c80b9de6777e2423`; the
555,437-byte output SHA-256 is
`560e4cae2cd338fd936c57d8f3aa9ecd663aa84736ce8ea5aa431a161066c4da`.

The packet inventory is
[`packet-manifest-0970-v1.json`](packet-manifest-0970-v1.json). Cleanup completed
under [`receipts/temporary-cleanup-0970-v1.json`](receipts/temporary-cleanup-0970-v1.json):
four authorized paths were handled, 21 regular files were removed, and
1,830,912 allocated bytes were freed while the raw inventory and authoritative
output remained unchanged. Final validation confirms every copied file is below
1 MiB, no compiled executable, packet, raw capture, or sidecar is included, and
the diagnostic output contains no adoption or qualification decision. The
unrelated working-tree files outside `docs/performance/results/change-0970/`
are preserved.
