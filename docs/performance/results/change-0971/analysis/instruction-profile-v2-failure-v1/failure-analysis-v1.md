# 0971 instruction-profile-v2 pilot forensic analysis

Created at `2026-10-08T03:31:31.706540+00:00`. This is a bounded forensic record for the already-terminal pilot. It
preserves the capture failure and does not qualify a CPU profile or make a latency or speedup
claim.

## Result

The pilot gate failed because the exact native marker
`doc_retention_probe_0942_cfb_profile::profiled_measured_case` was recovered in zero decoded
samples. The sealed packet reports `1387` instruction samples, total period
`6,935,000,000`, `lost_events=0`, `906` zero-frame/truncated records, and `618` unresolved
frames. Its marker sample count and marker period share are both zero. The capture manifest
therefore records `pilot_gate_failed`, requests no retry, and contains only child-01.

No native function is assigned to the measured edit region. The packet's ordinary target-DSO
frames are setup/preflight or other unresolved context; they are not exact marker ancestors.
The profile is unsuitable for phase attribution, Amdahl weighting, or any speedup inference.

## Independent offline checks

The sealed no-inline and inline perf-script decodes contain no textual occurrence of
`profiled_measured_case`, `run_case`, or `measure_one`. The target ELF is nevertheless present
in the perf build-ID list with BuildID
`e5c457384eeebd057932c444efda77618f137399`; `perf report` resolves ordinary target symbols
such as `litchi_doc::sprm::parse_sprms` and `TextExtractor::new`. `nm -n -C` resolves the
marker at `0xec4f0`, `run_case` at `0xf15f0`, and `measure_one` at `0xe6ca0`. Disassembly
shows the marker's direct call to `run_case` at marker-relative return address `0xec50b`.
Thus this failure is not explained by a missing target DSO, missing BuildID, or wholly broken
symbolization.

The raw `perf script -D` diagnostic examined the captured event records without launching a
process. It saw `1377` sample records matching its raw SAMPLE filter, including `891` target
DSO leaf-IP records with period `4,455,000,000`. Those `891` records are descriptive only and
must not be treated as a phase or owner population. In all of them, the captured raw 16 KiB
stack words contained zero marker-range words, zero marker direct-call return words, zero
`measure_one` caller-return words, and zero `run_case` range words. This corroborates that the
marker ancestry is absent from the captured raw callchain content. The `1377` raw-filter count
is not used as a replacement for the packet's `1387` sample count.

## Interpretation and limit

The evidence establishes a callchain/ancestry capture failure for this pilot: samples exist,
perf reports no loss, target code is mapped and symbolizable, but the marker and its callers
are absent from both decoded callchains and the raw stack words inspected. The offline evidence
cannot distinguish the underlying mechanism, such as stack-depth exhaustion versus a DWARF
unwind boundary or wrapper-specific unwind shape. The launch used the recorded Python helper
and then exec'd the target; that is provenance context, not a demonstrated causal explanation.

The result remains `unsupported` for native phase attribution. Preserve the sealed capture and
all derived diagnostics. Do not combine the `891` descriptive leaf records with wall-time or
phase metrics, and do not infer performance impact from this failed pilot. No retry or new
Office workload was launched in this analysis.

## Artifact locations

- Sealed packet and capture: `/home/zhuhe/litchi-performance-0971/capture/instruction-profile-v2/`
- Offline diagnostics: `/home/zhuhe/litchi-performance-0971/analysis/instruction-profile-v2-failure-v1/`
- Exact offline command log: `commands-v1.txt`
- Raw diagnostic generator: `raw-stack-scan-v1.py`
- Raw diagnostic result: `raw-stack-scan-v1.json`
- This analysis: `failure-analysis-v1.md`

The accompanying `receipt-v1.json` binds the sealed packet, binary, perf tool, decoded files,
and derived diagnostics by SHA-256. All large raw files remain outside the repository.
