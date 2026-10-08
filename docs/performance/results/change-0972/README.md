# 0972 offline decoder diagnosis packet

Status: **sealed final diagnostic packet**. This record contains the final v6 ABI-correct offline replay, PF_X-derived raw-stack correction, author seal, independent review, binding-resolution receipt, and cleanup receipt. It is diagnostic evidence from one retained 0971 child profile; it is not a new native measurement or a production change.

The work did not launch the retained executable, run an Office workload, build a target, mutate the sealed `perf.data`, or edit production source. The corrected replay reads the unchanged input and recovers 78 exact marker-containing chains. Stock and no-op retain 906 empty decoded chains and zero marker chains; corrected v6 has zero empty chains, 11,154 frame lines, 10,543 resolved frames, and 611 unresolved lines. All three modes have 1,387 sample headers, total period 6,935,000,000, and zero lost events.

The final replay supports an offline map-metadata/decoder interaction under the exact six-argument `dwfl_report_elf` ABI and comparator-aware v6 interposer. It does not prove the upstream perf/libdw root cause, a stack-depth cause, or what a new live capture would do. The 78 recovered marker chains remain below the original 100-sample gate, so there is no qualified cohort, phase or owner fraction, timing result, speedup claim, or adoption decision. The raw marker-period share in the decode summary is descriptive only.

The PF_X-derived raw scan corrects the historical one-page-high base and truncated executable bound. It reports 908 target-leaf records and potential address words (53 marker-range, 39 marker-return, 17 `measure_one`-return, and 54 `run_case`-range samples). These are raw word matches, not live ancestry or phase evidence; the v6 decoded chains are the authoritative ancestry result.

Final bounded files include:

- `analysis/diagnosis-v3.md` and `receipts/diagnosis-v3-receipt.json`;
- `analysis/dwfl-corrected-decode-v6.json`, `analysis/corrected-raw-stack-scan-0972-v2.json`, and the scanner source;
- v6 helper source/map, author/replay receipts, independent review, root binding resolution, ADR revalidation, and cleanup receipt;
- exact replay and scan commands, input evidence bindings, and a historical correction receipt for the invalid v1 shim.

Raw perf data, derived perf copies, verbose logs, decoded JSONL, retained executable, and compiled shim remain outside the repository. `provenance/external-artifact-descriptor-final-v1.json` binds those artifacts by absolute path, byte count, and SHA-256 and records the superseded v1-v5 helper lineage. The final file list and hashes are in `packet-manifest-0972-v1.json`.
