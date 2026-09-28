# 0818 — preserve untouched DOCX relationship XML

This batch follows 0817's failed independent admission. The previous turn made
progress by retaining exact outputs and separating DOCX ordering loss from
generated-manifest and XLSX calculation-closure reader assumptions. Base is
`7cdaba587b`; the original 0817 packet remains immutable.

The intended correction retains the existing main document Part and its source
relationship provenance during ordinary publication. OPC still decides whether
the final relationship binding matches the admitted source; genuinely changed
graphs must use validated serialization. No global authoring API contract is
relaxed. Independent admission continues to require exact untouched DOCX
relationship bytes and order.

Root owns Cargo, executable, and workload execution. Agents implement and
statically review the bounded source change and independent artifact oracle.
The first regression run must fail on unchanged production before applying the
fix. Fresh quality checks and a fresh six-case/thirty-output artifact export
must pass before claiming preservation repair. No latency improvement is
inferred from eliminating a serialization path; performance timing remains a
separate following measurement.

The 35 previously read normative inputs retain their exact hashes. Three
unrelated workspace files remain excluded. The broader OLE2/OOXML goal stays
active; iWork is excluded.

All six quality gates passed: 1,938 tests passed and 32 were ignored. Fresh
independent admission passes six corpora and thirty outputs. Untouched ZIP
metadata, compressed payloads, member order, and archive comments also pass.
The original regression's failure and the old-output auditor preflight remain
retained. Root removed both owned directories (8,894 files / 21,390,560,337 logical bytes)
after verifying the exporter. Offline replay after cleanup passes with
`python3 -B docs/performance/results/change-0818/validate.py --final`.
