# Review decision

Retain direct framing for verified Store/Deflate and shared Store payloads.
The formal allocation counters reproduce in all 120 media allocator observations:
publication requests 4,243,083 bytes instead of 21,057,867 and peaks 593,892 bytes
above entry instead of 17,406,388. Output identity, source/destination read work,
sink work and managed owner budgets are unchanged. Plain framing metadata costs
3,360 more allocated bytes and 1,720 more peak bytes, explicitly disclosed.

Independent production review found no material correctness issue in ownership,
trusted-token boundaries, framing reuse, ZIP64 layout preflight or partial sink
accounting. The initial reborrow compile failure and corrected source remain
available. Release tests and ASAN fuzz pass; the native self-pair produces the
prior exact golden output after correcting one evidence transcription error.
No independent native-pair or native-application coverage is promoted.

The ordinary bytes/media latency regression repeats: API +15.411%/+16.347% and
publication +32.611%/+32.769%, with substantially more candidate minor faults.
All 12 adverse API/publication percentile flags are retained. Separate perf
processes show matched low faults and candidate API medians around 24.25–24.35 ms,
below baseline 25.08–25.09 ms. Fixed-policy diagnostics improve API medians
2.626–10.266% across all four pairs. Together they demonstrate policy-sensitive
results, not a universal speedup or an exact historical causal reconstruction.
The memory saving is the acceptance reason; the original regression is a
material limitation, not a discarded outlier.

The additional range/media destination-open p99 +10.210% flag occurs only in R1;
its read work and code path are unchanged, and it does not repeat. The remaining
two flags are range/media open-tail improvements. With only 30 observations per
lane these are descriptive process tails, not population-tail guarantees. RSS
comparisons do not exceed 5%. Whole-process hardware counters and instruction
profiles are separate from ordinary timings and operation heap counters.

Evidence review requested exact source/gate/fixture identities, immutable driver
bindings, and transient containment plus per-file cleanup membership. The final
verifier implements those checks and rederives measurements and diagnostics.
The accidental native expected-SHA substitution is the only copied hash change
found by an additional read-only comparison; the corrected expected file is
byte-identical to 0455. Failed originals remain immutable.
