# 0450: Return decoded bytes and a verified token from one ZIP capture

Add `IndexedArchive::read_entry_precompressed_and_decoded_with_progress`.
It captures a Store/Deflate payload once, decodes it once, and returns both logical
bytes and the existing private-field verified compressed token. The prior API
still compares against caller-supplied decoded bytes through the shared verifier.

Ten deterministic cases use independently indexed identical sources. At roughly
1 MiB, source bytes fall from 2,097,297 to 1,048,654 for Store and 2,098,011 to
1,049,011 for Deflate. Calls fall 40→19 and 58→21 respectively. Exact decoded and
compressed bytes match; publication through the preservation writer adds no
Deflate recompression and preserves the existing member. See
[all ten cases](../results/change-0450/measurements.md).

These are I/O correctness assertions, not latency samples. Both complete compressed
and decoded payloads remain resident; no allocation-peak reduction is established.
The new method is not yet adopted by OPC/PPTX. Its integration must bind source
lineage/version, semantic validation and combined memory reservations through
cache, plan and publication lifetime. Matched end-to-end timing remains required.

Final validation passes 455 ZIP tests, 436 OPC tests, 59 targeted PPTX tests, strict
ZIP/fuzz lint, feature/workspace checks, warning-denied rustdoc, boundaries and
formatting. Five new tests cover output, short reads, cancellation, CRC/size/end
checks, ZIP64, transport and allocation refusal. The updated bounded fuzz target
passes 1,000 ASan/sancov runs. Initial compile and cache-state comparison drafts
remain disclosed in [validation notes](../results/change-0450/validation-notes.md).

Replay the sealed evidence without builds or the original fuzz workspace:
`python3 -B docs/performance/results/change-0450/verify.py --sealed --cleanup`.
The full non-iWork goal remains active; this primitive completes only the ZIP
prerequisite for the measured source-transfer opportunity.
