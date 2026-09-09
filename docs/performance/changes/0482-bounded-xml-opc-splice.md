# 0482: bounded XML auditing and decoded OPC insertion

The existing DOCX append lifecycle retains complete source/candidate XML and a
paragraph index. The 0481 allocation evidence and window contract identify a
streaming XML publication check and decoded OPC insertion path as prerequisites
for removing that ownership. This batch implements those shared primitives.
Public DOCX append integration remains the next workstream.

`xml-minifier` now audits a caller-supplied `BufRead` with a finite token and
parser-state envelope. Admission happens before proportional token growth.
The existing slice audit APIs keep their behavior; streaming errors follow
source order where a complete-input check cannot be reproduced in one pass.
Successful reports and authored lexical rules are checked differentially.

OPC prepares an insertion plan from authenticated source, fragment and candidate
bytes, independently audits XML source/candidate streams, then performs two
fresh decoded passes through ZIP's measure/emit replay. It retains the source
handle, proof scalars and a bounded fragment. Untouched members retain their
raw ZIP representation. Exact no-ops copy the original archive, and inverse
publication authenticates the current candidate before copying the retained
original. Resource limits, source changes, cancellation and partial sink errors
remain explicit. ZIP owns the scalar bounds for its preservation, decoder and
replay storage, including private layout sizes and locked backend envelopes.

The [evidence bundle](../results/change-0482/README.md) compares materialized
and streaming XML audit routes from the same executable and deterministic
generator. Normal and allocator builds are separate. This is a measured
enabler comparison, not a before/after DOCX transaction result. The batch adds
no parallel execution, ambient storage or source provider.

The materialized route calls `read_to_end` on an empty `Vec`. Its allocator
observations include geometric capacity growth and EOF probing, while the
logical materialization count records input length. These costs differ from
an exact-sized allocation based on ZIP metadata. The comparison therefore
cannot be projected directly onto the existing DOCX lifecycle. Streaming
retains a finite parser window, but the OPC plan still owns its explicitly
bounded fragment and package-level preservation metadata.

Changed XML source and candidate streams must satisfy the configured authored
compactness policy, not merely XML well-formedness. Exact no-ops can preserve
malformed or signed sources byte for byte. Supporting additional existing
source lexical forms in public DOCX append requires format-owned preservation
proofs and its own compatibility evidence.

The accepted 24-process matrix contains 720 successful samples. Normal mean
times are shown below; R1/R2 are separate process repeats.

| Input | Materialized mean ms, R1 / R2 | Streaming mean ms, R1 / R2 | Streaming change, R1 / R2 |
| --- | ---: | ---: | ---: |
| 64 KiB | 0.089028 / 0.088942 | 0.088799 / 0.089081 | −0.258% / +0.156% |
| 8 MiB | 10.996315 / 10.965352 | 11.390709 / 11.414038 | +3.587% / +4.092% |
| 128 MiB | 204.054625 / 203.475179 | 181.998491 / 181.923757 | −10.809% / −10.592% |

Every streaming allocator sample has a 65,587-byte operation heap increment,
seven allocation calls, no reallocation and zero net live bytes at exit.
Materialized peaks are 147,504, 16,793,648 and 268,451,888 bytes respectively.
The largest-input normal process RSS observations are 130.15–130.18 MiB for
materialized and 1.90–1.93 MiB for streaming; these include setup and teardown.
The [complete tables](../results/change-0482/measurements.md) retain intervals,
tails, allocated bytes, corpus hashes and every individual result.

Adverse review flags remain visible: allocator 64 KiB streaming R1 mean is
11.326% higher, normal 64 KiB materialized repeat RSS rises 10.108%, and
allocator 128 MiB streaming repeat RSS rises 12.757%. The small allocator
R1 contains a 372,081 ns sample, which was retained. No normal latency repeat
change crosses 5%. The 8 MiB latency increases are also reported despite
falling below that review threshold.

The accepted source has 24 required green gates, including 1,128 passing
Rust tests across XML, OPC, ZIP and the harness, with four existing ignored
tests; warning-denied Clippy/rustdoc, workspace, boundary and strict registry
checks; five evidence tests; and 10,000 ASan/libFuzzer iterations. Live evidence
verification checks all captures, pilots, accepted builds and source identities.

The [follow-up](../results/change-0482/next-work.md) preserves the full scope:
bounded DOCX scanning and publication, replayable large paragraph production,
durable patches, source/sink variants and end-to-end scaling evidence. The
full non-iWork goal remains open.
