# 0722 — writer-local DOCX structural fusion

0721 removed a complete structural traversal and improved edit medians, but
read controls regressed in all pairs and NumberedList edit region peak grew
2,480 bytes. This candidate confines dual-state parsing to the mutable writer.
The existing read-only namespace scanner and public alt scanner bodies must
remain byte-identical. Moving parser state into a private writer-only boundary
must preserve separate namespace/capture/limit rules, deferred range errors,
and both ordered MCE decisions. Explicitly releasing parser scratch before
MCE is a lifetime hypothesis, not an allocator saving claim.

Before any native captures, qualify the exact source with independent legacy
oracles, public facade/publication parity, debug tracing, all-feature DOCX
checks and the performance harness tests. Source maps, full binary identities,
and accepted constraint hashes bind every lane. Temporary tracing never enters
the native binaries.

Use CPU 12, generated-medium and NumberedList edit/lifecycle, two ABBA cycles,
100 warmups and 200 native samples per corpus/phase/stage. Generated-medium
paragraph listing and the pinned filesystem media paragraph count remain
independent interleaved read controls. The filesystem control retains its 100
warmup, 200 priming, 200 measured inner processes per invocation. Capture all stages
without resampling or excluding observations.

Every paired edit p50/mean must improve at least 3%; every lifecycle/read p50/mean
must regress no more than 3%. A separate allocator ABBA cycle matches the first
four stages after native capture, with zero warmups and three samples. Request
counts, requested bytes and per-operation peak above starting live bytes must
regress no more than 3%; net live bytes must not increase. Never sum phases or
subtract unpaired process counters. Keep all tails/repeat drifts above 5% visible.
Reject and restore baseline if any correctness or frozen performance gate fails.

Edit excludes open; lifecycle includes open, edit and atomic save. Owner drop,
output readback and cleanup remain outside both clocks and allocation regions.
No RSS, leak, cold-cache, hardware-counter, throughput or scaling claim is made.
Host allocator-exhaustion scheduling is not proven equivalent; explicit typed
resource policies and refusal precedence must remain preserved.
