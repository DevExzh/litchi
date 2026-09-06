# 0436 ODT ordinary-text Work batching

The baseline ODT sequential publisher keeps the measured operation allocator
peak at 420,091 bytes for 64, 8,192 and 32,768 paragraphs, but its scalar text
encoder repeatedly charges hierarchical Work. The fresh baseline pilots
measured p50 0.175881, 13.726897 and 54.492248 ms respectively. These three-
sample pilots establish a starting point, not a formal latency claim.

The previous 0435 streaming whole-process profile has 45.74% self samples in
ExecutionContext::consume. Both current baseline executable hashes exactly
match the 0435 candidate binaries, so the raw profile is retained with its
original capture provenance in preparatory/baseline-profile. The current Cargo
baseline build reused these cached artifacts. The profile includes setup,
hashing, warmups, measured calls and the oracle; its sample fraction is not
an operation-only causal estimate.

The hypothesis is that borrowed ordinary UTF-8 spans of at most 256 bytes can
remove repeated Work charges and scratch writes while retaining the existing
per-scalar cancellation checks. Spaces, tabs, line feeds and escaped scalars
keep their existing encoding. Checked arithmetic, paragraph and content XML
ceilings, and a rejected Work span charge fall back to the scalar writer
without retaining a speculative error. The budget owner already rolls back
rejected hierarchical charges. Scalar fallback preserves the first failing
scalar and content-before-paragraph limit precedence. Cancellation propagates;
the common generated-XML fragment boundary discards a failed fragment.

Only private ODT emission code changes. Successful output bytes, sink calls,
resource totals and allocation strategy must remain identical. Cancellation
polling remains per scalar, but asynchronous cancellation can occur during a
pending uncommitted span; internal scratch progress and Work at that instant
are not promised to reproduce the former scalar timeline. A span whose Work
charge has succeeded may complete up to 256 bytes before the next cancellation
check. Caller-visible
accepted sink progress retains the existing partial-publication contract.

Validation includes differential UTF-8, whitespace and entity output; every
threshold through representative text spans for local and parent/child Work
limits; cancellation; existing ODT integration and doc tests; lint, formatting
and crate-boundary gates. The formal same-API experiment uses CPU 2, one worker,
three shapes, normal and allocator binaries, 30 samples and three warmups,
two repeats in ABBA order (24 reports, 720 samples), plus four fresh process
profiles. Exact archive/content/styles/meta/semantic/sink identities are
prerequisites. All individual results, uncertainty and >5% regression/repeat
flags remain visible. Retain the optimization only if measurements support a
practically useful improvement without changing the resource contract.
