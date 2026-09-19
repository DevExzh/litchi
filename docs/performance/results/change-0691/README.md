# Change 0691 — refreshed PPTX MCE baseline

This packet separates native opened-presentation timing from allocation and
MCE identity diagnostics. It revalidates the repeated-processing opportunity
after 0653 and the intervening work. Production optimization is not yet part
of this packet; `performance_claim: none`.

Baseline production revision, source hashes and all 33 unchanged constraint
hashes are in `baseline.json`. `environment.json` describes the shared Linux
host. The corpus is owned archive bytes, with warm OS caches; source file I/O,
initial package construction and target discovery are outside phase timers.
The measured phases are capture, working clone, one shape-text edit, commit,
and apply. The enclosing total also includes timer/check overhead between
those calls. Save and reopen are excluded, so this is not a complete save
workflow measurement.

Run `prepare-control.py`, then `build.py` in the repository checkout. The
standalone probe is built with warnings denied, release optimization and two
Cargo jobs. It freezes native and allocation binaries separately. Run
`measure.py` for four independent process legs, reversing case order on odd
legs. Each process has five warmups and 100 measured fresh packages. Run
`profile.py` separately for native sampling, hardware counters and child RSS.
Adapt the absolute sibling scratch paths and CPU 12 consistently on another
machine. Build durations describe logistics, not clean build performance.

`prepare-control.py` changes the exact MCE namespace URI to an inert,
same-length spelling. Every member's before/after hash and replacement count
is recorded. Member names, timestamps and uncompressed lengths are preserved;
compressed bytes and lengths, flags and external attributes change on all
103 members (`control-metadata-check.json`). This control alters semantics and
is used only to isolate the marker-sensitive processing path. It is never a
preservation oracle for the original deck.

The allocator companion and temporary trace build are diagnostic instruments.
Their timings are excluded from native results. Pointer identities are
process-local observations, not persistent IDs or proof that foreign `Part`
implementations share their `blob()` and `blob_arc()` storage. Any future
reuse must prove identity at runtime and retain the source owner for its full
entry lifetime.

No physical cold-cache, concurrent scaling, remote-source, cross-platform,
native Office, or general PPTX performance claim is made. The broader GOAL
remains active; iWork is excluded.

`initial/` preserves the first complete capture before review strengthened
the untouched-part oracle and enforced `--locked` builds. All final measurements
were repeated with the corrected probe; do not pool the two captures.
The archived README's broad metadata-preservation wording is superseded by
the explicit control metadata differences above.
`audit.py --initial` checks the archived native bindings independently.
`trace-summary.py --verify-live-binary` additionally checks the temporary
trace executable while it exists. Its default validates retained receipts,
source/patch hashes and restoration after owned binaries are cleaned.
