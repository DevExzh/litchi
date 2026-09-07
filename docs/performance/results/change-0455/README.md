# Change 0455: unchanged ZIP transfer experiment

The root `runs/` directory retains six preliminary baseline lanes captured
before adding sink-summary output to the existing lifecycle journal. Their
original protocol, reports, command receipts, executable identities and source
manifests are unchanged. They are not mixed into the final timing comparison.

`formal/` holds the matched experiment with sink summaries captured after the
publication timer and allocator region close. The production ZIP accounting
API is checked directly for raw accepted bytes in preservation tests; the PPTX
lifecycle does not expose that counter and does not claim API-level attribution.
The ordinary and allocation-instrumented populations remain separate.

The retained candidate doubles the private ZIP preservation buffer to 64 KiB.
The [formal 720-sample matrix](formal/measurements.md) separates 480 ordinary
samples from 240 allocator samples. Media publication makes 256 fewer range
reads and sink writes with identical output and transferred bytes. Simulated
range API medians improve 4.775%/4.449%, with 32 KiB additional fixed stack.

The [240-sample confirmation](formal/confirmation/measurements.md), original
[hardware counters](formal/profile-summary.json), successful local instruction
profiles in `formal/stack-profiles-local/`, and
[fixed allocator-policy diagnosis](formal/diagnostic-summary.md) retain process
variability and the original adverse CPU observation. They do not establish a
CPU improvement. The failed initial stack-report attempt remains unchanged.
See [review](review.md) for disposition and limits.

`formal/verify.py --precleanup` verifies retained binaries, temporary fuzz inputs,
native outputs and raw profiles before cleanup. After `cleanup.py` inventories
and removes only `/tmp/litchi-goal-0455`, `formal/verify.py` replays portable
artifacts without those files. `precleanup.json`, `cleanup.json` and
`portable-verification.json` record actual outcomes. `SHA256SUMS` seals the final
bundle except itself. Shared build targets are retained. The full non-iWork goal
remains open; no coverage taxonomy row is promoted by this optimization.

Final precleanup and separate-copy portable verification pass. Cleanup removed
937 owned temporary files (1,032,052,420 bytes); shared targets were retained.
Run `python3 -B formal/verify.py --portable` from this bundle to additionally
validate the recorded portable pass, then `sha256sum -c SHA256SUMS` to check the
seal. No production sources changed after the measured candidate build.
