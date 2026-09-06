# Validation scope and retained draft history

Root-owned Rust work ran serially with Rust 1.98.1, four Cargo jobs, incremental
compilation disabled, release builds and one test thread. `check.py` retains
command, environment, log hash, complete source-manifest hashes and test counts.
Rust files were not edited during any owned build/test/fuzz job.

The first OPC all-target check failed because an existing direct private-loader
test needed the new final `None` argument. Its exact source and failed log are
retained. An initial rustfmt command incorrectly selected edition 2021 and
refused pre-existing let-chains; edition 2024 was then used. That formatting
failure has no separate machine receipt and is disclosed here.

The candidate all-feature OPC suite passed 466 tests. Initial strict clippy
then found a needless final `let` binding and unnecessary option dereference in
the refactored loader. The pre-lint source and failure remain retained. Final
checks use the cleaned source and the bounded OPC fuzz extension; they do not
reuse the earlier candidate pass as final-source evidence.

The default harness command covers 361 tests; three additional feature-gated
binaries are tested separately with all features to cover the other 20 tests.
This is recorded as two commands, not a single 381-test all-feature run.

Eight added tests cover independent cold I/O comparisons, warm allocation
sharing, managed budget pinning, exact/one-under admission, capped short reads,
oversized cache bypass, mid-read cancellation/source change, signed/encrypted/CRC
refusals, ordinary/combined single-flight coordination, provisional rollback,
and native-input compressed member transfer. Existing transfer tests retain
publication/partial-output coverage. No new semantic PPTX caller adopts this API.

The OPC fuzz target now bounds both eager and source-backed paths to 1 MiB input,
128 members, 1 MiB/member and 4 MiB aggregate decoded ZIP bytes. It attempts cold
combined capture before ordinary eager decoding and repeats successful captures
through the warm cache. The standalone copied target uses the repository OPC
dependency, whose source identity is checked before/after every command. Sixteen
retained seeds include thirteen earlier ZIP32/ZIP64 descriptor fixtures, one
existing native POI PPTX, and two deterministic Store/Deflate OPC packages.
A 1,000-run address-sanitized, coverage-instrumented smoke is bounded evidence,
not exhaustive fuzz coverage. Its exclusive temporary workspace, generated lock,
source/seed hashes and binary identity are recorded before cleanup.

No latency, peak RSS, allocator-peak, 5% regression clearance or native/default
coverage promotion follows from this batch. Matched complete PPTX timing and
broader goal gaps remain open.
