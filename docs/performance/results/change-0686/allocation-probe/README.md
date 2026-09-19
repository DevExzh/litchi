# 0686 allocation companion

Derived from the committed 0684 allocator probe. This external instrumentation
uses the system allocator and is not a production dependency. Successful
reallocations contribute their whole requested size to allocated bytes;
live/peak gauges track the size difference. Source and owner construction are
outside query gauges. Prior queries are dropped; the measured value/error and
owner remain alive through capture. Warm errors do not abort refusal scenarios.
No allocator-instrumented timing is used as native latency.

Usage: `xls0686-alloc owned|file q1|q2|q3|q8|visit FILE SHEET ROW COLUMN BUDGET`.
Build with `cargo build --manifest-path docs/performance/results/change-0686/allocation-probe/Cargo.toml --release --locked --offline`.
