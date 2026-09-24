# Calculation Chain checked publication

The Package and Workbook facades check the exact source and changed-package signature policy before cloning. A private publication method applies to that detached candidate, validates its resulting state, and avoids a second source capture and package clone. Publication remains atomic.

The isolated Rust 1.95.0 gate passed 780 unit/integration tests and nine doctests, strict all-target/all-feature Clippy, formatting, and whitespace checks. One native integration test and 11 existing doctests remain ignored by this gate.

`allocation-ab.tar.gz` retains both allocation harnesses, source manifests, raw rows, summaries, correctness/native inventories, and the original receipt. The measurement compares frozen commit c40 with the same two runtime files used here; the broader 9aa validation baseline is separate. For the large synthetic remove-apply workload, allocation decreased from 3,996,014 to 2,884,209 bytes and from 48,997 to 35,096 calls. Plan-and-commit counts were unchanged. These are allocation-only results from 900 rows across two profiles and 15 operations; they establish no latency, throughput, RSS, or syscall improvement.

The outer receipt binds the reviewed files, validation logs, and every archived evidence file. Original external evidence trees were preserved.

The private publication method trusts the preflighted caller and may modify its detached candidate before a failure. Its only production callers discard failed candidates and adopt them only after validation. No-op facade paths still clone; the measured optimization concerns changed patch application.
