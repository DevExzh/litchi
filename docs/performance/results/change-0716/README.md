# 0716 DOCX baseline stationarity diagnostic

This packet follows the rejected 0715 experiment with one newly built binary
from the restored, unchanged baseline. It does not rerun that candidate or
alter its rejection. The production and Rust benchmark source census remains
identical to 0713/0714 and the restored 0715 source.

The frozen matrix has eight blocks, two corpora and two warmup counts (10 and
100): 32 fresh native processes, each with 200 measured counting-publication
samples on CPU 12. Four cyclic orderings repeat twice, placing each treatment
twice at each position. All children and samples remain retained.

`analysis.json` reconstructs execution order from the raw sample permutation,
reports four sequential 50-sample means and two half means, and retains every
within-child and between-child spread above 5%. These are descriptive flags,
not a formal stationarity test or an optimization acceptance decision. Within
each block, warmup comparisons retain both corpus pairs without exclusions.

The external parent records `wait4` resource counters and before/after Linux
context snapshots. They cover the whole process, including setup, warmup,
untimed opening, verification and report writing. CPU and pressure snapshots
also include unrelated activity. They cannot attribute a serialization change
to scheduling, CPU frequency, hardware, allocator state or hash seeds.
Rust hash seeds and allocator/ASLR layout remain uncontrolled; Perl seed
variables do not control Rust hashing. Only the listed environment fields are
constrained or recorded. The generic report fields about fresh filesystem
children do not describe this ordinary-save route: its samples run in process.

No production or Rust harness change requires new format correctness runs.
`reused-verification.json` binds exact-source prior verification: 4,995 tests,
92 passed/46 ignored doctests, scoped Clippy and six repository evidence gates.
The seven preexisting PPTX/XLSB test Clippy exceptions remain disclosed.

Raw child artifacts, receipts, source/build/fixture identities, environment,
frozen plan and scripts, negative checks, review and cleanup records are covered
by the terminal artifact manifest. Replay requires no retained executable.

```sh
python3 -B docs/performance/results/change-0716/audit.py
python3 -B docs/performance/results/change-0716/artifact-seal.py --check
```
