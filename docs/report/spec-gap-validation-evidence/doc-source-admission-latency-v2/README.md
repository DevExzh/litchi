# DOC source-admission latency evidence

The corrected release comparison measures the frozen c859 control against the
eight-file source-admission overlay recorded by b2e2ff8e1. It does not compare
the complete trees of those commits. All 769 resolved source files and eight
fixtures per side are hashed; the archive includes the patch, locked harnesses,
raw results, build/environment provenance, and a portable reproduction wrapper.

The large-input candidate/control median ratios were:

| Operation | Corrected v2 | Root reproduction |
|---|---:|---:|
| Embedded snapshot open | 0.816 | 0.821 |
| Embedded open plus no-op finish | 1.227 | 1.232 |
| Embedded replacement plus outer commit | 0.971 | 0.959 |
| Common unique open plus finish | 0.800 | 0.807 |

The open-plus-finish slowdown is a compound observation including finish output
allocation. It does not identify a finish-only defect. Setup, source cloning,
common targets/limits construction, result equality checks and drops are outside
the timers. Each run has six alternating paired rounds, five warmups and twenty
samples per invocation: 96 invocations and 1,920 timed samples. Small-input
changes remain near the timing/noise scale. This is one host, one pinned core,
two synthetic sizes, one release profile and Rust 1.95.0.

Extract `evidence.tar.gz` and follow `publication/README.md`. The recommended
`publication/reproduce.py` verifies source/fixture hashes, rebuilds both locked
release binaries, runs fresh correctness checks, and times only those binaries.
It records flags, commands and binary hashes and uses disposable output paths.
Root executed that complete workflow: both fresh 34-check correctness gates
passed, timings completed, and the publication remained unchanged. The separate
`root-reproduction/` directory retains the full reproduction evidence.

`receipt.json` hashes every archived member. The historical source review came
from allocation-only work; its provenance note distinguishes that review from
this release latency evidence. Historical auxiliary gate-log names and source
receipt hashes refer to earlier evidence; the bundled source manifests and patch
are authoritative for reproduction. No native-producer acceptance, throughput,
peak-memory or broad performance improvement is claimed.
