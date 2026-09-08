# Retained diagnostics

- The initial Heaptrack export glob expected gzip, while both successful captures produced Zstandard. Exporter support was fixed; no capture rerun.
- The first Python test run caught fractional allocation counts being parsed as an integer prefix. Whole-line integer matching fixes this; initial failure and fixed runs remain.
- Initial full-guard analysis rejected empty optional source vectors. The established0467 comparison-copy treatment is restored with explicit equal-path audit; raw measured values and policy are unchanged.
- Initial live chronology validation wrongly required the candidate build before control A1. It now requires completion before candidate B1 and no overlap with any capture. A regression test covers both valid and invalid intervals.
- The targeted seven-row guard is rejected by strict ABBA because the realized CFB max_in_flight_reads vector varies. The complete error and all raw observations are retained in review-summary.json; no fallback accepted summary is generated.

- Two raw Heaptrack stderr lines retain original trailing spaces to preserve capture hashes. The scoped diff check passes for all other staged files.
