# Captured read-bound verifier correction

The revised capture completed 840 baseline and 4,080 candidate rows after its
136-case exact-value preflight passed. The captured standalone
`performance/verify.py` retained an older assumption that every metadata case
except scalar SHEET/ABS reads zero cells. That assumption does not cover the
six newly added computed ROW/COLUMN IF, IFERROR, and IFNA cases: each evaluates
two selected child reference cells before the metadata consumer rejects the
computed value array. Direct metadata operations still read zero cells.

The pre-capture case matrix already declares exactly two reads for these six
cases; the Rust preflight and independent semantic oracle agree. The observed
audit failure was `reads 2 outside expected bounds (0, 0)`, not a changed
measurement, provider retry, or runtime admission failure.

The independent root audit checks the captured per-case contract and exact
read totals while retaining the original captured verifier, input hashes,
preflight results, and raw stdout/time receipts unchanged. Use the root
`verify.py` for final retained-evidence acceptance. No new timing capture is
needed to correct this analysis-time expectation.
