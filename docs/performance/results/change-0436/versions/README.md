# Capture-time verifier

`verify-capture.py` is the exact outer verifier bound by every formal build
and capture receipt. Its SHA-256 is `409851684c1e68dd43e3950fc9f57797a56260afee71742effa391b460e4b838`.
The complete 24-report/720-sample/four-profile matrix passed this driver in
`checks/formal-matrix-verify.json` before the replay correction.

The current `verify.py` corrects the inherited cleanup set from five paths to
the three actual 0436 scratch directories. It validates the pinned historical
driver hash when checking capture/build bindings; measured inputs and receipt
hashes are unchanged. Its own bytes are covered by the final sealed inventory
and replay driver snapshots. When reconstructing fresh benchmark captures,
use this capture-time version as `verify.py` alongside the retained capture
and build-descriptor drivers, then use the current replay version for final
three-directory cleanup validation.
