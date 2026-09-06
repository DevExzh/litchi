# Replay checker corrections

Both recordings completed and passed the unchanged report verifier before the
new bundle replay checker was written. Their raw data, decoded stacks, reports,
protocol, capture driver, and receipts remain unchanged.

The first full portable replay failed because `verify.py` incorrectly required
the historical 0429 build's capture path to equal the new 0430 capture path.
The build records the original 0429 temporary executable; the 0430 protocol
records a separate copy with the same executable hash. The corrected checker
binds each path to its own experiment. The exact failed checker is
`01-verify.py`; its hash is retained in
`../portable-before-cleanup-export.json`, with the failed command receipt and
lossless log under `../checks/portable-before-cleanup.*`.

The second replay failed because the checker required `cycles:u` literally in
the top-level scope prose. The event is already fixed in the summary's protocol
and per-profile scope, while the prose states the warmup and sampled-count
boundary. The correction keeps the protocol event check and requires the
sample/warmup wording without duplicating the event spelling. The exact failed
checker is `02-verify.py`, bound by
`../portable-before-cleanup-v2-export.json` and its corresponding receipt/log.

The third full replay passed, including both report validators, stack
derivation, a copied bundle and mutation rejection. These corrections affect
only the newly written replay checker. They do not relax source, binary,
capture, numeric, output, stack, or hash checks and do not change any captured
observations. The failed commands remain in `expected-checks.json` with their
actual failed status.
