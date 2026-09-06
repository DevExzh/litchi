# 0433 seal and replay handoff

`seal.py` is the terminal bundle step. Before changing any file it rejects
`status: running` capture, state, and command receipts. It losslessly stores
logs as deterministic gzip (`mtime=0`) and recursively includes `.data` files
under both `before/profiles/` and `after/profiles/`. It writes
`expected-checks.json` from terminal command receipts and rebuilds a complete
`SHA256SUMS` inventory over every bundle file except the inventory itself.

`replay.py --tag <lowercase-tag>` requires that sealed inventory, launches the
0433 verifier with its actual `--portable-check --require-inventory` CLI from
the directory above the bundle, and retains the terminal output as a receipt
and log only after the verifier exits. It records driver hashes and verifies
that both the bound drivers and `SHA256SUMS` stayed unchanged; it never rewrites
the inventory. The retained receipt and log therefore need a later seal to be
included in a new inventory.

No script was executed while preparing this handoff. The coordinator should
run the normal seal, then external replay, and inspect the retained receipt
before any final reseal.
