# 0704 retention observer

This is the candidate-only observation lane for the bounded opened-PresentationML slide MCE projection. It has a separate Cargo workspace under `retention-probe/`; it never changes the frozen `probe0704` native source or its baseline/candidate binaries.

The observer makes no timing or allocator claim. It records the public retained charge at capture, snapshot clone, transaction creation and edits, commit, and publication. It also exercises value-preserving release on one snapshot clone, one transaction, and a separate commit while another snapshot reference remains alive. Every revision, semantic digest, patch-change flag, and edited text check is obtained through the public PPTX API.

The real fixture is `test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx`. Its fixed zero-based targets are slide `1`, shape `1`, and slide `2`, shape `0`. The generated fixture is the bounded `generated:12x8` recipe used by the frozen native 0704 probe; it contains no MCE marker and should retain zero bytes. Both fixtures run the default one-MiB policy, an explicit zero policy, and a one-byte policy for two repeats. Default/off and default/tiny revision, semantic, patch, and published-text parity is asserted inside the binary.

Run the candidate lane from the repository root after the production API candidate and the standalone lockfile are present:

```sh
python3 docs/performance/results/change-0704/source-census.py candidate
python3 docs/performance/results/change-0704/build-retention.py
python3 docs/performance/results/change-0704/run-retention.py
```

The build command uses `/home/zhuhe/code/litchi-target-0704` and publishes the observer as `/home/zhuhe/code/litchi-0704-bin/retention-observer`. `build-retention.py` records a complete file census of the PPTX, shared OOXML, and OPC production trees, the entire retention probe workspace, the build inputs, the binary hash, and deterministic source hashes. It records the Rust subset separately and requires that it match `source-census-candidate.json`; all other production files have a separate census and hash. It rechecks all of those maps after the build. `run-retention.py` rechecks them before execution, writes the raw receipts to `retention-probe/results/real.tsv` and `retention-probe/results/generated.tsv`, and records `retention-build.json` and `retention-runs.json`. The binary also compares the two fresh repeats field-for-field, including direct stage charges, release observations, semantic values, revisions, and exact serialized patch bytes.

The standalone lockfile is deliberately separate from the workspace `Cargo.lock`. The only owned build scratch is the target directory and copied observer binary named above; the raw observer outputs and hash receipts are packet evidence and remain retained. No result from this lane supports a performance, RSS, concurrency, native Office, or cross-platform claim.
