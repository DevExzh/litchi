# Fresh complete API capture and glossary return

Fresh all-feature, hidden-item rustdoc capture at `a3ec9799939071d535925133c483bc998aa9667c` completes all 68 crates with no reused artifacts. The source remains clean and uses the pinned 139,606-byte 4100 root lock. The coordinator independently verifies every receipt, manifest, raw artifact and log hash, package identity, required command flags and the sequential target-size chain. The raw JSON totals 1,098,607,324 bytes outside the repository; peak reported RSS is 2,995,376,128 bytes and the final target is 3,749,318,288 bytes. The [capture verification](results/0362ul18/capture-verification.json) records each external artifact's path, size and SHA-256.

The [84-row inventory](results/0362ul18/inventory-a3ec.tsv) and blank disposition scaffold are byte-identical to the preserved 082 capture. Independent regeneration also matches. A blank scaffold is not a reviewed disposition: its 84 unassigned rows fail the disposition validator. The separate [reviewed dispositions](results/0362ul18/reviewed-disposition-a3ec.tsv) retain the previously reviewed decisions, and the [independent review](results/0362ul18/independent-source-role-review.json) revalidates their exact identities and source roles. Of 21 proof-source files, twenty are unchanged; the changed IWA codec removes six obsolete aliases, and the reviewer rechecks all three affected wire-request rows against the surviving canonical APIs. No engineering work is reclassified as a deferral.

The [source-equivalence receipt](results/0362ul18/assembled-source-equivalence.json) binds the captured API to assembled candidate `c796d0acc24c7c6e6fd08980b9e31af6bb86b18b`. All production crate sources, crate manifests and root manifest match; the sole changed crate file contains one exact test-only rustfmt expansion. Both actual root locks match. Benchmark consumer repairs and documentation changes do not change that captured API. Later source repairs require their own validation; this receipt does not silently rebind them.

The supplemental glossary sweep runs all 59 registered profiles against the corresponding fresh all-feature JSON without Cargo or new rustdoc. Fifty-five profiles pass and four report findings. The [sweep receipt](results/0362ul18/glossary-sweep.json) retains the exact failed commands and hashes and binds all external results; the coordinator verifies all 118 output-log hashes. This supplemental sweep does not replace the gate's feature-specific builds.

Two DOCX profile failures identify one real API defect: the public `ArtifactStream::stream_to` implementation exposes `ArtifactStreamPublication`, violating the existing `*Publication` ban from 0355a. Engineering return 0362ul19 must retain the streaming capability while using a canonical report type and migrating every consumer together. The soapberry profile also has six stale scopes that omit `office`; those scope corrections must preserve the bans. The Keynote-minimal finding is an all-feature mismatch: `__invalidate_slide_preview` is gated by `internal-iwork-source`, absent from the actual minimal profile, and Keynote-full passes. It requires no source change or policy waiver.

To regenerate the capture, use a clean checkout of the recorded source and pinned root lock, fresh external output/target paths, and:

```sh
env -u LITCHI_CAPPED -u LITCHI_BLESS \
  CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true \
  LITCHI_MEM_MAX=20G \
  python3 tools/final_api_inventory.py --root "$CAPTURE_SOURCE" \
  --output "$CAPTURE_OUTPUT" --target-dir "$CAPTURE_TARGET" \
  --cargo /home/zhuhe/code/litchi-worktrees/.coord/bin/cargo --no-reuse
```

The full gate continues on its own immutable c796 checkout, including eventual feature-specific glossary checks and actual corpus verification. The raw report rename, native qualification and final glossary closure remain required. Only redaction audit 22 is deferred. Nothing is pushed.
