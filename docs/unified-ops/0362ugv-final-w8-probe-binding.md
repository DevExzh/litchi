# 0362ugv — W8 File probe identity binding

## Scope

This leaf repairs the stale identity pin used by the build-free W8 provenance
checks.  It does not change the production API, the probe workload, or any
native qualification result.  The source is pinned at
`2dfc4f134d94195dd9b41f2d806887b38ef10c58` and the repair is based on the
current probe source commit
`56dbd37551ba204e99199feeaf61a8dab848046d`.

The previous reviewed probe was the `d39d179353` source (74,747 bytes,
SHA-256
`3a01dd2cbb971b4eba54e2eccffa99e1b3c4557db14af03e61869f99f0a0cc27`).
The current probe differs from that source only by the XLSB consumer migration
required by the deferred selector API: three routes now call
`.resolve().expect("xlsb worksheet resolution")` before worksheet mutation
(six inserted lines, no changed lines).  The route vocabulary, 109-route
matrix, fixture constants and markers, `into_file()` extraction, native/facade
dispatch, publication phase, semantic assertions and timing wrappers are
unchanged.  The current probe is therefore the reviewed workload with a
necessary consumer-only API adaptation, rather than a workload or proof
substitution.

## Exact bindings

The current probe blob is 75,035 bytes with SHA-256
`21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0`.

The following executable bindings now use that identity:

* `tools/w8_file_provenance.py` updates `REVIEWED_PROBE_SHA256`.
* `tools/w8_original_baseline.py` updates the independently reviewed probe
  source commit, byte count and SHA-256.  The pristine original baseline
  remains `f2bbbc8d07ab5da26c180f1ff74813d39f632d54`; only the corrected probe
  source pin advances to the API-migrated consumer source.
* `docs/unified-ops/results/0362h/run.py` updates its fail-closed workload
  identity check so endpoint manifests generated for the current probe are
  accepted.

Historical 0362ha/he/hf receipts retain their original d39 source identity;
they describe the earlier qualification input and are not rewritten by this
binding repair.

## Frozen failure and validation

The frozen pre-repair protocol log was
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/native-protocol-tests-2df.log`
(14,075 bytes, SHA-256
`ee015827821a3c6aa231c93481384729d9255c35e287b1e451d5ed24221b426d`).  It
reported 11 errors, all from the stale reviewed-probe identity: the current
probe was `21c8e383...` while the tools and runner required `3a01dd2c...`.

After the binding update, the affected build-free checks were run with:

```text
python3 -m unittest -v tools.test_w8_file_provenance tools.test_w8_original_baseline
```

The preserved raw log is outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/final-w8-probe-binding/w8-binding-tests.log`:
3,141 bytes, SHA-256
`5e31470bd6bd178a03770f7bf7f95293dd0e69431b36dc4eb45c4f0dcb9bd8d7`.
All 18 tests passed, with zero failures, errors or skips.  The route tests
continued to prove the complete unique 109-route matrix and the adapter test
continued to prove only the four legacy result-field translations.

No Cargo command, native build, native execution, timing run or qualification
claim was made for this repair.
