# 0362uva — Final gate tooling registration packet

This source-only packet prepares the next safe-gate runner while the active
runner remains owned by the coordinator. It is based on staging commit
`cf8e83e7169f5d0d005e4f33eb54c4701889a2eb` and does not change `PROGRESS.md`,
the active runner, its receipts, or any Cargo/build limit.

## Registration

The prepared runner keeps 53 steps. Its `tooling-unit` command contains the
original nine modules exactly once:

```text
tools.test_check_api_glossary
tools.test_final_api_inventory
tools.test_officecli_corpus
tools.test_officecli_capability_delta
tools.test_no_3rdparty_references
tools.test_check_crate_boundaries
tools.test_non_iwork_gate
tools.test_unified_bench_matrix
tools.test_check_iwork_leaf_public_api
```

It adds these five modules, also exactly once:

```text
tools.test_api_debt_inventory
tools.test_validate_api_disposition
tools.test_w8_file_provenance
tools.test_w8_original_baseline
tools.test_w8_projection_policy
```

The prepared command starts with the glossary, inventory and disposition
modules, matching the persistent CI step in `.github/workflows/rust-ci.yml`.
The verifier checks the complete 14-module
argv, set equality, uniqueness, and the unchanged 53-step count. The runner
diff changes only this `tooling-unit` argv; the helper and Cargo wrapper are
copied byte-for-byte.

## External identity

The active runner was not touched. Its current and prepared identities are:

| Artifact | Active/current bytes | Active/current SHA-256 | Prepared bytes | Prepared SHA-256 |
| --- | ---: | --- | ---: | --- |
| `run_safe_gate.py` | 32,791 | `0691c5cb469d9f59caf67c63b748e4d2c80ea5b58e9b99cbde0e171fa5e8a372` | 32,966 | `0d8111c04bc9666566dab7c4f281a9f3fd7bb41efd64bd92703450abeee0bf14` |
| `safe_chunked_tests.py` | 8,750 | `70fdb56c96e4ed9e1d1e0d2b1e9c7f1db148a2115fb25246d13464569290cc35` | 8,750 | `70fdb56c96e4ed9e1d1e0d2b1e9c7f1db148a2115fb25246d13464569290cc35` |
| `bin/cargo` | 2,068 | `495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d` | 2,068 | `495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d` |

The prepared external copy is
`/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-next-20261024/`.
The tracked minimal patch is
`docs/unified-ops/results/0362uva/safe-gate-tooling.patch`, 1,614 bytes,
SHA-256
`cceb3b7dbcc0c6e3b90f4b12a32652cd9f7f3fb9b07be71f1a5b81746fd1bcbe`.
Applying it to a fresh copy of the active runner reproduces the prepared
runner byte-for-byte.

## Build-free verification

The exact command was:

```text
python3 -B docs/unified-ops/results/0362uva/verify_safe_gate.py \
  --runner /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-next-20261024/run_safe_gate.py \
  --root /home/zhuhe/code/litchi-worktrees/uo2-coordinator-metadata-integration \
  --scratch /home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-next-20261024/verification-tooling
```

It passed with:

```text
PASS: 53 steps, inventory tests, cap isolation, preserved child/failure receipts
```

The verifier is 14,047 bytes with SHA-256
`5c4583bdef66b4eda17f4fac5b3e56e98ab02223750150b9f9fef64daaa827c0`.
The post-commit external log is
`/home/zhuhe/code/litchi-worktrees/.coord/replay-integration/gate-safe-next-20261024/verification-tooling-postcommit.log`,
81 bytes with SHA-256
`ddeed317add455825af1593f549c74a979cb3715a4a17792b5911f770db06b11`.
No Cargo command was run. The prepared copy is for coordinator adoption only
after the active gate exits; this record makes no gate-pass or acceptance
claim.
