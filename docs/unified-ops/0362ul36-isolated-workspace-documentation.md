# 0362ul36 — isolated workspace documentation passes

Date: 2026-10-24. The isolated `wsdoc` selection passes on frozen source
`9781cf6d87e4d261d1dd9670e144e2db127d869b`. This is row 34 of 53, not a
full-gate result. The actual workspace documentation includes
`litchi-iwa-protos`, whose generated crate index remains in the target.

The Cargo invocation is `doc --offline --locked -q --workspace --no-deps
--all-features`, with `RUSTDOCFLAGS=-Dwarnings`, one build job, incremental
compilation disabled, and the 20-GB memory-capped shim. It ran alone, exited
zero after 2,138.861 seconds, and produced no diagnostics. The source remains
clean and all six recorded Cargo locks are unchanged.

The final target occupies 3,380,852,837 bytes. The largest of 34,909 live
size samples is the same value, below the target limits. This is target
storage, not a peak-RSS measurement. Three post-exit observations show no
owned processes or builders and stable target size. The coordinator also
verified that runner PID 2239045, Cargo PID 2239350 and both guard identities
had terminated before reviewing the separate row-36 invocation.

## Retained evidence

Endpoint:
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781-isolated-wsdoc-f5f3-rerun/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `plan.json` | 24,733 | `807ebff8affb09930e7d4efe6899bc109e6b53f7968056b5fd60cd68ee8a6c72` |
| `state.json` | 7,659 | `72e7a5a8733ab992801d69c33a7c4568e51c850507d332751eb7e894509dae38` |
| `34-wsdoc-attempt-1792878617796613288-2239045.log` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `34-wsdoc-attempt-1792878617796613288-2239045-cargo-targets.tsv` | 129 | `d637acaded9d799a1f10f1db4b43a1ad327ad13d67fd44328df9556b24acaf26` |
| `target/doc/litchi_iwa_protos/index.html` | 4,775 | `f86c9c18b91d04350e1a76d89c3bdd0391dfc64e80a88dc9f95c8c90eccc9c71` |

The coordinator independently verifies these artifact identities, the passing
row's command and environment, source and lock checks, target observations,
and terminated process identities. The Rust crates, root Cargo manifest and
lock, Cargo configuration and Rust toolchain configuration have no diff
between tested 9781 and integration `49b83790a99f7d22c936e7825140601ce81d7140`.
New tooling registration is separately reviewed in 0362ul35; it is not
silently attributed to the frozen documentation source.

## Runner and reproduction

The run uses the retained f5 runner from
`/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/`:

- `run_safe_gate.py`: SHA-256
  `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`;
- `safe_chunked_tests.py`: SHA-256
  `3e03aa2227ac0002c3052cf5080d032cd2e9150f584c926b9aa0d499175ec67f`;
- `bin/cargo`: SHA-256
  `495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`.

Reproduce with that runner's `--only wsdoc`, the frozen source at
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781cf6d87/source`,
and fresh `--target` and `--state` directories. Keep compilation and native
measurement jobs stopped for the isolated documentation selection. Do not
reuse or overwrite the accepted endpoint.

The earlier `full-gate-9781-isolated-wsdoc-f5f3` endpoint is preserved as an
aborted partial invocation: Cargo and rustc started before the coordinator
handoff interrupted it. It is not accepted or described as setup-only. The
older resource and ops-gate failures likewise remain separate evidence.
Row 36 (core/ops/facade documentation), the remaining gate rows, actual corpus
replay and native performance qualification retain their own acceptance
requirements.

## Separate core/ops/facade documentation result

Row 36 also passes on unchanged 9781 after row 34's processes terminate.
Its command is `cargo doc --offline --locked -q --no-deps --all-features
-p litchi-core -p litchi-ops -p litchi`, using the same f5 runner, warnings
policy, memory cap and single-job environment. The isolated run exits zero
in 612.126 seconds with no diagnostics. All three generated crate indexes
exist. All six locks remain unchanged and the source remains clean.

Endpoint:
`/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/full-gate-9781-isolated-rustdoc-f5f3/`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `plan.json` | 24,683 | `d35fd6f91ff6a7c830f1a369262633049399a6acc20fb5a175752c10ed202673` |
| `state.json` | 7,786 | `318300698b440815aabd16584c42a2cddb486c86059b8cdd4da58578b8d8e07e` |
| `36-rustdoc-core-ops-facade-attempt-1792880788406988113-2531775.log` | 0 | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` |
| `36-rustdoc-core-ops-facade-attempt-1792880788406988113-2531775-cargo-targets.tsv` | 125 | `0a86bf35163cc805efed1b999730076f605f54b04e504086319f20fa52a30402` |

Final and maximum sampled target storage are 1,915,267,412 bytes across
23,415 live samples. No guard errors occurred. Three post-exit samples show
stable target size, no owned processes and no builders. Runner PID 2531775
and Cargo PID 2531997 have exited. The coordinator independently verifies
the receipts, passing row, source/lock equality, generated indexes and
terminated runner. Rows 34 and 36 are complete for their recorded source;
no other row or performance result is inferred from them.
