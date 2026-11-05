# 0362ul29 bounded gate tooling

This record contains only source and tests for the repaired external gate
helper. It is bound to the immutable production source
`9781cf6d87e4d261d1dd9670e144e2db127d869b`.

The failed predecessor is retained outside this checkout at
`full-gate-9781cf6d87/`. Its row 25 failure was the helper's RTF example
chunk, which reached `20,763,352,771` bytes before the old post-command
retention ran. The original receipts remain unchanged:

* row log SHA-256: `a8981c731f51841ffdd4446657f4407d5db79bbcdde7f5c005c630172e3baf6c`;
* Cargo trace SHA-256: `e1c8cf9044f3dfd81e1da0175ce30517e5344099698887e22337affd9a7d063c`;
* blocked chunk log SHA-256: `ad4e24317ad3672ac34058882a59ace6c694a017eea97a2098927b30c60b2cbe`.

## Retention contract

`safe_chunked_tests.py` enables Linux child-subreaper mode and waits for Cargo,
its complete process group, and every adopted descendant (including a child
that calls `setsid`) before moving any reported output. Successful
test/example ELF executables in exactly `target/debug/deps` or
`target/debug/examples` (including stale hashed variants from earlier feature
builds) and only their exact `.d`/`.rmeta` sidecars are atomically moved beside
the target into a per-invocation, per-chunk quarantine namespace.  The
namespace is derived from the resolved log directory and gate attempt, so
different suites may safely reuse chunk indices; a repeated direct invocation
reuses an existing artifact only after verifying identical bytes and mode.
their bytes, mode, SHA-256, original relative path, command and chunk index
are recorded in `manifest.jsonl`. Caches, rlibs, build outputs, shared
libraries, non-ELF files, `target/debug/build`, and release artifacts remain
in the Cargo target. A failed chunk stops the helper immediately; failed,
aborted or hung commands retain their artifacts in place and produce a blocked
receipt. If descendant cleanup has to escalate after its lifetime timeout, the
helper treats that cleanup itself as a failed chunk even when all descendants
are eventually reaped; no artifact from that chunk is retired.

Both the helper and the outer runner sample the target while the command's
root process is running. The outer runner also samples through the remaining
descendants after that root exits; the helper alone does not provide that
post-root live guard. The reviewed invocation therefore uses the outer runner
to stop its identity-guarded owned tree (including adopted `setsid` children)
at `19,000,000,000` bytes, before the hard `20,000,000,000` byte cap. A
command starts only below `18,000,000,000` bytes, with at least 12 GB free
disk. The helper uses chunk size 2; planning tests
assert that all test and example targets occur exactly once, so shrinking the
chunks cannot drop work. No historical target cleanup is used.
Target-size sampling retries transient `du` races caused by Cargo renames; a
persistent measurement error or any unexpected watcher exception is recorded
with its last sample and process cleanup, then fails the command and prevents
artifact retirement.

## Preserved command-count receipt

The immutable predecessor's `ops-tests/checks.json` is 98,519 bytes with
SHA-256 `bc8efb85911c0d4089a5f8687d4ca846fdb663c92209895d6db8ba63ba8a2c2d`.
Parsing its completed commands gives 167 `--test` selectors, 72
`--example` selectors, and one `--lib` command. At chunk size 2 this is 121
selector/build chunks (84 test chunks + 36 example chunks + the library
command), plus the one `--doc` command: 122 helper commands. This
checks-derived count is the recorded basis; the historical 132-command
statement is not used as a receipt or as evidence of completed work.

## Source binding and locks

The rerun must use a fresh target/state/quarantine with the exact source above
and these offline lock identities:

```
Cargo.lock                                      4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1
tools/perf-baseline/Cargo.lock                  eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c
tools/native-resave/Cargo.lock                  51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378
crates/litchi-core/fuzz/Cargo.lock              a9e134aabbb8ec0462c811d0f388f6260c5c4c8733d5038f550c456919d5bc3e
crates/litchi-ops/fuzz/Cargo.lock               ccf4cec8234ff5a72377ccbde14a25598a81ff49e4d8f2e9ffb68ef2e9e9a7aa
crates/litchi/fuzz/Cargo.lock                   26ad5846ddca5e564e636c65d6377b1eeb14c21c7278ea0efd14941ae403c35b
```

The exact immutable rerun command is:

```
python3 docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py \
  --root (deleted-coord-artifact) \
  --target (deleted-coord-artifact) \
  --state (deleted-coord-artifact)
```

This is a fresh 53-step selection. If a quiet-window rerun is split after
the non-isolated rows, the remaining selections are, in order:

```
tests-ops,tests-facade,tests-core,clippy-leaves,clippy-ods,tests-leaves,
tests-ods-full,clippy-workspace,tests-workspace-excluding-ods-ops,wsdoc,
ppt-alone,rustdoc-core-ops-facade,cli-tests,python-bindings,feature-matrix,
feature-matrix-clippy,fuzz,goldens,golden-clean,perf-harness-check,
native-resave-check,api-glossary,iwork-leaf-api,tooling-unit,crate-boundaries,
non-iwork-gate,bench-matrix,corpus-dispositions
```

The documentation selections remain isolated and must be granted a quiet
window; they are not skipped or relabeled from the previous guarded failure.

The copied runner/helper hashes are recorded by the runner's `plan.json`; the
original coordinator-only scripts and the failed source/target are not
modified by this record.
