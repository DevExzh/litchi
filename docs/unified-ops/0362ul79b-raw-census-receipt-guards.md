# 0362ul79b — raw-census runner and resource receipt guards

## Review repairs

The 79a receipt bound the source, feature profile, tools, output, and ignored
lock hash, but two capture-integrity details were still self-attested. This
follow-up keeps the manifest incomplete and adds no owner or disposition
exception:

- The approved ignored `Cargo.lock` must be a regular file. Every component of
  its path is checked with `lstat`; a symlinked lock or symlinked parent fails
  even when the followed bytes have the approved SHA-256.
- Each profile receipt now binds the exact f5 runner, gate-local cap wrapper,
  and underlying memory-capped Cargo shim by absolute path and SHA-256. The
  wrapper and underlying shim must be regular executables and retain the
  target-trace and `systemd-run`/`MemoryMax` markers.
- The receipt carries the exact f5 environment identity, including offline and
  locked command policy, target path, debug/incremental settings, one build
  job, one test thread, 20G memory cap, external `TMPDIR`, attempt marker, and
  trace path. Fixed target/free-space/RSS guard values are copied from the
  approved policy and resource receipts must remain within them, including a
  recorded live-target peak and an explicit no-overflow flag.
- The receipt carries the cap wrapper's complete tab-separated execution trace
  with bytes/SHA-256. The validator reparses the actual file, compares the
  recorded rows byte-for-byte by parsed order, and requires every row to be a
  `rustdoc` invocation for the exact profile target. A changed command,
  target, trace, wrapper, or environment fails closed.

The trusted policy is recorded in `tools/raw_api_coverage.json` and currently
identifies runner SHA-256
`1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`, gate
wrapper SHA-256
`495e24c6f60b238378386d86ac4158cba904992ce8870a9494f2c4a0eedd256d`, and
underlying shim SHA-256
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`.
The final 66-profile manifest and these external identities must be
regenerated/rebound after the source freeze; this commit makes no completeness
claim.

## Source-only checks

```text
python3 -m py_compile tools/raw_api_coverage.py tools/test_raw_api_coverage.py
exit 0
python3 -m unittest tools.test_raw_api_coverage
Ran 16 tests ... OK
python3 tools/raw_api_coverage.py --root . \
  --manifest tools/raw_api_coverage.json \
  --glossary tools/api_glossary.json --json
exit 0; complete=false; source_failures=[]; current_source_failures=[];
missing_receipts=15; missing_inputs=15; missing glossary profiles=53;
missing_workspace_packages=0
```

Adversarial fixtures cover lock symlinks, wrong runner identity, changed
offline policy, altered trace command/target, stale source, wrong features,
wrong artifact bytes, wrong seeded lock, and ignored vendor/manifest/fixture
overlays. No Cargo, rustdoc, native build, or measurement was run.
