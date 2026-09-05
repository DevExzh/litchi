# Change 0414 evidence

This bundle records ZIP64 preservation capability and descriptive ZIP32
regression guards. It authorizes no performance improvement claim.

- Control: `6b632726b8abacb4eb28dc14f3211bc6265206aa`.
- Candidate: `a4b7f849b9f34ba000eb912c69e63bad03a71773`.
- `protocol.json` fixes the original eight-row ABBA capture, 300 samples/30
  warmups. `guard-recheck-protocol.json` records the follow-up triggered by two
  original threshold crossings, using 1,000 samples/100 warmups.
- `identities.json` binds protocol hashes, candidate source bytes and unchanged
  files responsible for qualified gates. Each report contains the executable
  hash, clean Git revision, environment, corpus and raw samples.
- `guards` and `guard-recheck` retain all reports/catalogs as lossless Zstandard
  originals, plus `/usr/bin/time -v` process observations and exact capture argv.
  `raw-artifacts.json` binds decompressed lengths and SHA-256 hashes.
- `checks` retains losslessly compressed test/build/gate logs, the exact workspace test lockfile,
  the test-only control patch and the control's sparse integration test source.

Verify the inventory from this directory with `sha256sum --check SHA256SUMS`.
From the repository root, recompute both summaries with Python 3 and `zstd`:

```sh
python3 docs/performance/results/change-0414/summarize.py > /tmp/0414-summary.json
cmp docs/performance/results/change-0414/summary.json /tmp/0414-summary.json
python3 docs/performance/results/change-0414/summarize.py --followup > /tmp/0414-followup.json
cmp docs/performance/results/change-0414/guard-recheck-summary.json /tmp/0414-followup.json
```

The script verifies lossless originals, corpus bindings, revisions, clean
worktrees, CPU affinity, sample counts/order, recomputed percentiles and matched
sink summaries. Its temporary decompressions are cleaned on exit. Sink summaries
are byte/write-count observations; they are not cross-revision output hashes.
The harness also performs its per-run deterministic output oracles.

For a fresh capture, use clean detached worktrees at the two revisions. Build
each harness separately with:

```sh
env CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 \
  RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes' \
  cargo +1.98.1 build --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
```

Copy each executable to a distinct location before another build can replace
it. Run the argv in the retained time files, in A1/B1/B2/A2 order and with
`RUSTUP_TOOLCHAIN=1.98.1`, from the corresponding clean worktree. Adapt temporary
paths to the replay machine. Serialize builds/tests/analysis and measurement.
Original captures used CPU 2 on the shared EPYC 9R45 KVM guest; this is a warm
generated in-memory guard, not cold/remote/native/concurrent evidence.

Candidate verification used the following commands with four build jobs and
the lockfile in `checks/Cargo.lock` (restore it only in a disposable worktree):

```sh
cargo +1.98.1 test --locked -p soapberry-zip -p litchi-opc -p litchi-odf-common --all-features --all-targets
env RUSTDOCFLAGS='-D warnings' cargo +1.98.1 test --locked -p soapberry-zip -p litchi-opc -p litchi-odf-common --all-features --doc
env RUSTDOCFLAGS='-D warnings' cargo +1.98.1 doc --locked --no-deps -p soapberry-zip -p litchi-opc -p litchi-odf-common --all-features
cargo +1.98.1 clippy --locked -p soapberry-zip -p litchi-opc -p litchi-odf-common --all-features --all-targets -- -D warnings -A clippy::chunks_exact_to_as_chunks -A clippy::err_expect -A clippy::bool_assert_comparison -A clippy::large_enum_variant -A clippy::redundant_pattern_matching
```

Tests: 1,202 passed; doctests: 43 passed, two ignored. New lint findings were
fixed. The unexempted Clippy gate remains red on pre-existing findings; command
exemptions do not modify repository lint policy. The full-workspace formatter
reports unchanged DOCX glossary fixture differences; both modified ZIP and OPC
file sets pass their edition-specific rustfmt checks. Boundary, coverage, strict
claim and report-classification logs are retained separately.

To reproduce the before-state failures in a disposable control worktree, decompress and apply
`checks/control-opc-tests.patch.zst`; create `crates/soapberry-zip/tests` and decompress
`checks/control-sparse-test.rs.zst` there as `preservation_zip64_promotion.rs`.
Using the same test lockfile, run:

```sh
cargo +1.98.1 test --locked -p soapberry-zip --test preservation_zip64_promotion -- --nocapture
cargo +1.98.1 test --locked -p litchi-opc topology_add_at_zip32_entry_count_boundary -- --nocapture
```

Both commands intentionally exit 101: the sparse offset promotion test fails
with `UnsupportedPreservation`, and both OPC count-promotion tests fail with
their existing typed refusals. Only test code is transplanted. The candidate
uses the committed, extended tests and succeeds, including repeated owned
append/reopen through 65,537 members.
