# 0497 validation plan

This plan validates the bounded DOCX paragraph-tail filesystem route after the
focused source and test checkout is frozen. Run every command from the
repository root, with the release toolchain and the locked manifests. Keep the
Cargo lanes serial when they share a target directory or a performance CPU
lock.

The first focused DOCX lane is:

```sh
cargo test --release --locked -p litchi-docx \
  --test source_backed_tail_append_stream -- --test-threads=1
```

It exercises `ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` through real sibling-temporary
publication. The cases cover exact bytes and reopen for new and existing
destinations, durable forward patch and inverse authorization, replacement of
the open source path, a Unix hardlink alias, stale and mutated sources,
cancellation, a late managed output-budget failure after a temporary prefix
was accepted, symlink and nonregular destinations, Unix permission
preservation, temporary-file cleanup, and managed memory/object release. The
late failure must report accepted partial output while leaving the destination
unchanged; a preflight-only limit refusal is insufficient evidence for that
case.

Run the existing OPC atomic adapter tests separately so the injected
post-replacement directory-sync failure remains visible:

```sh
cargo test --release --locked -p litchi-opc --lib atomic -- --test-threads=1
cargo test --release --locked -p litchi-opc \
  --test source_backed_path \
  --test source_part_splice \
  --test source_part_splice_replay \
  --test source_backed_topology \
  --test source_read_ahead \
  --test source_artifact_restore \
  --test source_part_transfer \
  --test phys_pkg_borrowed \
  --test python_zip64_source \
  --test external_zip64_source -- --test-threads=1
```

The OPC cases cover raw member and framing preservation, source freshness,
replay publication, exact inverse restoration, path-backed sources, aliases,
ZIP64 preservation, and the `OpcError::Committed` contract. The DOCX route has
no public parent-directory-sync fault injector, so the generic OPC unit test is
the evidence for that branch; adding a route-specific test would require an
unrelated injection API.

Run the full DOCX default and all-feature gates:

```sh
cargo test --release --locked -p litchi-docx --lib
cargo test --release --locked -p litchi-docx --all-targets -- --test-threads=1
cargo test --release --locked -p litchi-docx --all-features --all-targets -- --test-threads=1
cargo check --release --locked -p litchi-docx --all-features --all-targets
cargo clippy --release --locked -p litchi-docx --all-features --all-targets -- -D warnings
env RUSTDOCFLAGS=-Dwarnings cargo doc --release --locked -p litchi-docx --all-features --no-deps
```

The focused compatibility vector can be rerun before the all-targets lane to
make failures easier to attribute:

```sh
cargo test --release --locked -p litchi-docx \
  --test source_backed_tail_append \
  --test source_backed_tail_append_stream \
  --test source_backed_managed_document_edit \
  --test source_backed_managed \
  --test source_backed \
  --test source_backed_semantic \
  --test source_backed_story_text \
  --test source_backed_secondary_story_text \
  --test source_backed_paragraph_copy \
  --test source_backed_paragraph_removal -- --test-threads=1
```

Run the corresponding OPC compile, test, lint, and documentation gates:

```sh
cargo test --release --locked -p litchi-opc --all-targets -- --test-threads=1
cargo test --release --locked -p litchi-opc --all-features --all-targets -- --test-threads=1
cargo check --release --locked -p litchi-opc --all-features --all-targets
cargo clippy --release --locked -p litchi-opc --all-features --all-targets -- -D warnings
env RUSTDOCFLAGS=-Dwarnings cargo doc --release --locked -p litchi-opc --all-features --no-deps
```

The existing perf harness owns the route-level production checks. Its binary
tests run the same `run_case` path for deterministic, memory-store, and
file-store authored providers, and verify that counting and atomic publication
remain separate routes with exact candidate, semantic, untouched-member, and
inverse oracles:

```sh
cargo test --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin docx_replayable_tail_append -- --test-threads=1
cargo clippy --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --all-targets -- -D warnings
env RUSTDOCFLAGS=-Dwarnings cargo doc --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --lib --no-deps
cargo test --release --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --features allocator-metrics --lib -- --test-threads=1
```

Run the retained DOCX stream fuzz lane through its standalone, content-bound
helper. Preparation copies the unchanged target and the authenticated 0485
90-seed corpus into `/home/zhuhe/.cache/litchi-goal-0497/fuzz-target/fuzz`,
binds the manifest to the 0497 `after` checkout, and copies the primary fuzz
lock. The build uses the private
`/home/zhuhe/.cache/litchi-goal-0497/fuzz-target/target` directory and the
0485 ASan/coverage flags; it never shares the normal Cargo target. The run
command retains each starting corpus, mutated post-corpus, raw stdout/stderr,
artifact directory, command vector, source binding, and binary hash. It runs
exactly 10,000 cases at `-max_len=65536` for seeds 497 and 498.
The retained executable is
`/home/zhuhe/.cache/litchi-goal-0497/retained/fuzz/source_backed_tail_append_stream`;
cleanup may remove the private package, Cargo target, run copies, and TMPDIR
after verification while preserving that executable and the evidence-side
seed records.

```sh
python3 -B docs/performance/results/change-0497/fuzz.py prepare
python3 -B docs/performance/results/change-0497/fuzz.py build
python3 -B docs/performance/results/change-0497/fuzz.py run
python3 -B docs/performance/results/change-0497/fuzz.py verify
```

The coordinator owns the build and fuzz executions. A successful `verify`
requires both runs to exit zero, retain every authenticated starting seed, and
contain no sanitizer failure marker in the raw terminal. The helper's custody
receipt binds the copied target, standalone manifest, lock, after-crate tree,
and retained 0485 source records. These commands provide bounded sanitizer
evidence for this target; they do not claim arbitrary corpus coverage, native
application acceptance, or a performance improvement.

The ordinary fuzz target compile checks remain useful for the workspace lanes:

```sh
cargo check --release --locked \
  --manifest-path crates/litchi-docx/fuzz/Cargo.toml \
  --bin source_backed_tail_append_stream
cargo check --release --locked \
  --manifest-path crates/litchi-opc/fuzz/Cargo.toml \
  --bin parse_opc
```

These commands prove that the workspace fuzz targets still compile against the
changed APIs. The standalone helper above is the separate sanitizer-backed
evidence and retains its own seed, toolchain, terminal, and source/build
receipts.

Run the Python evidence validators after the source/build/capture manifests
are frozen:

```sh
python3 -B -m unittest discover \
  -s docs/performance/results/change-0497 \
  -p 'test_*.py' -v
```

For the optional syscall profile, verify the retained profile without starting
another child:

```sh
python3 -B docs/performance/results/change-0497/profile.py verify strace1
```

The profile is descriptive evidence. Its atomic scope must contain the actual
destination, `.litchi-*.tmp` sibling, file synchronization, rename, and parent
directory synchronization; its counting scope must contain no atomic
destination or sibling. The report must retain the raw trace and terminal
records and must remove the private replay and temporary trees before the
profile is accepted.

Formatting is a separate repository gate and must be recorded against the
same frozen source manifest:

```sh
cargo fmt --all -- --check
```

The result is validation evidence only. It does not establish crash
durability, a Windows same-path or hardlink guarantee from Unix-only cases,
native Word round trips, a performance improvement, or a speedup claim for the
new atomic route. The timing matrix must keep the atomic route as an after-only
capability unless an independently valid before implementation exists.
