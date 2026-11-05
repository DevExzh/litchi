# 0362rm — DOCX raw codec boundaries

- Date: 2026-10-23.
- Status: source and tracked-consumer migration complete; runtime, allocation,
  and performance qualification remain pending the coordinator's released
  Cargo lane.
- Parent: 0362rc/0362r and the 0351 API contract.

The DOCX free package mutators and detached XML codecs now have one explicit
raw boundary. The normal package owner continues to use its focused `Edit`
families and publishes a package `Commit`. Direct callers that intentionally
work with an `OpcPackage` or detached codec use the corresponding `raw`
module and pass that family's typed `&Limits` policy.

## Public boundary

The following raw paths retain the historical capability with required typed
limits:

| Capability | Raw entry point |
| --- | --- |
| chart graph load/store | `chart::raw::{load,store}` |
| font XML and document font graph | `font::raw::{parse,write,read,put,remove}` |
| glossary XML and complete opaque graph | `glossary::raw::{read,write,load,put,remove}` |
| web-settings XML and package graph | `web::raw::{parse,read,write,put,remove}` |
| stylesWithEffects package graph | `styles::effects::raw::{load,put,remove}` |

Detached stylesWithEffects resources use `Resource::new(xml, &Limits)` and
`Snapshot::from_xml(owner, xml, &Limits)`. Web snapshots use
`web::Snapshot::from_xml(xml, &Limits)`. The former public
`Resource::from_xml` spelling and the old free package codec exports were
removed; all tracked consumers now use the canonical raw path or owner edit.
The public chart and glossary source snapshots also require their raw policy:
`chart::Snapshot::{load,read}(package, document, &Limits)` and
`glossary::Snapshot::{load,read}(package, &Limits)`. Their package reads pass
through the same raw admission checks before graph allocation.

The glossary graph remains transferable without flattening its opaque parts.
`glossary::raw::Part::new`, `replace_data`, and `set_relationships` also take
the required policy. Name, content-type, payload, relationship-count, and
relationship-metadata bounds are checked before the part retains new strings,
payload ownership, or a replacement relationship vector. Raw graph publication
checks newly serialized catalog XML once before staging it, and web/font raw
publication reuses the same serialized bytes for policy admission and package
staging. Untouched package payloads continue to be shared by the existing OPC
graph code.

The source-closure admission walks the complete owned relationship graph under
the selected policy before creating the retained `HashSet`, role map, queue, or
auxiliary projection. It charges the selected part, payload, relationship,
string, and aggregate metadata budgets as it discovers each target; owner main,
root, relationship ID, and literal target metadata are charged too. Direct
`Part::new` admission includes its own name/content-type metadata. Graph
matching retains owner relationship identity, so equal root XML with a changed
owner or changed auxiliary part cannot be accepted as an unchanged graph.
Glossary XML is admitted before either DOM parse, including the MCE-projected
bytes. Font parsing follows the same before-MCE and projected-byte admission;
the font count and text budgets are checked before semantic tree allocation.
Chart raw load/store route through the limits-aware package helpers, and the
web/effects package paths carry their selected policy through the source read.

The package modules retain only crate-private canonical implementation helpers;
there is no public alias for a former top-level `put`, `remove`, `store`,
`parse`, `read`, or `write` bridge. The raw XML scanners keep the format hard
ceilings while enforcing the caller's selected byte, depth, node, attribute,
event, text, graph, and metadata bounds.

## Static validation

- `rustfmt --edition 2024` passed for every changed Rust source file.
- `git diff --check` passed.
- Static searches found no tracked consumer of the removed public free codec
  exports or `Resource::from_xml`.
- The bounded Cargo receipt below covers the glossary raw-helper focus and the
  complete `litchi-docx` library. The coordinator must still run the full
  all-feature DOCX consumers, glossary/inventory capture, Clippy, and paired
  raw lifecycle probes before acceptance.

## Runtime receipt

The allocated DOCX Cargo lane supplied the following bounded test receipts.
The copied logs are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/` and
are intentionally not tracked.

| Run | Source commit / tree | Command | Exit | Tests | Log size / SHA-256 |
| --- | --- | --- | ---: | ---: | --- |
| Glossary focus | `e2c7bf5808b12262a8d1640a44060fe82dabf135` / `c86d4f5d01c9c9c68fbecceb5734d1a0c8d11896` | `cargo test -p litchi-docx --lib glossary::tests -- --test-threads=1` | 0 | 56 passed | 5,268 bytes / `a068fd70b38a9710cc90437733abc144ff396f7492a2a374901101429ebf3a82` |
| DOCX library | `e2c7bf5808b12262a8d1640a44060fe82dabf135` / `c86d4f5d01c9c9c68fbecceb5734d1a0c8d11896` | `cargo test -p litchi-docx --lib -- --test-threads=1` | 0 | 1,488 passed | 133,098 bytes / `1806f5df52d0402ed73549c27cbb648e45f14b70caf88fd0e8c380702c6dba88` |
| Glossary focus (generated-root preflight) | `e43e9a8e9a7557b455962fa23773ecb0c669ab36` / `af46fd46d8c02f3df315049d77105a42862da22a` | `cargo test -p litchi-docx --lib glossary::tests -- --test-threads=1` | 0 | 58 passed | 5,477 bytes / `fcc7dd3b6a3dae4575ccdc8169f481456dfc588363cec7b671667d2937fcb338` |
| Glossary focus (prepared admission reuse) | `3b3a7c66fbfa8b5364bfa9578ea60079d9a6ea62` / `f2564573ba96a3393ccd10e265f6eda27ea9001c` | `cargo test -p litchi-docx --lib glossary::tests -- --test-threads=1` | 0 | 59 passed | 5,507 bytes / `64a7533b1a24093c68ea2c27b89e9c35422815b09aa00d470f2e8fff467d0ba0` |
| DOCX library (final) | `3b3a7c66fbfa8b5364bfa9578ea60079d9a6ea62` / `f2564573ba96a3393ccd10e265f6eda27ea9001c` | `cargo test -p litchi-docx --lib -- --test-threads=1` | 0 | 1,491 passed | 133,446 bytes / `3b29d2ae77464350c7e6469c922f223c28c6dd52344e32c4f943149d0e7e75f5` |
| Raw public integration (diagnostic) | `3b3a7c66fbfa8b5364bfa9578ea60079d9a6ea62` / `f2564573ba96a3393ccd10e265f6eda27ea9001c` | `cargo test -p litchi-docx --test glossary_api -- --test-threads=1` | 101 | compile blocked before tests by pre-existing `dead_code` denies | 1,143 bytes / `4c635ec9b025d918b1f5f68ba9cf35c53b2c0314dde3ea0d504c033386bb29e0` |
| Raw public integration (lint-override diagnostic) | `3b3a7c66fbfa8b5364bfa9578ea60079d9a6ea62` / `f2564573ba96a3393ccd10e265f6eda27ea9001c` | `RUSTFLAGS=-Adead_code cargo test -p litchi-docx --test glossary_api -- --test-threads=1` | 0 | 1 passed | 4,338 bytes / `954ce4d662bfe801f05677eff19384e4b554004c318008be544c973176a0dc35` |

The seven bounded runs used `PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH`,
`CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-docx-p0/target`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`,
`CARGO_PROFILE_DEV_DEBUG=0`, `CARGO_PROFILE_TEST_DEBUG=0`,
`RUST_TEST_THREADS=1`, `LITCHI_MEM_MAX=20G`, and an external `TMPDIR`.
The earlier focused run used `TMPDIR=/tmp/litchi-docx-glossary-e2c7bf5808`,
the generated-root preflight focused run used
`TMPDIR=/tmp/litchi-docx-glossary-final-preflight`, the prepared-admission
focused run used `TMPDIR=/tmp/litchi-docx-glossary-prepared3`, and the library
run used `TMPDIR=/tmp/litchi-docx-lib-e2c7bf5808`; the final library run used
`TMPDIR=/tmp/litchi-docx-lib-final-3b3a7c66fb`, and the raw integration run used
`TMPDIR=/tmp/litchi-docx-glossary-api-final-3b3a7c66fb`; its lint-override
diagnostic used `TMPDIR=/tmp/litchi-docx-glossary-api-final-override-3b3a7c66fb`
and `RUSTFLAGS=-Adead_code`.
The earlier focused log is
`glossary-56.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/glossary-56.cargo.log`);
the generated-root preflight focused log is
`glossary-58-e43e9a8e9.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/glossary-58-e43e9a8e9.cargo.log`),
the prepared-admission focused log is
`glossary-59-3b3a7c66fb.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/glossary-59-3b3a7c66fb.cargo.log`),
the final library log is
`docx-lib-1491-3b3a7c66fb.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/docx-lib-1491-3b3a7c66fb.cargo.log`),
and the raw integration diagnostic is
`glossary-api-integration-3b3a7c66fb.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/glossary-api-integration-3b3a7c66fb.cargo.log`).
The lint-override integration diagnostic is
`glossary-api-integration-override-3b3a7c66fb.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/glossary-api-integration-override-3b3a7c66fb.cargo.log`).
The earlier library log is
`docx-lib-1488.cargo.log`
(`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/docx-lib-1488.cargo.log`).
The final target directory measured 2,564,794,152 bytes (2.4G).

The generated-root preflight rerun used source `e43e9a8e9a`; the prepared-
admission rerun at `3b3a7c66fb` covers the incoming-catalog preflight,
transaction preparation, and admitted-read reuse changes. The final full
library receipt at `3b3a7c66fb` passed all 1,491 tests. The existing
`glossary_api` integration target did not reach its test binary because the
workspace's denied-unused lint reports the pre-existing private helpers
`codec::read`, `codec::write`, `codec::package::plan_write`, and
`codec::xml::parse_dom`; blame places those helpers before this wave. With
only that pre-existing lint suppressed, its one raw-public test passed; this
override is diagnostic evidence, not a replacement for the normal gate. No
Rust source changes were made after the final runs. All copied logs remain
available for historical comparison in the same evidence directory.

No executable, fixture, or large performance artifact is part of this record.
