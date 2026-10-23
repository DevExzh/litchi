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
- Cargo, rustdoc, the API glossary capture, and native performance runs were
  intentionally not started while the coordinator's target and host-disk hold
  was active. The coordinator must run the full all-feature DOCX consumers,
  glossary/inventory capture, Clippy, and the paired raw lifecycle probes
  before acceptance.

## Runtime receipt

The allocated DOCX Cargo lane supplied the following bounded test receipts.
The copied logs are outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-helper-final-evidence/` and
are intentionally not tracked.

| Run | Source commit / tree | Command | Exit | Tests | Log size / SHA-256 |
| --- | --- | --- | ---: | ---: | --- |
| Glossary focus | `7fb00ee4ad2f6229613b58ad78d548bc116f0a6d` / `9e8e69f0529f5e2e639fa6361412f508200615dd` | `cargo test -p litchi-docx --lib glossary::tests -- --test-threads=1` | 0 | 55 passed | 5,178 bytes / `e187a2fa98baa8dd2800d83b150bc0c3088466a2d95953fedbe27c6beff1f18b` |
| DOCX library | `d331043548139488f0127e245b945ba6dedf1b0b` / `52f0673013ea98a3793d162de2420af1cfbe8f2c` | `cargo test -p litchi-docx --lib -- --test-threads=1` | 0 | 1,487 passed | 133,008 bytes / `15f8474416a1330e11cb0eaa818c3115f7cb4d0cb3be2774fd2681ff03367306` |

Both commands used `PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH`,
`CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-docx-p0/target`,
`CARGO_BUILD_JOBS=1`, `CARGO_INCREMENTAL=0`,
`CARGO_PROFILE_DEV_DEBUG=0`, `CARGO_PROFILE_TEST_DEBUG=0`,
`RUST_TEST_THREADS=1`, `LITCHI_MEM_MAX=20G`, and an external `TMPDIR`.
The focused run used `TMPDIR=/tmp/litchi-docx-glossary-7fb00ee4ad`; the
library run used `TMPDIR=/tmp/litchi-docx-lib-54d6037fcc`.
The focused log is
`glossary-55.cargo.log` (`/tmp/litchi-docx-glossary-7fb00ee4ad/cargo.log`)
and the library log is
`docx-lib-1487.cargo.log` (`/tmp/litchi-docx-lib-54d6037fcc/cargo.log`).
The final target directory measured 2,195,105,726 bytes (2.1G).

After the full-library receipt, the source advanced through
`4e447dd42e`, `1a567f58c0`, and `7fb00ee4ad`; the final-source focused receipt
above covers those admission/validator-only changes. No test was rerun for
this record-only update.

No executable, fixture, or large performance artifact is part of this record.
