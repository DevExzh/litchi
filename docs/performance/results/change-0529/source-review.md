# 0529 XML attribute probe candidate source review

status: UNAPPLIED production candidate; bounded implementation review

The candidate implements the reviewed zero/one-attribute probe in a disposable
copy of `crates/xml-minifier/src/audit.rs`. The repository Rust source remains
unchanged. The candidate is intended for later measurement only; this review
makes no performance or production-retention claim. OLE2/OOXML remains the
active priority, with ODF deferred and iWork excluded.

## Identity and custody

| item | identity |
| --- | --- |
| baseline revision | `3f4c7be06159dc3d742d9c815800dc55bf7c2db6` (`3f4c7be06`) |
| baseline source | `crates/xml-minifier/src/audit.rs` |
| baseline source SHA-256 | `40c7dfea199aa8f2e1951d0d75cd9a47e84f13bf49f2abcd1232230b64b355ac` |
| baseline source Git blob | `fb1ee700a1b218ef832d582dbd946be2fc78a529` |
| candidate copy | `/home/zhuhe/litchi-goal-0529-draft/crates/xml-minifier/src/audit.rs` |
| candidate copy SHA-256 | `507feadd1d48e4a5a1c26a009137f8ec46ff7a9d922d0ea5a5b6c1f11229ca89` |
| candidate source Git blob | `a7e4b4ee9fc6707b6432158931c316ec745dab49` |
| unapplied patch | `production.patch` |
| patch SHA-256 | `ca0970a03ab76e0bd6ebffd4ced9501fadb656ac4bb55acb515a2371d471e509` |
| source manifest | `source-manifest.json` |

The patch contains one release-path file and applies cleanly under
`git apply --check`; no hunk was applied. The disposable copy is the only Rust
candidate source. No build, test, benchmark, allocator run, or capture was
performed.

The candidate is bound to the following dependency inputs:

| input | SHA-256 or identity |
| --- | --- |
| workspace `Cargo.toml` | `911a52cf6932b81550bc9ffd6e522c327dec178297ee9bc46d2ccedb693d4885` |
| `crates/xml-minifier/Cargo.toml` | `7cdf4fac21a50ce8c39fc20d0af89c3ca24f869b74d95c6cd449249f0e28b797` |
| workspace `Cargo.lock` | `9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a` |
| `quick-xml` lock entry | `0.41.0`, checksum `e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1` |
| quick-xml `Cargo.toml` | `0e7df0b5caa523509bb47a3ce3cb282e49e64dcf59e39cd12ab8e3aacd732700` |
| quick-xml `src/events/attributes.rs` | `c46448f11d7dba312e6ad2177dc31deab2ce51282e9f1e4069d3b35175c18518` |
| quick-xml `src/events/mod.rs` | `b5e38bbfc0d87b2fa49b2dedcdb1e14d42064e7d41049c1b568d6f86f7731abb` |

## Exact candidate shape

The patch imports quick-xml's public `Attribute` type and changes the shared
`inspect_attributes` helper at `audit.rs:1504` as follows:

1. It creates `tag.attributes()`, disables duplicate tracking on that probe
   with `with_checks(false)`, and reads at most two iterator results.
2. `None` on the first result returns the inherited `Space` immediately.
3. A first `Ok(attribute)` followed by `None` drops the probe and sends the
   saved attribute through the new `inspect_attribute` helper. That helper is
   the sole owner of the existing aggregate `Attributes` charge and
   `xml:space` decode/validation logic.
4. A first `Err`, or any second result (`Ok` or `Err`), drops the probe and
   calls `inspect_attributes_checked`, which creates a fresh default checked
   iterator from the tag start and sends every successfully parsed attribute
   through the same helper.

The probe does not mutate `State`, charge an attribute, or decode `xml:space`.
It is dropped before single-attribute processing or checked replay, so the
probe's iterator state and quick-xml duplicate scratch cannot coexist with the
fallback. The patch does not add a duplicate checker, dependency, public API,
or unsafe code.

## Static contract review

Both verifier entry points already call `inspect_attributes`: the streaming
`Start`/`Empty` arms at `audit.rs:827-860` and the authored slice arms at
`audit.rs:1326-1357`. `check_start` still runs before it in both paths, and
all event, position, token, depth, text, and stream-buffer logic is untouched.
No other attribute consumer is changed.

The existing checked loop's observable ordering is retained by
`inspect_attribute`: it charges `Resource::Attributes` with `checked_add` and
only then decodes and validates `xml:space`, preserving the current
`Error::Limit`, `Error::Malformed`, and streaming `StreamError::Audit`
boundaries. The inherited scope is returned for an empty tag, and one valid
attribute receives exactly the same charge and `Space` transition as before.

The pinned quick-xml 0.41.0 implementation supports the fallback proof.
`BytesStart::attributes` creates a fresh iterator
(`events/mod.rs:286-289`); `with_checks` only flips
`IterState::check_duplicates` (`events/attributes.rs:580-589`). Syntax
parsing is unchanged in `IterState::next` (`events/attributes.rs:1222-1358`),
and duplicate checking occurs before parsing a value
(`events/attributes.rs:1287-1293`). One successful attribute cannot be a
duplicate. A second successful result may be a duplicate or another attribute,
and a second error may be the malformed second value; both cases therefore
replay the checked iterator. Replaying also preserves duplicate-before-
malformed-value precedence and the existing bounded error text. A first error
cannot be a duplicate because no earlier key was accepted, so its replay has
the same result.

The common helper affects both slice and streaming auditors. The current source
creates a new quick-xml attribute iterator for each tag in both paths; the
streaming path reuses its event/capture buffers but does not reuse quick-xml's
per-tag duplicate-key `Vec`. The candidate therefore removes that key-vector
work only for the proven zero/one-result case and adds a probe plus replay for
tags with two or more results. This is a measurement hypothesis, not a
removable-work or throughput claim.

## Risks and required validation

No source-level contract blocker was found in this bounded implementation
review. Compilation remains required before application because this review
was explicitly source-only, especially to confirm the borrowed `Attribute`
value remains usable after `drop(probe)` and the helper's inferred lifetimes.

Before any application or acceptance, focused differential checks should cover
both slice and arbitrary-chunk streaming inputs with:

* zero attributes, one ordinary attribute, one `xml:space` attribute with both
  valid values, escapes, normalization, and invalid values;
* malformed first attributes, duplicate names, a valid second attribute,
  malformed second values, and duplicate-before-malformed precedence;
* duplicate checks below and above quick-xml's 32-name threshold;
* exact aggregate attribute-limit boundaries, including limits reached on the
  first or later accepted attribute, plus token/error offsets and typed error
  details; and
* compactness/layout errors, nested inherited space, BOM and split chunks,
  authored versus streaming report parity, and the existing memory envelope.

The candidate must remain unaccepted until those checks and the frozen
OLE2/OOXML performance protocol establish useful end-to-end evidence. The
repository source is still at the exact baseline revision and the patch remains
unapplied.
