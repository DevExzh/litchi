# 0440 ODP `ElementAttrs` parser audit

Reviewed at baseline `dd0681cf8` without changing the workspace or running
builds/tests/scripts.

## Current contract

`ElementAttrs` in `crates/litchi-odp/src/codec/parser/codec/xml/validation.rs`
shares one `quick_xml::Attributes` iterator per element. It caches raw
attributes, the local-name view, and an owned namespace result. The cache is
used by the multi-query handlers in `shape_builder`, event listeners, media,
transition properties/sound, and style-definition collection; one-shot
`Parser::get_attr` also goes through it. Shape handling then calls
`drawing_attributes`, which must combine the cached prefix and the remaining
iterator in source order.

The cache is sound only while the event and resolver scope are unchanged. No
caller may read another event, clear/reuse the event buffer, or pop/rebind the
namespace resolver between `ElementAttrs::new` and the final lookup/harvest.
The existing call sites satisfy this by doing straight-line queries inside the
current event. If namespace snapshots become borrowed, make that invariant a
Rust lifetime/API property by constructing the cache with the reader (or keep
an equivalent owner token); otherwise a future caller can retain a namespace
slice after `NsReader` advances.

## Hard semantic and security obligations

1. Resolve attributes with the attribute rule: an unprefixed attribute is
   `Unbound` even when a default element namespace is active. Prefix aliases
   must compare by URI, and a child rebind must win over an outer binding.
   Unknown prefixes remain non-matches with the current behavior. Do not use
   the existing `NsClass` unchanged: it omits `SVG`, `XLINK`, `XML`, and
   `SMIL`, all of which are queried by production handlers. Collapsing those
   URIs to `Other` would drop geometry, links, `xml:id`, and transition
   properties. The generic unit test also queries a bound custom URI, so a
   compact known-namespace token needs a lossless fallback (or the generic
   helper contract must be deliberately narrowed and tested).

2. Preserve lazy first-match/error order. A match before a malformed raw
   attribute succeeds; a query whose scan first reaches that malformed
   attribute returns `invalid XML attribute: ...`; a malformed attribute first
   reached by the drawing harvest returns `invalid ODP shape attribute: ...`.
   Duplicate detection must remain at the underlying `Attributes` iterator's
   position. Do not eagerly materialize or decode all attributes, because that
   would turn a later malformed/duplicate/value error into an earlier failure.

3. Preserve value semantics. Matching values must still use
   `decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())`,
   return owned `String`s, and normalize entities/whitespace identically.
   Nonmatching, modeled, and foreign attributes are skipped without decoding.
   Keep exact error wrappers for raw-attribute, value, and invalid local-name
   failures.

4. Preserve `drawing_attributes`: cached and uncached portions remain in
   document order; DRAW/SVG/DR3D/TABLE classification and modeled-name skips
   remain byte-identical; the first error and its wrapper remain unchanged.
   It must not harvest attributes already consumed after a failed value decode
   or hide a malformed continuation.

5. Keep existing parser limits and unknown-prefix handling. A compact cache or
   query index must not remove namespace-declaration/attribute/value limits,
   create a lossy UTF-8 conversion of raw URI/name bytes, or turn malformed
   namespace syntax into a successful known-namespace match.

There is one existing retry edge worth pinning in the differential oracle:
`get` consumes a matching raw attribute, then `lookup` can return a value-decode
error before `self.parsed.push(cached)` (lines 372–375). A second query on the
same `ElementAttrs` can therefore skip that attribute and return `None` instead
of reproducing the one-shot decode error. The harvest has the analogous
continuation ordering at lines 492–496. Normal callers abort on `?`, but a
replacement must not make this behavior less deterministic; ideally add a
retry test and retain the first error state.

## Focused differential/adversarial tests

- Run identical query traces against the current one-shot oracle and the new
  cache: queries in document order and reverse order, repeated queries,
  missing queries, and DRAW→PRESENTATION fallback. Include prefix aliases,
  nested shadowing, default namespace plus unprefixed attributes, empty
  bindings, unknown prefixes, and a bound custom URI.
- Put a matching attribute before/after: a malformed raw attribute, a duplicate
  raw QName, an invalid entity/encoding in a nonmatching value, and an invalid
  value in the matching attribute. Compare both returned values and exact
  `Error::InvalidFormat` text at every query in the trace.
- Compare `drawing_attributes` with the pre-0215 fresh-scan oracle for
  interleaved cached/uncached attributes, modeled attributes with bad values,
  foreign attributes with bad values, namespace shadowing, duplicate keys, and
  malformed errors in both the cached prefix and continuation. Assert source
  order and no double decoding.
- Cover every production attribute namespace used by `ElementAttrs`: office,
  draw, presentation, style, script, text, SVG, DR3D, table, SMIL, XLINK, and
  XML. Existing parser fixtures already exercise many aliases; keep those when
  changing the representation.
- Add a large-attribute adversary at the existing accepted limit and just over
  it where the handler has a limit, plus a long foreign namespace URI. Check
  bounded memory/typed rejection rather than allowing a new eager `Vec` or
  namespace copy path to grow without the current policy.

## Likely remaining work worth profiling

The current cache removes repeated `Attributes` iteration and resolver calls,
but each new query still linearly replays `self.parsed` from index zero. A
multi-query shape therefore remains O(n·k) cheap comparisons; a query index or
per-query first-match position could remove that work, but it must remain lazy
so malformed/duplicate/error precedence is unchanged. Matching values are also
decoded again for repeated queries. Finally, one-shot `get_attr` calls still
construct and grow a cache while scanning to one result. Measure these three
paths separately before claiming a larger optimization; the owned bound-URI
`Vec<u8>` copies are a distinct allocation/copy cost and can be removed with a
compact, lossless namespace representation.

## Root resolution after applying the candidate

Only bound namespace ownership changes: the cache retains arbitrary resolver
URI slices under an explicit reader lifetime, with no namespace classification
shortcut. The root removed the draft's one-shot lookup rewrite and inspected
the production diff for unchanged iterator advancement, duplicate/malformed
error handling, value decoding, cached-prefix order and harvest continuation.
The unrelated decode-error retry behavior is unchanged. Three focused tests
use a direct iterator/resolver oracle and check actual URI pointer identity,
default-namespace exclusion, and rebinding after the cache ends. All 18
attribute tests and 352 ODP tests pass; strict owner Clippy and rustdoc pass.
No dependency, public API, unsafe block, validation or publication path changed.
