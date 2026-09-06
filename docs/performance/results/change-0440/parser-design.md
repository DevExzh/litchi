# 0440 ElementAttrs ownership review

Current `ElementAttrs` in `crates/litchi-odp/src/codec/parser/codec/xml/validation.rs` (banked 0210/0211/0212/0215) should retain its single shared `Attributes` iterator and per-attribute resolution cache. Historical 0212's cache removes repeated resolver calls and is already the central multi-lookup win.

## Smallest candidate

Change only `ResolvedAttributeNamespace::Bound(Vec<u8>)` to a borrowed resolver URI, for example `Bound(Namespace<'reader>)` (or `Bound(&'reader [u8])`). Make the cache carry two lifetimes: the event lifetime for `RawAttribute`/`LocalName`, and a resolver-borrow lifetime for the URI. `resolve_at_scan` receives `&'reader NsReader<&[u8]>` and stores the `Namespace` slice returned by `resolve_attribute` without `to_vec()`.

`ElementAttrs` already documents that no XML event may be read between `new`/the first `get` and its final lookup/harvest. The borrowed resolver URI is valid exactly under that condition: quick-xml stores namespace values in its resolver buffer and a later event can mutate/truncate it. Tying the second lifetime to `&NsReader` makes that invariant compiler-visible; it also prevents the caller from advancing a mutable reader while the cache lives. Existing callers only use the reader immutably until the helper returns, so no public API change is required. Keep `Unknown`/`Unbound` behavior and `local_name` borrow unchanged.

Preserve `resolve_attribute` call order. Continue mapping `Unknown(_)` to `Unknown`; do not change quick-xml's unknown-prefix allocation/error behavior in this tranche. Keep the `Vec<ResolvedAttribute>` cache: replacing URI ownership is the bounded one-change seam; a stack cache/index is a separate resource/ordering design.

## One-shot callers

`Parser::get_attr` previously constructed an `ElementAttrs` and therefore stored every scanned attribute in `parsed`, including the URI ownership copy, even though it returns after one match or exhaustion. The draft restores its old direct loop (the pre-0210 implementation), so one-shot page/image/text:s lookups do not allocate the cache. This direct loop preserves exact malformed-attribute order/message, first-match behavior, resolver call order, decoder normalization, and `None` exhaustion. It is kept separate from the borrowed cache so one-shot calls have no reader lifetime extension.

## Error/ordering invariants

- The iterator remains one shared document-order stream for multi-lookups; first matching attribute still wins.
- A malformed attribute still reports the generic `invalid XML attribute: ...` at the first `get` that reaches it; later misses replay the stored message. Harvest continuation still uses `invalid ODP shape attribute: ...`.
- Duplicate detection stays in quick-xml's iterator and occurs at the same scan position.
- Only matching values are decoded, with `XmlVersion::Implicit1_0` and the same decoder; URI borrowing must not cause a predecode or eager value owner.
- `drawing_attributes` traverses cached prefix then continuation in original order; modeled/foreign attributes remain skipped without value decoding.
- Default namespace remains excluded for attributes (`Unbound`); unknown prefixes remain nonmatches.

## Required focused coverage (source review only)

Compare the cache/direct oracle for: bounded resolver URI lifetime (including a later event only after cache drop), shadowed prefix, default-namespace unqualified attr, unknown prefix, bound custom URI, malformed-before/after match, duplicate-after-first-match, malformed value, repeated cached hit, and `drawing_attributes` order/error parity. Measure the 0440 candidate only after the frozen before matrix; reject it if normal p50 or medium/large allocation metrics fail the declared >=5% gate.
