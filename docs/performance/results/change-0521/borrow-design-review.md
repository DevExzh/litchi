# XLSX XML borrow design review

Reviewed `crates/litchi-xlsx/src/cell_values/validation.rs` and the local
`quick-xml` 0.41.0 source. The proposed change is sound for this function:
`NsReader<&[u8]>::read_event()` returns an event borrowing the input slice
(`reader/slice_reader.rs:75-76,244-246`), and
`NamespaceResolver::resolve_event` returns that event unchanged while a bound
namespace borrows only the resolver (`name.rs:933-967`). The existing
`into_owned()` and `reader.resolver().clone()` therefore add per-event copying
without changing the validation inputs.

`NsReader` deliberately delays namespace popping until the next read after an
`Empty` or `End` event (`reader/ns_reader.rs:64-98`). Resolving immediately
after `read_event()` preserves the current scope for both event kinds; the
current validation branches do not mutably access the reader while the
borrowed `ResolveResult` is live. Parser and namespace-push errors occur before
resolution, so their mapping and ordering remain unchanged. `Unknown` prefixes
still make the same owned allocation.

Focused tests should cover a prefixed child that temporarily rebinds a root
foreign prefix to SpreadsheetML, then an `End` and an `Empty` sibling after the
scope is popped; the child must pass and the later foreign-prefixed sibling
must reject. Also retain cases for an unknown prefix, invalid reserved binding,
and mismatched closing tag to confirm error behavior and ordering. No build or
capture was run for this source-only review.
