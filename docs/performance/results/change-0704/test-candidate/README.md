# 0704 retention test candidate

`mce_retention_tests.rs` is the focused test module for the bounded opened
PPTX slide-MCE projection candidate. The source copy is included by
`opened/mod.rs` only under `#[cfg(test)]`; this directory keeps the review
artifact identical to that integrated module.

The module assumes this narrow `cfg(test)` hook under `opened::model`:

```rust
pub(crate) mod mce_retention_test_hooks {
    pub(crate) fn slide_output(
        snapshot: &super::Snapshot,
        slide_index: usize,
    ) -> Option<std::sync::Arc<Vec<u8>>>;

    pub(crate) fn checked_charge_for_test(
        entries_capacity: usize,
        output_capacity: usize,
    ) -> Option<usize>;
}
```

The hook must return only a retained owned default-profile slide output. It
must return `None` for marker-free, disabled, released, or budget-fallback
entries. Tests compare the returned `Arc` by identity and use public snapshot
methods for semantic and budget assertions. The hook must not add a runtime
observer, counter, global cache, or ordinary public API. The checked-charge
hook must call the same checked production arithmetic used for table admission
so the tests exercise entry-metadata multiplication and output-plus-metadata
overflow directly.

The module covers:

- default/disabled policy, marker-free admission, exact and minus-one budget
  boundaries, tiny ceilings, and checked arithmetic overflow;
- output capacity greater than logical length and aggregate metadata charging,
  with partial admission falling back without refusal;
- fresh snapshot isolation, clone/release/drop lifetimes, no-op commit and
  publication, and one/two-edit unchanged-hit/changed-miss Arc identity;
- aggregate partial admission and semantic/revision preservation;
- root, name, notes, relationship, and content-type typed refusals, including
  changed metadata candidates that reuse the same slide payloads; and
- a foreign `Part` whose visible `blob()` does not alias `blob_arc()`, plus a
  separate equal-byte allocation, proving foreign source ownership is not
  pinned or reused.

The counted foreign-part fixture compares the cached and disabled capture
routes and asserts that retention adds no `blob()`/`blob_arc()` observation.
The private capture helper used by the metadata tests keeps all validation in
the normal publication path; it does not bypass root, relationship, notes, or
content-type checks.
