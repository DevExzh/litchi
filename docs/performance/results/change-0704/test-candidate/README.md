# 0704 retention test candidate

`mce_retention_tests.rs` is the focused test module for the bounded opened
PPTX slide-MCE projection candidate. It remains outside `crates/` until the
production owner adds the cache and the test-only Arc inspection seam.

The module assumes this narrow `cfg(test)` hook under `opened::model`:

```rust
pub(crate) mod mce_retention_test_hooks {
    pub(crate) fn slide_output(
        snapshot: &super::Snapshot,
        slide_index: usize,
    ) -> Option<std::sync::Arc<Vec<u8>>>;
}
```

The hook must return only a retained owned default-profile slide output. It
must return `None` for marker-free, disabled, released, or budget-fallback
entries. Tests compare the returned `Arc` by identity and use public snapshot
methods for semantic and budget assertions. The hook must not add a runtime
observer, counter, global cache, or ordinary public API.

The module covers:

- default/disabled policy, marker-free admission, exact and minus-one budget
  boundaries, tiny ceilings, and `usize::MAX` arithmetic;
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
