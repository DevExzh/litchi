# 0481 DOCX scanner local-name borrow source review

This review covers the proposed `scan_document` change in
`crates/litchi-docx/src/source_backed/paragraph_copy.rs`: remove the three
`checked_clone(..., "XML local name")` calls in the `Start`, `Empty`, and
`End` event arms, and pass a borrowed local name to `next_scope` and the
scope comparison. The review is read-only and is a source/correctness handoff;
it is not measurement evidence.

## Lifetime constraint and required form

`quick_xml`'s `BytesStart::local_name()` and `BytesEnd::local_name()` return a
copyable `LocalName<'_>` whose bytes are borrowed from the current parser
event. The event itself is backed by `buffer`, which is cleared only after the
whole `match` arm completes (`paragraph_copy.rs:1593-1710`). No local name is
stored in `Layout`, `word_namespace`, the scope stack, or any returned error.
The borrow is therefore valid for every use inside the arm and ends before
`buffer.clear()`.

There is one Rust lifetime trap in a tempting shorthand. Binding
`start.local_name().as_ref()` or `end.local_name().as_ref()` directly to a
slice produces `E0716` because the temporary `LocalName` wrapper is dropped at
the end of the statement. The planned candidate avoids this by first binding
the wrapper in each arm, for example:

```rust
let local = start.local_name();
// ...
if local.as_ref() == b"document" { /* ... */ }
let scope = next_scope(&stack, local.as_ref(), false)?;
```

The same form applies to `Empty` and `End`. Alternatively,
`local_name().into_inner()` is an explicit borrowed-slice projection in the
pinned `quick_xml` version, but retaining the `LocalName` binding makes the
temporary lifetime and the three event arms visibly symmetric. The wrapper
contains only the existing event slice; binding it performs no allocation.

## State machine and error behavior

With the wrapper binding, the scanner compares exactly the same unqualified
bytes against `document`, `body`, `p`, `r`, and `t` through `next_scope`, and
the `End` arm compares the same bytes against `scope_local`
(`paragraph_copy.rs:1766-1800`). Prefix stripping remains owned by
`quick_xml`; no namespace or lexical spelling is normalized. The start/empty/end
event offsets, paragraph ranges, body boundary, stack transitions, and final
well-formedness checks are independent of local-name ownership.

The optimization intentionally removes the recoverable allocation failure for
the resource named `XML local name`. It also changes allocation-sensitive error
precedence: under an injected allocation failure, the old code could fail while
copying a local name before resolving its namespace, whereas the borrowed form
cannot fail at that point and reports the same namespace/scope refusal that an
ordinary successful allocation would reach. Parser syntax errors still become
`ComplexDocument` before an event is produced; XML-byte, event, depth, and
paragraph limits remain in the same checks; stack/range/Word-namespace
allocations retain their existing `Error::Allocation` paths. This is a
deliberate strengthening of the hot scanner's ownership behavior, not a
semantic change to malformed-input acceptance.

The `Start` arm must continue to own `word_namespace`: that value survives
multiple event iterations and cannot be replaced by a borrow from the current
event. Only the per-arm local tag name is safe to borrow. The borrowed slice
must not be moved into `word_namespace`, `Layout`, a patch, or a returned
diagnostic.

## Contract and ownership checks

The change is consistent with ADR 0003's immutable snapshots and borrowed
views, ADR 0005's bounded parser/resource policy, and ADR 0006's fail-closed
validation and preservation rules. It does not alter source bytes, paragraph
range offsets, source identity, dialect resolution, attribute validation,
publication, or the exact lexical fragments later copied by the edit path.
The XML-byte limit bounds the backing source, while the event reader owns the
current event buffer; clearing that buffer after the arm cannot invalidate any
persisting scanner state because none is retained.

## Existing coverage and review gap

The focused DOCX suite already exercises resource limits and failure atomicity
(`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:260-327`), complex
paragraph/document and hostile-namespace refusals (`:428-473`), both Strict
and Transitional namespace declarations (`:475-518`), exact paragraph ranges,
inverse replay, and publication preservation. The shared scanner is also
reached by durable patch validation, so those paths should be included in the
same release check.

I found no meaningful functional regression test that should be added solely
for this borrow. A test asserting an allocation count or the absence of a
temporary `Vec` would mirror the implementation and belongs in the measured
allocator harness. The focused suite should still compile and run against the
actual candidate, while the direct `let local = event.local_name().as_ref();`
form should remain avoided; inline comparisons such as
`event.local_name().as_ref() == root` are valid because their borrow does not
outlive the expression.

## Verdict

The planned candidate is semantically safe and removes three bounded,
per-event local-name allocations from the scanner. Its wrapper-first binding
has no source-level correctness blocker. Root should run the focused DOCX
tests and the existing scanner/publication evidence against the fresh
candidate source; no additional parser work is required for this change.
