# 0813 direct event success/error match candidate

This directory is an archive-only source candidate based on commit
`03305a82aa` (`perf(pptx): localize scanner event payload copies in native
samples`). It changes only the archived copy of
`crates/litchi-pptx/src/notes/codec.rs`; the production worktree remains
unchanged. `candidate/before/codec.rs` is an exact copy of the current
production source at that commit. `candidate/after/codec.rs` contains the
bounded source experiment and is not a production adoption or a performance
claim.

The candidate changes the `scan_processed_xml` event transport from

```rust
match reader.read_event().map_err(xml_error)? {
    Event::Start(element) => { /* existing body */ }
    // existing event arms
}
```

to a direct match on the reader result:

```rust
match reader.read_event() {
    Ok(Event::Start(element)) => { /* existing body */ }
    // every existing event arm is wrapped in Ok
    Err(error) => return Err(xml_error(error)),
}
```

Every successful event body, arm order, scope update, resolver operation,
limit check, element inspection, relationship collection, and refusal branch
is retained byte-for-byte. The `DocType`/`PI` alternative is represented as
`Ok(Event::DocType(_)) | Ok(Event::PI(_))`; the wildcard is `Ok(_)`. The only
new behavior spelling is the final error arm, which applies the same
`xml_error` adapter that `map_err(...)?` applied to a read failure. No reader,
resolver, namespace, attribute, buffering, output, or public API contract is
changed.

The candidate adds no test source. The existing buffered scanner oracle,
refusal corpus, and differential tests already exercise the scanner's event
success and error outcomes. Root owns the source-format check, focused and
full quality gates, fresh candidate measurements, semantic/resource review,
and the final production disposition. This archive does not authorize
retaining the rewrite without those gates, and it makes no timing, allocation,
RSS, instruction, or workflow claim.

The source review follows the accepted ADR index and the previously completed
architecture review, especially ADRs 0001, 0002, 0005, 0006, 0008, 0011,
0013, 0024, 0030, 0031, and 0032. The candidate introduces no dependency,
unsafe code, ambient state, public API, budget, or ownership change.
