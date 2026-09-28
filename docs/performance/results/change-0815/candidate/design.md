# 0815 borrowed event arm binding candidate

This directory is an archive-only source candidate based on commit
`55bb2ead34` (`perf(pptx): attribute remaining event arm payload copies`). It
contains snapshots of only
`crates/litchi-pptx/src/notes/codec.rs`; the production worktree remains
unchanged. `candidate/before/codec.rs` is an exact copy of the current
production source at that commit. `candidate/after/codec.rs` contains the
bounded source experiment and is not a production adoption or a performance
claim.

The candidate keeps the existing direct `Reader::read_event()` result match
and changes only the `Start` and `Empty` bindings in `scan_processed_xml`:

```rust
Ok(Event::Start(ref element)) => { /* existing body */ }
Ok(Event::Empty(ref element)) => { /* existing body */ }
```

Since each `element` is already a `&BytesStart`, the two arms pass
`element` directly to `resolver.push` and `inspect_element`. This removes the
four redundant `&element` expressions while preserving the event lifetime
through both calls. The patch therefore has six source spelling changes:
two `ref` bindings and four argument borrows. Resolver push/pop scheduling,
node and depth limits, attribute accounting, root tracking, event order,
typed error mapping, and all refusal branches remain unchanged.

The buffered scanner oracle, `inspect_element_oracle`, differential corpus,
refusal tests, and all test source are byte-identical to the before snapshot.
No test is added because the existing borrowed-versus-buffered coverage
already exercises both `Start` and `Empty` paths and their error/refusal
boundaries. The candidate introduces no dependency, unsafe code, ambient
state, public API, ownership, or budget change.

Root owns the source-format check, focused and full quality gates, fresh
ordinary/profile evidence, semantic and resource review, and the final
production disposition. This archive does not authorize retaining the
rewrite without those gates and makes no timing, allocation, RSS,
instruction, or workflow claim.
