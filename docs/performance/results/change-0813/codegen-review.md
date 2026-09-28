# 0813 code-generation review: direct `Result<Event>` dispatch

This is an independent static review of the 0813 source candidate and the
exact ordinary/profile scanner symbols recorded in
`codegen-before/` and `codegen-after/`. The review covers mechanism and source
semantics only. It does not make a workflow timing, allocation, RSS, tail, or
retention claim. No Cargo, formatting, workload, profiler, or binary command
was run while making this review.

## Source identity and semantic scope

The current worktree source is byte-identical to the archived candidate after
source (`codec.rs` SHA-256
`007c2429b44027c37bb6f648f8a2841bfce7ef74b164aa70ed08fd06248b4f0d`). The
archived before source is
`466e504588843a0b2538fc25ce3378b6fb9bf13aa9b0d74efebcb73ca0ebab5e` and the
candidate patch is
`da23ae2c0b7e92c0b58ad41c79b3617f3e430604bb6f5f7f8356936bcd11085b`.

The candidate changes one dispatch expression in `scan_processed_xml`:

```rust
match reader.read_event().map_err(xml_error)? {
    Event::Start(element) => { /* existing body */ }
    // existing event arms
}
```

becomes:

```rust
match reader.read_event() {
    Ok(Event::Start(element)) => { /* existing body */ }
    // existing event arms wrapped in Ok
    Err(error) => return Err(xml_error(error)),
}
```

The archived before and after files retain the same parser construction and
configuration. The `pending_pop` operation still runs before every read;
`Start` and `Empty` still push the resolver before node, depth, root, and
attribute work; `End` still schedules the resolver pop and decrements depth;
and the DTD, processing-instruction, and CDATA refusals remain in their prior
arms. `Eof`, the root/depth postcondition, limit checks, relationship
collection, checked-attribute handling, and all existing event-arm bodies are
unchanged. The new `Err(error)` arm applies the same `xml_error` adapter that
`map_err(xml_error)?` applied to a read failure.

The full test suffix in both archives is byte-identical (33,693 bytes,
SHA-256 `78e16e72fbb50e24dc682f4420a593cce5f1d613bd8dbcbf8650ddd75579bb40`).
That suffix includes the buffered `NsReader` differential oracle, malformed
and refusal corpus, namespace cases, and limit-ordering checks. This source
review therefore finds no changed parser, resolver, lifetime, error-ordering,
oracle, public API, dependency, unsafe-code, or iWork scope.

## Exact code-generation evidence

The packet’s `codegen-analysis.json` defines one bounded static window for each
leg and variant. The window starts at the sole
`quick_xml::reader::Reader<R>::read_event_impl` call in the recorded
`scan_processed_xml` symbol and ends at its first indirect `jmp *`, with a
maximum of 81 instructions. The counts below are properties of those exact
disassemblies, not runtime frequencies:

| leg | variant | scanner symbol bytes | symbol instructions | vector moves in window | binary SHA-256 |
| --- | --- | ---: | ---: | ---: | --- |
| before | native | 2,616 | 487 | 12 | `2ee9b24e15d65b83d68eef9783782906d52105cab8e10a8d82c2b16063b00bb6` |
| before | profile | 2,616 | 487 | 12 | `18b18523b28f39cfd0bb13516bf96fe0705e34b8b811f3ca53e35a71c54185bc` |
| after | native | 2,320 | 434 | 0 | `e2ae857291a6ed96400522f821d7f6d406d57a27bd55cbf62127fc6487550c58` |
| after | profile | 2,320 | 434 | 0 | `8b2fa9c6677382a4991ac8d91d15300bc0bb361806cee844e7f04487749ff1fa` |

In each before window, the reader return is followed by three successive
40-byte aggregate copy chains before event dispatch. They are visible as the
three groups of vector moves from the reader event slot at `0x68`, through the
`r14` and `r13` temporaries, and into the stack event slot at `0x90`, followed
by the indirect dispatch. The analysis counts 12 vector move instructions in
this pre-dispatch window. In each after window, the reader result’s error
discriminant is checked and the event discriminant is dispatched directly;
the three aggregate chains are absent and the window contains zero vector
move instructions. The first indirect dispatch moves from offset 714 in the
before rows to offset 615 in the after rows (relative offsets in the recorded
scanner symbols; the profile rows show the same shape).

The scanner symbol consequently shrinks by 296 bytes, or about 11.3 percent,
and 53 instructions, or about 10.9 percent, in both ordinary and profile
builds. The after assembly still contains payload moves in individual event
arms after dispatch; this review makes the narrower claim that the
`read_event_impl` to first event dispatch copy chain was removed. It does not
claim that every event payload move in the whole scanner disappeared.

The four rows are bound to the immutable per-leg receipts and assembly
artifacts in [`codegen-analysis.json`](codegen-analysis.json), whose source
and binary bindings are checked by [`codegen-gate.json`](codegen-gate.json).
The receipt identities are:

* before receipt: `0aee36660617395fdc9c4fa3c9953be8c50b5b68903de4464875808bc9fbf23a`
* after receipt: `2c1e37820da04d246c7e6645f6a701554955b028270472fa0631ea767b2aedef`

## Review result and limits

**Static code-generation result: PASS.** Both ordinary and profile release
symbols show the intended mechanism change: direct matching of the reader
`Result` removes the pre-dispatch aggregate copy chain identified in the 0812
instruction review. The source comparison shows that the error conversion and
successful event behavior remain represented by the same semantic branches.

This result authorizes the code-generation stage of the frozen 0813 protocol;
it does not authorize production retention by itself. The native and
allocation captures, owner-scoped profile publication, independent raw result
readers, resource checks, quality gates, and the adoption policy remain
required. Static symbol size and instruction-window counts cannot establish
runtime frequency, causal cycle savings, a universal speedup, or a semantic
equivalence proof beyond the source and existing oracle scope.
