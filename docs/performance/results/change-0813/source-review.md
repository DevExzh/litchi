# 0813 source review: direct `Result<Event>` matching

This is a read-only review of the single 0813 candidate. The candidate is
limited to `crates/litchi-pptx/src/notes/codec.rs`; it changes the scanner's
event dispatch from `read_event().map_err(xml_error)?` followed by an
`Event` match to a direct match on `read_event()` with `Ok(Event::...)` arms
and one `Err(error)` arm. It does not change the public API, parser
configuration, dependency graph, test oracle, unsafe code, or any iWork
source. No Cargo, formatting, test, workload, or profiler command was run for
this review, and this document makes no timing or adoption claim.

The archived before and after files are the complete production source used
by the 0813 candidate archive. Their SHA-256 identities are:

* before: `466e504588843a0b2538fc25ce3378b6fb9bf13aa9b0d74efebcb73ca0ebab5e`
* after: `007c2429b44027c37bb6f648f8a2841bfce7ef74b164aa70ed08fd06248b4f0d`

The first differing byte is inside `scan_processed_xml`, at the event match.
The complete source prefix through `fn scan_processed_xml(` is byte-identical
in both archives (SHA-256
`c049ebdb54fda7e558206af17f35b0d3cc22183b8064cf208df94689972fe75a`). The
complete `#[cfg(test)] mod tests` suffix is byte-identical as well (33,693
bytes, SHA-256
`78e16e72fbb50e24dc682f4420a593cce5f1d613bd8dbcbf8650ddd75579bb40`). Thus
the candidate has no unreviewed changes in the non-scanner prefix or in any
test helper or test case.

## Changed transport boundary

The before source contains:

```rust
match reader.read_event().map_err(xml_error)? {
    Event::Start(element) => { /* existing body */ }
    // existing Event arms
}
```

The after source contains the mechanically equivalent dispatch shape:

```rust
match reader.read_event() {
    Ok(Event::Start(element)) => { /* existing body */ }
    Ok(Event::Empty(element)) => { /* existing body */ }
    Ok(Event::End(_)) => { /* existing body */ }
    Ok(Event::DocType(_)) | Ok(Event::PI(_)) => { /* existing refusal */ }
    Ok(Event::CData(_)) => return Err(invalid("CDATA is rejected")),
    Ok(Event::Eof) => break,
    Ok(_) => {},
    Err(error) => return Err(xml_error(error)),
}
```

Every existing event arm body is preserved. The candidate only moves the
error conversion into an explicit arm. `read_event()` still returns a
borrowed `Event` over the processed input slice, so this spelling does not
introduce a scratch buffer, an owned event, or a changed lifetime. The parser
continues to be `Reader::from_reader(processed)` with the same configuration.

The loop state machine remains in the same order. `pending_pop` is consumed
before each read, including a read returning `Eof` or a parser error. `Start`
and `Empty` still push the resolver before node, depth, root, or attribute
work. `Empty` and `End` still set `pending_pop` at the same point. Resolver
errors still pass through `quick_xml::Error::from` and `xml_error`; parser
errors now take that same `xml_error` conversion in the explicit `Err` arm.
All depth, node, raw and processed byte, attribute, root, refusal, and
relationship checks remain in their prior order.

## Oracle and test boundary

The entire test module is byte-identical, including the independent buffered
`NsReader` oracle, the oracle inspector, differential outcome comparison,
malformed/refusal inputs, namespace-scope cases, and limit-ordering tests.
The candidate therefore does not weaken the semantic comparison while it
changes the production event dispatch. The retained test-only `NsReader`
import and oracle are test machinery; they are not evidence that the
production scanner still uses `NsReader`.

The candidate is specifically intended to answer the 0812 mechanism
hypothesis: whether matching the `Result<Event>` directly changes the release
scanner's event-result copy chain. A different source spelling, a smaller
binary, or a changed instruction count is only mechanism evidence. It cannot
establish a workflow benefit. If ordinary and profile disassembly show no
reduction in the relevant event-result movement, this hypothesis should be
rejected and the exact baseline restored. If code generation changes, the
candidate still requires the full six production gates, the independent
qualification and numerical readers, the native and allocation gates, and
the final disposition policy before retention.

The 0811 frame-pointer perturbation and sampled instruction attribution limit
any interpretation of the mechanism result. This review does not assign a
phase fraction, causal instruction cost, universal speedup, tail, RSS, or
cold-cache claim. The final disposition must remain external to this static
review and must agree with the retained source after the candidate trial.

**Static review result: PASS.** The archived source shows one narrowly scoped
direct `Result<Event>` match, byte-identical non-scanner prefix and complete
test suffix, and no source-level change to resolver, limits, errors, or
oracle semantics. Runtime gates, code-generation evidence, independent
replay, resource checks, and retention remain root-owned.
