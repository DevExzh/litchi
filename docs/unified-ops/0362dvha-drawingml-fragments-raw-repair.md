# 0362dvha — DrawingML raw-fragment review repairs

This return packet repairs the independent review findings for the
0362dvh DrawingML raw-fragment migration. It is intentionally a successor
record: the original migration receipt remains unchanged.

`formatcode2` now admits rewritten and detached output from the selected
output policy before encoding or reserving the output buffer. `write_to` and
`write_attribute_to` stream source slices and bounded replacement fragments
directly to the caller's sink. The scanner carries the selected lexical and
decoded-value budgets through text, CDATA, XString decoding, bindings and
attribute projection; XML entity and attribute normalization are preflighted
without materializing an over-budget lexical value.

Color raw constructors (`Value::unknown` and `Unknown::from_xml`) require the
caller-owned `color::Limits`. Color writes preflight their exact typed output,
enforce selected transform/node/depth bounds, and revalidate opaque values
under the selected policy. Transform writes likewise preflight exact output
and stream sink output without a full intermediate buffer. PPTX shape
transforms and DrawingML text colors derive subordinate policies from their
retained owner limits; spreadsheet-drawing's public `Emitter::new` validates
its geometry policy before constructing state.

The repair adds negative admission tests for unchanged sinks, lowered value
and lexical budgets, typed and opaque color structural limits, and source
fidelity under the bounded streaming paths. No native workload or ABBA run was
performed in this source-repair lane.

Focused validation on the capped offline target:

```text
cargo check --offline --all-targets -p litchi-drawingml -p litchi-spreadsheet-drawing -p litchi-pptx -p litchi-xlsb
cargo test --offline -p litchi-drawingml --lib --test chart_formatcode2 --test chart_formatcode2_adversarial --test chart_formatcode2_xml_misc --test transform
cargo fmt --all -- --check
git diff --check
```

All commands pass after the repair. The final workspace gate, fresh inventory,
and native qualification remain coordinator-owned assembled-source checks.
