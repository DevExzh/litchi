# 0362dvha — DrawingML raw-fragment review repairs

This return packet repairs the independent review findings for the
0362dvh DrawingML raw-fragment migration. It is intentionally a successor
record: the original migration receipt remains unchanged.

The source commit for this repair was `12e99a7c0b`; the later PPTX
limit-direction correction is recorded separately in `0362dvhb`.

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

Focused diagnostic validation was run on the capped offline target. The
commands used the environment and target below; they passed `--offline` but
did not pass `--locked`:

```text
source (deleted-coord-artifact)
export CARGO_TARGET_DIR=(deleted-coord-artifact)
export CARGO_BUILD_JOBS=1 RUST_TEST_THREADS=1 CARGO_NET_OFFLINE=true
cargo check --offline --all-targets -p litchi-drawingml -p litchi-spreadsheet-drawing -p litchi-pptx -p litchi-xlsb
cargo test --offline -p litchi-drawingml --lib --test chart_formatcode2 --test chart_formatcode2_adversarial --test chart_formatcode2_xml_misc --test transform
cargo test --offline -p litchi-spreadsheet-drawing --lib shape
cargo test --offline -p litchi-drawingml --lib
cargo fmt --all -- --check
git diff --check
```

All listed commands passed after the repair. The ignored leaf `Cargo.lock`
observed during the later receipt check was 139,592 bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`;
12e's command-time pre/post lock hashes were not captured, so no pinned-lock
claim is made. No external raw command logs or command-time target byte
measurements were retained for this receipt. The shared target's current
post-inspection apparent size is 777,184,046 bytes. These are diagnostic leaf
checks (`6d`) only; assembled `4100` workspace validation, fresh inventory,
and native qualification remain coordinator-owned and required.
