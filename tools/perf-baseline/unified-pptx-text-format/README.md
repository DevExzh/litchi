# Paired PPTX run formatting

This standalone harness (record 0252) compares one native primary
`Edit::set_text_properties` call on the first run of the first shape with one
unified `File::set("/slide[1]/shape[1]/paragraph[1]/run[1]", props)` call
setting `bold`, `size` (24pt) and `color` (1F4E79), on generated one-slide
PPTX corpora of 1, 10, 100, 1000 and 10000 direct text boxes, for both eager
and source owners, in ABBA order (default nine rounds).

```sh
BUILD=/tmp/litchi-unified-pptx-text-format-target
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-pptx-text-format/Cargo.toml
taskset -c 0 "$BUILD/release/unified-pptx-text-format-baseline" --rounds 9 > out.json
```

Corpus generation, opening, target parsing, property preparation, warm-up,
serialization and equivalence checks are outside the timer. Routes end in
`run_format_v1`. Never commit the compiled binary.
