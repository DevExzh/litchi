# unified-pptx-elements

Paired native-leaf versus unified-API measurements of the record 0321 PPTX
elements: theme, master name, table column width and insertion, shape
hyperlink writes, and hyperlink, equation and theme reads (routes and
equivalence checks in `src/main.rs`).

```sh
CARGO_NET_OFFLINE=true cargo build --release \
    --manifest-path tools/perf-baseline/unified-pptx-elements/Cargo.toml
taskset -c 7 target/release/unified-pptx-elements-baseline --rounds 9 [--counts 1,50] [--only ROUTE]
```

The JSON report goes to standard output; record 0321 keeps two runs under
`docs/unified-ops/results/0321/`. The binary is not committed.
