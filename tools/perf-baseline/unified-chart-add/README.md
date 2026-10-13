# Paired chart add (record 0270)

This standalone harness compares the native leaf chart add with one unified
`File::add` call, on eager owners that already hold `N` charts (default
`0,10,100`):

- PPTX: `opened::Edit::add_chart(0, &Spec, "Chart N+1", bounds)` and one
  commit, versus `File::add("/slide[1]", chart{chartType, title, legend,
  categories, data, x, y, width, height})`;
- XLSX: `Spec::to_xml` plus `Package::edit_charts`, `Edit::add` and one
  commit, versus `File::add("/Sheet1", chart{..., x, y, width, height})`.

Corpus generation, opening, warm-up, serialization and the equivalence
check (exact bytes, else equal decoded archive members) are outside the
timed interval. Each sample is one add; legs alternate ABBA; medians are the
upper middle.

```sh
BUILD=/tmp/litchi-unified-chart-add-target
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release \
  --manifest-path tools/perf-baseline/unified-chart-add/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
taskset -c 7 "$BUILD/release/unified-chart-add-baseline" --rounds 9 > chart-add.json
```

Route names: `pptx_chart_add_v1`, `xlsx_chart_add_v1`.
