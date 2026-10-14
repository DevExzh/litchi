# Paired SmartArt writes (record 0284)

This standalone harness compares native leaf SmartArt writes with the
unified verbs, on eager owners that already hold `N` SmartArt graphics
(default `0,10,50`; the node routes use at least one):

- `pptx_smartart_add_v1`: `opened::Edit::add_smartart(0, &Spec, "Diagram
  N+1", bounds)` and one commit, versus `File::add("/slide[1]",
  smartart{layout, text, x, y, width, height})`;
- `pptx_smartart_node_text_v1`: the hand-written leaf edit of the first
  graphic's second node (`Edit::smartart`, `Spec::from_parts`, the text
  change, `Spec::{data_xml, drawing_xml}`, `Edit::set_smartart`, one commit),
  versus `File::set("/slide[1]/smartart[1]/node[2]", {text})`;
- `pptx_smartart_node_add_v1`: the same leaf sequence appending a node,
  versus `File::add("/slide[1]/smartart[1]", node{text})`.

Corpus generation, opening, warm-up, serialization and the equivalence
check (exact bytes, else equal decoded archive members) are outside the
timed interval. Each sample is one operation; legs alternate ABBA; medians
are the upper middle.

```sh
BUILD=/tmp/litchi-unified-smartart-target
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR="$BUILD" cargo build --release --locked \
  --manifest-path tools/perf-baseline/unified-smartart/Cargo.toml
LITCHI_SOURCE_COMMIT="$(git rev-parse HEAD)" \
LITCHI_RUSTC_VERSION="$(rustc --version)" \
LITCHI_TARGET="$(rustc -vV | sed -n 's/^host: //p')" \
taskset -c 7 "$BUILD/release/unified-smartart-baseline" --rounds 9 > smartart-writes.json
```
