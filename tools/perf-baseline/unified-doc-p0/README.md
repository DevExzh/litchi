# Paired DOC P0 operations (record 0258)

The standalone target builds `crates/litchi-ops/examples/doc_p0_bench.rs`.
It compares the primary native DOC owner (`litchi_doc::Package::edit` and the
native `Document` reader) with the unified `File` call that wraps it:

| Scenario | Native leg | Unified leg |
|---|---|---|
| `paragraph_text_set` | `edit.replace_paragraph` + commit + publish | `File::set("/body/p[N]", text)` + publish |
| `run_format_set` | `set_character_property(Bold)` + `set_character_value(Size)` on a `BodySpan` | `File::set("/body/p[N]/r[2]", {bold, size})` |
| `paragraph_insert` | `edit.insert_paragraph` | `File::add("/body", p at N)` |
| `paragraph_text_get` | `Document::body_paragraph_text` | `File::get("/body/p[N]")` + view text |
| `run_bold_get` | `Document::body_paragraph` + `runs()[1].bold()` | `File::get(".../r[2]")` + format `bold` |

Each document has `N` paragraphs (1, 10, 100, 1,000) of two runs; the text
replacement uses a sibling document of single-run paragraphs. Owners, parsed
paths and properties are prepared outside timing and reused, so story-index
memos are warm in both legs. Both legs' published bytes are compared once
before timing. Two warm-up rounds, then nine ABBA rounds: 18 samples per leg.

```sh
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=<target> cargo build --release \
  --manifest-path tools/perf-baseline/unified-doc-p0/Cargo.toml
taskset -c <cpu> <target>/release/unified-doc-p0-baseline --rounds 9 > receipt.json
```

Run at least three trials pinned to one idle CPU. Do not commit the binary;
record its SHA-256, the source commit and this build command instead.
