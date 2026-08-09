# litchi-markdown

Format-agnostic Markdown emission helpers for the Litchi office-formats library.

## Overview

`litchi-markdown` provides bounded, exact-source CommonMark/GFM snapshots and
reversible block edits, plus the `ToMarkdown` trait and configuration types
used by Litchi's higher-level format adapters. It deliberately has no knowledge
of any concrete document format.

Snapshots expose source-ranged top-level blocks, nested block nodes, inline
nodes, and a deterministic link/image/footnote reference graph. Transactions
can stage multiple disjoint block, nested-block, and inline replacements plus
appends, validate reference-definition dependencies, and publish atomically
into bounded undo/redo history. Versioned JSON patches carry bounded,
replay-verified semantic operations; independently prepared patches can be
joined or inspected through a non-mutating three-way merge plan with
deterministic structured conflicts.

Reference-aware transfer preflights recursively include definitions missing
from a destination and refuse conflicting definitions before returning a fully
validated commit. Projection preflights report exact source ranges for raw
HTML, tables, task lists, and footnotes unsupported by a declared target.

```rust
use litchi_markdown::reader::Snapshot;

let source = Snapshot::read("# Original\n\nBody")?;
let mut edit = source.edit();
edit.replace_block_with_text(0, "# literal heading marker")?;
let commit = edit.commit()?;
assert_eq!(commit.snapshot().source(), "\\# literal heading marker\n\nBody");

let restored = commit.snapshot().apply(&commit.patch().inverse())?;
assert_eq!(restored.snapshot().source(), source.source());
# Ok::<(), litchi_markdown::reader::Error>(())
```

The offline release gate vendors all 652 normative CommonMark 0.31.2 examples
and all 670 examples from a pinned GitHub `cmark-gfm` specification. It checks
normative CommonMark rendering, supported GFM-extension rendering, exact source
ranges, deterministic reparsing, reversible edits, and project-authored
real-document roundtrips. Exact provenance, hashes, attribution, licenses, and
explicit semantic limits are recorded in `tests/data/PROVENANCE.md`.

## Reproducible release gate

Run the focused gate from the repository root:

```sh
sh crates/litchi-markdown/scripts/release-gate.sh
```

## Usage

```toml
[dependencies]
litchi-markdown = "0.0.1"
```

```rust
use litchi_markdown::{MarkdownOptions, TableStyle, ToMarkdown};

fn render<T: ToMarkdown>(value: &T) -> String {
    let opts = MarkdownOptions::new().with_table_style(TableStyle::Markdown);
    value.to_markdown_with_options(&opts).expect("Markdown conversion failed")
}
```

## Features

- `ToMarkdown` trait for converting types to Markdown.
- `MarkdownOptions` plus `FormulaStyle`, `ScriptStyle`, `StrikethroughStyle`, `TableStyle` enums for tuning the output.
- Unicode helpers for rendering super- and subscript characters.

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
