# litchi-markdown

Format-agnostic Markdown emission helpers for the Litchi office-formats library.

## Overview

`litchi-markdown` provides the `ToMarkdown` trait and configuration types used by Litchi's higher-level format crates (and the `litchi` umbrella crate) to render Office documents and presentations as Markdown. It deliberately has no knowledge of any concrete document format; per-format `impl ToMarkdown for ...` blocks live alongside their respective format crates.

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
