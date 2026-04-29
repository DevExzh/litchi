# litchi-iwa

Apple iWork archive parser for `.pages`, `.numbers`, and `.key` files.

## Overview

`litchi-iwa` reads Apple iWork bundles using their IWA (iWork Archive) layout: a ZIP container holding Snappy-compressed, protobuf-encoded object streams along with media assets and metadata. It exposes a unified `Document` API that handles all three iWork applications, plus lower-level access to archives, the object reference graph, and structured content (tables, slides, sections).

## Usage

```toml
[dependencies]
litchi-iwa = "0.0.1"
```

```rust
use litchi_iwa::Document;

let doc = Document::open("document.pages")?;
let text = doc.text()?;
let stats = doc.stats();
println!("objects: {}", stats.total_objects);

let structured = doc.extract_structured_data()?;
println!("{}", structured.summary());
# Ok::<(), litchi_iwa::Error>(())
```

## Features

- Parse Pages, Numbers, and Keynote bundles from a path or in-memory bytes
- Snappy decompression and protobuf decoding of `.iwa` streams
- Text extraction across all iWork applications
- Structured-data extraction: tables (with CSV export), slides, sections
- Media asset discovery and extraction

## Build Requirements

This crate compiles protobuf definitions via `prost-build`. The `protoc` compiler must be available on `PATH`:

- Debian / Ubuntu: `apt install protobuf-compiler`
- macOS (Homebrew): `brew install protobuf`

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
