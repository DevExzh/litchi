# litchi-core

Shared types, traits, and utilities used by every format crate in the Litchi workspace.

## Overview

`litchi-core` is the foundation of the [Litchi](https://github.com/DevExzh/litchi)
office-formats library. It provides the unified `Error`/`Result` types,
file-format detection, BOM and encoding helpers, document metadata, and
length/style primitives shared across the OLE, OOXML, ODF, iWork, and RTF
format crates.

Most users will pull this crate in transitively via `litchi` or one of the
format crates rather than depending on it directly.

## Usage

```toml
[dependencies]
litchi-core = "0.0.1"
```

```rust
use litchi_core::{FileFormat, Result};
use litchi_core::bom::strip_bom;

fn inspect(bytes: &[u8]) -> Result<()> {
    let format = FileFormat::detect(bytes);
    let body = strip_bom(bytes);
    println!("format = {:?}, body = {} bytes", format, body.len());
    Ok(())
}
```

## Features

- Unified `Error` and `Result` types built on `thiserror`
- File-format detection via the `FileFormat` enum
- BOM detection and stripping for UTF-8/16/32 streams
- Document metadata, shape, and style/length primitives shared across formats
- SIMD-friendly helpers and zero-copy XML slice types

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
