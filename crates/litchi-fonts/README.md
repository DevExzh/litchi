# litchi-fonts

Font discovery, OpenType metadata, glyph subsetting, Office embedding, and
OOXML font obfuscation for the Litchi office-formats library.

## Overview

`litchi-fonts` provides the format-independent font-handling layer used when
generating Office documents that need to embed or subset typefaces. It wraps
`font-kit` for system font enumeration, `allsorts` for OpenType table parsing,
and `roaring` bitmaps for compact glyph-coverage tracking. It is consumed by
the document-family writers (`docx`/`pptx` font embedding) inside the
[Litchi](https://github.com/DevExzh/litchi) workspace.

The API is organized by ownership:

- `discovery::Loader` resolves system faces and extracts typed metadata.
- `subset::mapping` maps Unicode requests to glyph IDs; `Allsorts` reduces
  OpenType programs.
- `embedding::prepare` creates owned, validated font programs, while
  `embedding::powerpoint::data` publishes the EOT wrapper.
- `obfuscation::apply` and `obfuscation::remove` implement the OOXML transform.

## System Dependencies

On Ubuntu/Debian, install the FreeType and Fontconfig development packages:

```bash
sudo apt install pkg-config libfreetype6-dev libfontconfig1-dev
```

## Usage

```toml
[dependencies]
litchi-fonts = "0.0.1"
```

```rust
use litchi_fonts::{FontError, Loader};

fn load(family: &str) -> Result<usize, FontError> {
    let loader = Loader::new();
    let font = loader.load_system_font(family)?;
    Ok(font.data.len())
}
```

## Features

- System font discovery via `discovery::Loader`
- OpenType property extraction (panose, charset, family, pitch, Unicode signature)
- Glyph-set collection through the `CollectGlyphs` trait, backed by `RoaringBitmap`
- Unicode-to-glyph mapping and pluggable font subsetting via the `subset` module
- Owned, license-checked font preparation and PowerPoint EOT publication
- SIMD-accelerated OOXML font obfuscation

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
