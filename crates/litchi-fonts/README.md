# litchi-fonts

Font discovery, loading, and subsetting for the Litchi office-formats library.

## Overview

`litchi-fonts` provides the font-handling layer used when generating Office
documents that need to embed or subset typefaces. It wraps `font-kit` for
system font enumeration, `allsorts` for OpenType table parsing, and
`roaring` bitmaps for compact glyph-coverage tracking. It is consumed by
the OOXML writer (`docx`/`pptx` font embedding) inside the
[Litchi](https://github.com/DevExzh/litchi) workspace.

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
use litchi_fonts::{FontLoader, FontError};

fn load(family: &str) -> Result<usize, FontError> {
    let loader = FontLoader::new();
    let font = loader.load_system_font(family)?;
    Ok(font.data.len())
}
```

## Features

- System font discovery via `FontLoader`
- OpenType property extraction (panose, charset, family, pitch, Unicode signature)
- Glyph-set collection through the `CollectGlyphs` trait, backed by `RoaringBitmap`
- Pluggable font subsetting via the `FontSubsetter` trait

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
