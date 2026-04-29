# litchi-imgconv

Pure decoders and converters for the image formats embedded inside Microsoft Office documents.

## Overview

`litchi-imgconv` parses BLIP (Binary Large Image or Picture) records and
the metafile formats they wrap — Enhanced Metafile (EMF), Windows Metafile
(WMF), and Macintosh PICT — and converts them to modern raster formats
(PNG, JPEG, WebP) or SVG. It is a leaf crate in the
[Litchi](https://github.com/DevExzh/litchi) workspace, depending only on
`litchi-core` plus the `image`, `flate2`, `bytes`, `xml-minifier`, and
`zerocopy` crates. The OLE/Escher integration glue lives in the umbrella
`litchi` crate.

## Usage

```toml
[dependencies]
litchi-imgconv = "0.0.1"
```

```rust
use litchi_imgconv::{Blip, convert_blip_to_png};

fn render(blip_bytes: &[u8]) -> litchi_core::error::Result<Vec<u8>> {
    let blip = Blip::parse(blip_bytes)?;
    convert_blip_to_png(&blip, Some(800), None)
}
```

## Features

- BLIP record parsing (`Blip`, `BitmapBlip`, `MetafileBlip`, `BlipType`)
- BLIP store table parsing (`BlipStore`, `BlipStoreEntry`)
- EMF, WMF, and PICT metafile decoding to PNG/JPEG/WebP
- Optional SVG output for vector metafiles
- Optional resizing with high-quality Lanczos3 filtering

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
