# litchi-imgconv

Bounded vector-first conversion for image formats embedded in Microsoft Office documents.

OfficeArt BLIP and BStore grammar lives in `litchi-odraw`. This crate consumes
its borrowed `image::Blip` views and performs RFC1950 decompression, DIB/WMF
adaptation, EMF/WMF/PICT conversion, resizing, and raster encoding under explicit
resource ceilings. EMF and WMF raster output is produced by first generating SVG
and rasterizing that SVG; the crate never substitutes a placeholder image.

Host crates can expose borrowed or move-owned `litchi_odraw::image::File`
values. Import `litchi_imgconv::Convert` to add the short `decode`, `png`,
`jpeg`, `webp`, `svg`, and `extract` codec operations without introducing a dependency
from the format-neutral OfficeArt crate back to this codec crate.

```rust
use litchi_imgconv::{Options, to_png};
use litchi_odraw::image::Blip;

fn render(blip_bytes: &[u8]) -> litchi_core::error::Result<Vec<u8>> {
    let blip = Blip::parse(blip_bytes)
        .map_err(|error| litchi_core::error::Error::ParseError(error.to_string()))?;
    to_png(&blip, Options::default().width(800))
}
```

For raw metafiles, use `convert_metafile` for a typed result that carries its
actual format, MIME type, extension, and selection report. `OutputFormat::Auto`
is vector-first: it selects SVG unless the parsed metafile is exclusively bitmap
painting, in which case it selects PNG.

```rust
use litchi_imgconv::{InputFormat, Options, OutputFormat, convert_metafile};

let emf = std::fs::read("diagram.emf")?;
let output = convert_metafile(&emf, InputFormat::Emf, OutputFormat::Auto, Options::default())?;
assert_eq!(output.mime_type, "image/svg+xml");
# Ok::<(), Box<dyn std::error::Error>>(())
```

The default limits are safe for ordinary Office documents and can be replaced
through `Options::limits` when a trusted workload needs larger images. The
format-specific legacy helpers remain available and now use the same bounded
SVG-backed raster path.

See [FORMAT_COVERAGE.md](FORMAT_COVERAGE.md) for the specification inventory,
bitmap matrix, strict-fidelity policy, and the operations that intentionally
return `Unsupported` instead of producing misleading SVG.

Licensed under the Apache License, Version 2.0.
