# litchi-imgconv

Bounded decoders and renderers for image formats embedded in Microsoft Office documents.

OfficeArt BLIP and BStore grammar lives in `litchi-odraw`. This crate consumes
its borrowed `image::Blip` views and performs RFC1950 decompression, DIB/WMF
adaptation, EMF/WMF/PICT rendering, resizing, and raster encoding under explicit
resource ceilings.

Host crates can expose borrowed or move-owned `litchi_odraw::image::File`
values. Import `litchi_imgconv::Convert` to add the short `decode`, `png`,
`jpeg`, `svg`, and `extract` codec operations without introducing a dependency
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

The default limits are safe for ordinary Office documents and can be replaced
through `Options::limits` when a trusted workload needs larger images.

Licensed under the Apache License, Version 2.0.
