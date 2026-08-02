# litchi-ppt

`litchi-ppt` owns the typed reader, writer, and host projections for legacy
Microsoft PowerPoint (`.ppt`) compound files.

The crate depends on `litchi-cfb` for storage, `litchi-odraw` for OfficeArt,
`litchi-ograph` for embedded chart grammar, and the shared security crates for
inert encryption, signing, and VBA handling. It has no dependency on DOC, XLS,
OOXML, an async runtime, or the former `litchi-ole` migration host.

## Usage

```toml
[dependencies]
litchi-ppt = "0.0.1"
```

```rust,no_run
use litchi_ppt::Package;

let mut package = Package::open("slides.ppt")?;
let presentation = package.presentation()?;
println!("slides: {}", presentation.slide_count());
# Ok::<(), Box<dyn std::error::Error>>(())
```

Use `litchi_ppt` directly or `litchi::ppt` through the umbrella crate. The
ownership move is intentionally breaking: there is no `litchi_ole::ppt` or
`litchi::ole::ppt` compatibility path.

## Features

- Typed `.ppt` reader, writer, and checked mutation APIs
- Slides, masters, text, shapes, OfficeArt images, charts, media, and OLE objects
- Inert encryption, signing, and VBA project handling
- Optional `formula` feature for MathType (MTEF) extraction

## License

Licensed under the Apache License, Version 2.0. Part of the [Litchi](https://github.com/DevExzh/litchi) workspace.
