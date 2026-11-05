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

let package = Package::open("slides.ppt")?;
let presentation = package.presentation()?;
println!("slides: {}", presentation.slide_count());
# Ok::<(), Box<dyn std::error::Error>>(())
```

`Package` and `Presentation` are immutable, cheaply cloned owners. Package
publication preserves the exact input bytes; encrypted pass-through uses
`to_encrypted_bytes()` explicitly. Low-level integrations that need the generic
mutable reader use `litchi_ppt::raw::Reader<R>`.

Opened changes use `Package::edit()` and the types in `package::edit`:

```rust,no_run
use litchi_ppt::package::edit::Position;
use litchi_ppt::Package;

let source = Package::open("slides.ppt")?;
let mut edit = source.edit()?;
edit.move_slide(Position::new(0), Position::new(1))?;
let commit = edit.commit()?;
let (updated, patch) = commit.into_parts();
updated.save("reordered.ppt")?;
let restored = updated.apply(&patch.inverse())?;
# Ok::<(), Box<dyn std::error::Error>>(())
```

Edits retain source policies and native dependency checks. Dropping an edit
leaves the source unchanged; a failed staged operation preserves earlier
successful stages. `EditOptions::max_operations` selects a stricter finite
operation budget.

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
