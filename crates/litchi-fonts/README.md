# litchi-fonts

Portable font metadata, Office embedding, OOXML font obfuscation, and optional
native discovery or glyph subsetting for the Litchi office-formats library.

## Overview

`litchi-fonts` provides the format-independent font-handling layer used when
generating Office documents that need to embed or subset typefaces. Its default
build is portable: it provides the model, explicit resolver-based preparation,
EOT publication, and OOXML obfuscation without native font packages. Optional
backends add `font-kit` system discovery and `allsorts` subsetting. It is
consumed by the document-family writers (`docx`/`pptx` font embedding) inside the
[Litchi](https://github.com/DevExzh/litchi) workspace.

The API is organized by ownership:

- `embedding::Resolver` and `prepare_with` create owned, validated font
  programs from an application-provided source under a required
  `PreparationLimits` policy; `embedding::prepare` does the same through the
  native resolver when `automatic` is enabled. `embedding::powerpoint::encode`
  publishes the EOT wrapper under a required `powerpoint::Limits` policy and
  explicit intent.
- `discovery::Loader` resolves system faces with a 64 MiB program limit when
  the `discovery` feature is selected.
- `subset::mapping` maps Unicode requests to glyph IDs; `Allsorts` reduces
  OpenType programs when the `subset` feature is selected.
- `obfuscation::apply` and `obfuscation::remove` implement the OOXML transform.

## System Dependencies

On Ubuntu/Debian, install the FreeType and Fontconfig development packages:

```bash
sudo apt install pkg-config libfreetype6-dev libfontconfig1-dev
```

## Usage

```toml
[dependencies]
litchi-fonts = { version = "0.0.1", features = ["automatic"] }
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

- `default`: portable model, explicit `embedding::Resolver` preparation, EOT,
  and OOXML obfuscation. No native font dependencies.
- `discovery`: system `Loader` backed by `font-kit`; requires the platform
  font packages listed above.
- `subset`: `allsorts` glyph mapping and OpenType subsetting.
- `automatic`: combines `discovery` and `subset`, enabling
  `litchi_fonts::prepare` for the conventional system-font workflow. Hosts
  must pass an explicit `PreparationLimits` value.

Use `embedding::Mode::{Full, Subset}` to choose complete or subset font
programs. `Subset` requires the `subset` feature; explicit source integrations
can otherwise use `prepare_with` with `Mode::Full`.

`embedding::PreparationLimits::max_subset_working_bytes` is a separate,
finite budget for temporary memory used by the subsetter. It is independent
of `max_program_bytes` and `max_total_program_bytes`: the latter two remain
the limits on each retained program and on the aggregate retained output.
The current standard default is 1 GiB, pending corpus and bounded-build
validation; callers may select a narrower or larger value through their owner
options. Subset admission charges the source and requested glyph-ID staging
before mapping, then passes the selected budget through the checked subsetter
adapter before parser, index, or writer allocations.

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
