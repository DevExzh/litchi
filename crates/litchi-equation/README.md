# litchi-equation

Math equation conversion for office documents.

## Overview

`litchi-equation` converts the equation notations Office files store to and
from LaTeX: **OMML** (Office Math Markup Language, in `.docx`/`.pptx`/`.xlsx`),
binary **MTEF** (MathType and Equation Editor equations in the
`Equation Native` stream of OLE objects in `.doc`/`.ppt`/`.xls`) and
**MathML** (OpenDocument formula objects). Parsing and writing are done by
the [`mathema`](https://github.com/DevExzh/mathema) converter; this crate
adds a typed `Error`, caller-selected `Limits` and host-shaped conversions.
It is part of the [Litchi](https://github.com/DevExzh/litchi) workspace.

## Usage

```toml
[dependencies]
litchi-equation = "0.0.1"
```

```rust
use litchi_equation::{Limits, OmmlHost};

fn convert(equation_native: &[u8]) -> litchi_equation::Result<(String, String, litchi_equation::Omml)> {
    let from_omml = litchi_equation::omml_to_latex(
        "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
        &Limits::DEFAULT,
    )?;
    let from_mtef = litchi_equation::mtef_to_latex(equation_native, &Limits::DEFAULT)?;
    let omml = litchi_equation::latex_to_omml(r"\frac{a}{b}", OmmlHost::WordprocessingMl, &Limits::DEFAULT)?;
    Ok((from_omml, from_mtef, omml))
}
```

## Features

- `omml_to_latex`, `mathml_to_latex`: OMML and MathML to LaTeX
- `read_mtef`, `mtef_to_latex`: `Equation Native` streams and bare MTEF 3/5
  data to LaTeX, with the MTEF version and generating product (`MtefInfo`)
- `latex_to_omml`: LaTeX to `m:oMath` for Word (`w:rPr`) or DrawingML
  (`a:rPr`) hosts, reporting formatting the host cannot express
- `Limits`: input bytes, parsed nodes and nesting depth for every conversion
- One typed `Error` covering malformed input, unsupported constructs and
  exceeded limits

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
