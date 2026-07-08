# litchi-formula

Mathematical formula parsing and conversion between OMML, MTEF, and LaTeX.

## Overview

`litchi-formula` provides parsers and converters for the math formats
encountered in Office documents: **OMML** (Office Math Markup Language,
used in `.docx`/`.pptx`), **MTEF** (the binary MathType Equation Format
embedded in legacy `.doc`/`.ppt` files), and **LaTeX** as the canonical
output. Internally it uses an arena-allocated AST so OMML and MTEF input
share the same conversion path. It is part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.

## Usage

```toml
[dependencies]
litchi-formula = "0.0.1"
```

```rust
use litchi_formula::{omml_to_latex, mtef_to_latex, FormulaError};

fn convert_examples(mtef_bytes: &[u8]) -> Result<(String, String), FormulaError> {
    let from_omml = omml_to_latex("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
    let from_mtef = mtef_to_latex(mtef_bytes)?;
    Ok((from_omml, from_mtef))
}
```

## Features

- OMML parser (`OmmlParser`) for modern Office math markup
- MTEF parser (`MtefParser`) for legacy MathType binary streams
- LaTeX writer (`LatexConverter`) over a shared arena-allocated AST
- One-shot helpers: `omml_to_latex`, `mtef_to_latex`
- Unified `FormulaError` covering all parser/converter failures

## License

Licensed under the Apache License, Version 2.0. Part of the
[Litchi](https://github.com/DevExzh/litchi) workspace.
