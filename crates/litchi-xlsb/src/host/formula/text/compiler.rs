//! Public formula-text compiler facade implementation.

use super::super::{Error, MAX_CELL_FORMULA_BYTES, ParsedFormula, Result};
use super::model::{CompilationContext, FormulaEncoding};

/// Compiles standards-defined Excel formula text to XLSB RPN tokens.
///
/// The compiler supports literals, A1 references and ranges, parentheses,
/// arithmetic/comparison/concatenation operators, percent, and the built-in
/// non-macro built-in functions from [MS-XLSB]'s `Ftab` table, and typed array
/// constants. Unsupported constructs return an error; they are never replaced
/// by a cached value.
pub struct Compiler<'a> {
    pub(super) input: &'a str,
    pub(super) offset: usize,
    pub(super) context: Option<&'a CompilationContext<'a>>,
}

impl<'a> Compiler<'a> {
    pub fn compile(formula: &'a str) -> Result<ParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, None)
    }

    pub(crate) fn compile_with_context(
        formula: &'a str,
        context: &'a CompilationContext<'a>,
    ) -> Result<ParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, Some(context))
    }

    /// Compile a shared formula, encoding relative A1 references as
    /// `PtgRefN`/`PtgAreaN` offsets from the first cell in the shared range.
    pub fn compile_shared(formula: &'a str, base_row: u32, base_col: u32) -> Result<ParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, None)
    }

    pub(crate) fn compile_shared_with_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: &'a CompilationContext<'a>,
    ) -> Result<ParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, Some(context))
    }

    fn compile_shared_with_optional_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: Option<&'a CompilationContext<'a>>,
    ) -> Result<ParsedFormula> {
        if base_row >= 1_048_576 || base_col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "shared formula base ({base_row}, {base_col})"
            )));
        }
        Self::compile_with_encoding(
            formula,
            FormulaEncoding::Shared { base_row, base_col },
            context,
        )
    }

    fn compile_with_encoding(
        formula: &'a str,
        encoding: FormulaEncoding,
        context: Option<&'a CompilationContext<'a>>,
    ) -> Result<ParsedFormula> {
        let input = formula.strip_prefix('=').unwrap_or(formula).trim();
        if input.is_empty() {
            return Err(Error::InvalidFormula(
                "formula expression is empty".to_string(),
            ));
        }
        let mut compiler = Self {
            input,
            offset: 0,
            context,
        };
        let expression = compiler.parse_comparison()?;
        compiler.skip_spaces();
        if compiler.offset != compiler.input.len() {
            return Err(compiler.error("unexpected trailing input"));
        }

        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        Self::emit(&expression, &mut rgce, &mut rgcb, encoding)?;
        if rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "compiled formula is {} bytes; maximum is {MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(ParsedFormula { rgce, rgcb })
    }
}
