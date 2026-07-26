#![allow(missing_docs)]
// Formula Module - Mathematical Formula Parsing and Conversion
//
// This module provides comprehensive support for parsing and converting
// mathematical formulas between different formats:
//
// - **OMML** (Office Math Markup Language): XML-based format used in modern Office files
// - **LaTeX**: Standard mathematical typesetting format
// - **MTEF** (MathType Equation Format): Binary format used in legacy OLE files
//
// The module uses a common Abstract Syntax Tree (AST) representation to enable
// efficient conversion between formats.
//
// # Example
//
// ```ignore
// use litchi::formula::{Formula, OmmlParser, LatexConverter};
//
// // Parse OMML
// let formula = Formula::new();
// let parser = OmmlParser::new(formula.arena());
// let nodes = parser.parse("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
//
// // Convert to LaTeX
// let mut formula = Formula::new();
// formula.set_root(nodes);
// let mut converter = LatexConverter::new();
// let latex = converter.convert(&formula)?;
// ```

/// Abstract Syntax Tree for Mathematical Formulas
///
/// This module defines a comprehensive AST for representing mathematical formulas
/// that can be parsed from OMML, LaTeX, and MTEF formats and converted between them.
///
/// The design is inspired by the plurimath Ruby project but adapted for Rust's
/// type system and performance characteristics.
pub mod ast;
/// LaTeX Converter
///
/// This module converts our formula AST to LaTeX format.
/// LaTeX is a widely-used typesetting system for mathematical formulas.
pub mod latex;
/// MTEF (MathType Equation Format) Parser
///
/// This module parses the binary MathType Equation Format (MTEF) used in
/// legacy OLE documents (.doc, .ppt, etc.) into our formula AST.
///
/// MTEF is a private data stream format developed by Design Science for
/// storing mathematical equations.
///
/// References:
/// - http://rtf2latex2e.sourceforge.net/MTEF5.html
/// - rtf2latex2e source code
mod mtef;
/// OMML (Office Math Markup Language) Parser
///
/// This module parses Microsoft Office Math Markup Language (OMML) into our AST.
/// OMML is used in modern Office documents (.docx, .pptx, etc.) to represent
/// mathematical formulas.
///
/// This implementation provides comprehensive OMML parsing with:
/// - High-performance streaming XML parsing
/// - Modular element handlers for different OMML constructs
/// - Comprehensive attribute parsing
/// - Memory-efficient arena-based allocation
/// - Support for all OMML elements and properties
///
/// Reference: https://devblogs.microsoft.com/math-in-office/officemath/
mod omml;

// Re-export public API
pub use ast::{
    AccentType, Fence, Formula, FormulaBuilder, LargeOperator, MathNode, MatrixFence, Operator,
    SpaceType, StyleType, Symbol,
};
pub use latex::{LatexConverter, LatexError};
pub use mtef::{MtefError, MtefParser};
pub use omml::{OmmlError, OmmlParser, OmmlWriter};

/// Conversion error that wraps all possible formula errors
#[derive(Debug)]
#[non_exhaustive]
pub enum FormulaError {
    Omml(OmmlError),
    Latex(LatexError),
    Mtef(MtefError),
}

impl std::fmt::Display for FormulaError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FormulaError::Omml(e) => write!(f, "OMML error: {}", e),
            FormulaError::Latex(e) => write!(f, "LaTeX error: {}", e),
            FormulaError::Mtef(e) => write!(f, "MTEF error: {}", e),
        }
    }
}

impl std::error::Error for FormulaError {}

impl From<OmmlError> for FormulaError {
    fn from(e: OmmlError) -> Self {
        FormulaError::Omml(e)
    }
}

impl From<LatexError> for FormulaError {
    fn from(e: LatexError) -> Self {
        FormulaError::Latex(e)
    }
}

impl From<MtefError> for FormulaError {
    fn from(e: MtefError) -> Self {
        FormulaError::Mtef(e)
    }
}

/// High-level conversion functions
/// Convert OMML to LaTeX
///
/// # Example
/// ```ignore
/// let latex = omml_to_latex("<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>")?;
/// println!("LaTeX: {}", latex);
/// ```
pub fn omml_to_latex(omml: &str) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let parser = OmmlParser::new(formula.arena());
    let nodes = parser.parse(omml)?;

    let mut formula = Formula::new();
    formula.set_root(nodes);

    let mut converter = LatexConverter::new();
    Ok(converter.convert(&formula)?.to_string())
}

/// Convert MTEF binary data to LaTeX
///
/// # Example
/// ```ignore
/// let latex = mtef_to_latex(mtef_data)?;
/// println!("LaTeX: {}", latex);
/// ```
pub fn mtef_to_latex(mtef_data: &[u8]) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), mtef_data);
    let nodes = parser.parse()?;

    let mut formula = Formula::new();
    formula.set_root(nodes);

    let mut converter = LatexConverter::new();
    Ok(converter.convert(&formula)?.to_string())
}

/// Convert OMML to MTEF (not yet implemented)
///
/// This function is planned for future implementation.
pub fn omml_to_mtef(_omml: &str) -> Result<Vec<u8>, FormulaError> {
    unimplemented!("OMML to MTEF conversion is not yet implemented")
}

/// Convert MTEF binary data to OMML
///
/// Parses the MathType equation and serializes the resulting AST as an
/// `m:oMath` fragment suitable for embedding in OOXML documents.
///
/// # Example
/// ```ignore
/// let omml = mtef_to_omml(mtef_data)?;
/// println!("OMML: {}", omml);
/// ```
pub fn mtef_to_omml(mtef_data: &[u8]) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), mtef_data);
    let nodes = parser.parse()?;

    let mut writer = OmmlWriter::new();
    Ok(writer.write_nodes(&nodes)?.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_omml_to_latex() {
        let omml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let result = omml_to_latex(omml);
        assert!(result.is_ok());
    }

    #[test]
    fn test_formula_creation() {
        let formula = Formula::new();
        assert!(formula.root().is_empty());
        assert!(formula.display_style());
    }

    /// A minimal MTEF 5 stream: OLE equation header plus a SIZE and END record.
    fn minimal_mtef5() -> Vec<u8> {
        let mut data = vec![0u8; 28];
        data[0] = 0x1c; // cb_hdr
        data[2] = 0x00; // version
        data[4] = 0x03; // cf_id
        data[5] = 0x00;
        data[6] = 0x00;
        data[7] = 0x00;
        data[8] = 0x02;
        data.extend_from_slice(&[
            0x05, // mtef_version = 5
            0x01, // platform = Windows
            0x01, // product = MathType
            0x01, // version
            0x00, // version_sub
            0x00, // application_key (empty, null-terminated)
            0x00, // inline
            0x09, // SIZE
            0x00, // END
        ]);
        data
    }

    #[test]
    fn mtef_to_omml_emits_a_namespaced_omath_fragment() {
        let omml = mtef_to_omml(&minimal_mtef5()).expect("minimal MTEF 5 converts");
        assert!(
            omml.starts_with("<m:oMath "),
            "expected an m:oMath root, got {omml}"
        );
        assert!(
            omml.contains("officeDocument/2006/math"),
            "expected the OMML namespace declaration, got {omml}"
        );
        assert!(omml.ends_with("</m:oMath>"), "unclosed fragment: {omml}");
    }

    #[test]
    fn mtef_to_omml_rejects_no_data_without_panicking() {
        // Too short for the OLE equation header: the parser yields no nodes
        // rather than panicking, so the conversion still produces a fragment.
        let omml = mtef_to_omml(&[0u8; 4]).expect("short input is not fatal");
        assert!(omml.starts_with("<m:oMath "));
        assert!(omml.ends_with("</m:oMath>"));
    }
}
