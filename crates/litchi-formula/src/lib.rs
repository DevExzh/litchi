#![allow(missing_docs)]
// quick-xml's checked attribute iteration is quadratic on hostile tags; read
// attributes through `omml::xml_attributes` (records 0764 and 0770, workspace
// `clippy.toml`).
#![cfg_attr(not(test), deny(clippy::disallowed_methods))]

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
// let parser = OmmlParser::new(formula.arena(), omml::Limits::DEFAULT);
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
/// - <http://rtf2latex2e.sourceforge.net/MTEF5.html>
/// - rtf2latex2e source code
pub mod mtef;
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
/// Reference: <https://devblogs.microsoft.com/math-in-office/officemath/>
pub mod omml;

// Re-export public API
pub use ast::{
    AccentType, Fence, Formula, FormulaBuilder, LargeOperator, MathNode, MatrixFence, Operator,
    SpaceType, StyleType, Symbol,
};
pub use latex::{DEFAULT_MAX_DEPTH, LatexConverter, LatexError, LatexParseError, LatexParser};
pub use mtef::{MtefError, MtefParser, MtefWriteError, MtefWriteOptions, MtefWriter};
pub use omml::{OmmlError, OmmlParser, OmmlWriter};

/// Conversion error that wraps all possible formula errors
#[derive(Debug)]
#[non_exhaustive]
pub enum FormulaError {
    Omml(OmmlError),
    Latex(LatexError),
    LatexParse(LatexParseError),
    Mtef(MtefError),
    MtefWrite(MtefWriteError),
}

impl std::fmt::Display for FormulaError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FormulaError::Omml(e) => write!(f, "OMML error: {}", e),
            FormulaError::Latex(e) => write!(f, "LaTeX error: {}", e),
            FormulaError::LatexParse(e) => write!(f, "LaTeX parse error: {}", e),
            FormulaError::Mtef(e) => write!(f, "MTEF error: {}", e),
            FormulaError::MtefWrite(e) => write!(f, "MTEF write error: {}", e),
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

impl From<LatexParseError> for FormulaError {
    fn from(e: LatexParseError) -> Self {
        FormulaError::LatexParse(e)
    }
}

impl From<MtefWriteError> for FormulaError {
    fn from(e: MtefWriteError) -> Self {
        FormulaError::MtefWrite(e)
    }
}

/// High-level conversion functions
/// Convert OMML to LaTeX
///
/// # Example
/// ```ignore
/// let latex = omml_to_latex(
///     "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
///     &omml::Limits::DEFAULT,
/// )?;
/// println!("LaTeX: {}", latex);
/// ```
pub fn omml_to_latex(omml: &str, limits: &omml::Limits) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let parser = OmmlParser::new(formula.arena(), *limits);
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
/// let latex = mtef_to_latex(mtef_data, &mtef::Limits::DEFAULT)?;
/// println!("LaTeX: {}", latex);
/// ```
pub fn mtef_to_latex(mtef_data: &[u8], limits: &mtef::Limits) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), mtef_data, *limits);
    let nodes = parser.parse()?;

    let mut formula = Formula::new();
    formula.set_root(nodes);

    let mut converter = LatexConverter::new();
    Ok(converter.convert(&formula)?.to_string())
}

/// Convert OMML to MTEF
///
/// Parses the `m:oMath` fragment and serializes the resulting AST as an MTEF 5
/// stream carrying the 28-byte `Equation.3` OLE header, ready to store as an
/// `Equation Native` stream. Use [`MtefWriter`] directly for a bare payload.
///
/// # Example
/// ```ignore
/// let mtef = omml_to_mtef(
///     "<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>",
///     &omml::Limits::DEFAULT,
///     &mtef::Limits::DEFAULT,
/// )?;
/// ```
pub fn omml_to_mtef(
    omml: &str,
    input_limits: &omml::Limits,
    output_limits: &mtef::Limits,
) -> Result<Vec<u8>, FormulaError> {
    let formula = Formula::new();
    let parser = OmmlParser::new(formula.arena(), *input_limits);
    let nodes = parser.parse(omml)?;

    let writer = MtefWriter::new(*output_limits);
    Ok(writer.write_nodes(&nodes)?)
}

/// Convert LaTeX to OMML
///
/// # Example
/// ```ignore
/// let omml = latex_to_omml(r"x^2 + \sqrt{y}", &omml::Limits::DEFAULT)?;
/// ```
pub fn latex_to_omml(latex: &str, limits: &omml::Limits) -> Result<String, FormulaError> {
    let parser = LatexParser::new();
    let nodes = parser.parse(latex)?;

    let writer = OmmlWriter::new(*limits);
    Ok(writer.write_nodes(&nodes)?.to_string())
}

/// Convert LaTeX to MTEF
///
/// Produces an MTEF 5 stream with the `Equation.3` OLE header, as
/// [`omml_to_mtef`] does.
///
/// # Example
/// ```ignore
/// let mtef = latex_to_mtef(r"\alpha + \beta", &mtef::Limits::DEFAULT)?;
/// ```
pub fn latex_to_mtef(latex: &str, limits: &mtef::Limits) -> Result<Vec<u8>, FormulaError> {
    let parser = LatexParser::new();
    let nodes = parser.parse(latex)?;

    let writer = MtefWriter::new(*limits);
    Ok(writer.write_nodes(&nodes)?)
}

/// Convert MTEF binary data to OMML
///
/// Parses the MathType equation and serializes the resulting AST as an
/// `m:oMath` fragment suitable for embedding in OOXML documents.
///
/// # Example
/// ```ignore
/// let omml = mtef_to_omml(
///     mtef_data,
///     &mtef::Limits::DEFAULT,
///     &omml::Limits::DEFAULT,
/// )?;
/// println!("OMML: {}", omml);
/// ```
pub fn mtef_to_omml(
    mtef_data: &[u8],
    input_limits: &mtef::Limits,
    output_limits: &omml::Limits,
) -> Result<String, FormulaError> {
    let formula = Formula::new();
    let mut parser = MtefParser::new(formula.arena(), mtef_data, *input_limits);
    let nodes = parser.parse()?;

    let writer = OmmlWriter::new(*output_limits);
    Ok(writer.write_nodes(&nodes)?.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_omml_to_latex() {
        let omml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let result = omml_to_latex(omml, &omml::Limits::DEFAULT);
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
        let omml = mtef_to_omml(
            &minimal_mtef5(),
            &mtef::Limits::DEFAULT,
            &omml::Limits::DEFAULT,
        )
        .expect("minimal MTEF 5 converts");
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
        let omml = mtef_to_omml(&[0u8; 4], &mtef::Limits::DEFAULT, &omml::Limits::DEFAULT)
            .expect("short input is not fatal");
        assert!(omml.starts_with("<m:oMath "));
        assert!(omml.ends_with("</m:oMath>"));
    }
}

#[cfg(test)]
mod conversion_tests {
    use super::*;
    use std::borrow::Cow;

    /// `omml_to_mtef` used to be an `unimplemented!()` panic in the public API.
    #[test]
    fn omml_to_mtef_produces_a_stream_the_reader_accepts() {
        let omml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let mtef = omml_to_mtef(omml, &omml::Limits::DEFAULT, &mtef::Limits::DEFAULT)
            .expect("OMML converts to MTEF");
        assert!(!mtef.is_empty(), "expected a non-empty MTEF stream");

        let latex =
            mtef_to_latex(&mtef, &mtef::Limits::DEFAULT).expect("the writer's output parses back");
        assert!(latex.contains('x'), "round trip lost the operand: {latex}");
    }

    #[test]
    fn latex_to_omml_emits_a_namespaced_fragment() {
        let omml = latex_to_omml(r"x^2", &omml::Limits::DEFAULT).expect("LaTeX converts to OMML");
        assert!(omml.starts_with("<m:oMath "), "got {omml}");
        assert!(omml.contains("officeDocument/2006/math"), "got {omml}");
        assert!(omml.ends_with("</m:oMath>"), "unclosed fragment: {omml}");
        assert!(omml.contains('2'), "lost the exponent: {omml}");
    }

    #[test]
    fn latex_to_mtef_round_trips_through_the_reader() {
        let mtef = latex_to_mtef(r"a+b", &mtef::Limits::DEFAULT).expect("LaTeX converts to MTEF");
        let latex =
            mtef_to_latex(&mtef, &mtef::Limits::DEFAULT).expect("the writer's output parses back");
        for expected in ['a', '+', 'b'] {
            assert!(
                latex.contains(expected),
                "round trip lost {expected}: {latex}"
            );
        }
    }

    /// A LaTeX fragment survives a full OMML round trip.
    #[test]
    fn latex_survives_a_round_trip_through_omml() {
        let omml = latex_to_omml(r"\frac{a}{b}", &omml::Limits::DEFAULT).expect("converts to OMML");
        let latex = omml_to_latex(&omml, &omml::Limits::DEFAULT).expect("converts back to LaTeX");
        assert!(latex.contains("\\frac"), "lost the fraction: {latex}");
        assert!(
            latex.contains('a') && latex.contains('b'),
            "lost operands: {latex}"
        );
    }

    /// Malformed LaTeX must be reported, never panic.
    #[test]
    fn malformed_latex_is_an_error_not_a_panic() {
        assert!(latex_to_omml(r"\frac{a", &omml::Limits::DEFAULT).is_err());
        assert!(latex_to_mtef(r"{{{", &mtef::Limits::DEFAULT).is_err());
    }

    #[test]
    fn raw_boundaries_reject_input_and_output_before_codec_work() {
        let omml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let input_limits = omml::Limits::new(omml.len(), 32, 256, 256, 1024, 4096)
            .expect("finite OMML boundary policy");
        assert!(omml_to_latex(omml, &input_limits).is_ok());

        let too_small_input = omml::Limits::new(omml.len() - 1, 32, 256, 256, 1024, 4096)
            .expect("finite OMML boundary policy");
        assert!(matches!(
            omml_to_latex(omml, &too_small_input),
            Err(FormulaError::Omml(OmmlError::LimitExceeded {
                kind: omml::LimitKind::InputBytes,
                ..
            }))
        ));

        let tiny_mtef_input =
            mtef::Limits::new(4, 32, 256, 256, 1024, 4096).expect("finite MTEF input policy");
        assert!(matches!(
            mtef_to_latex(&[0; 5], &tiny_mtef_input),
            Err(FormulaError::Mtef(MtefError::LimitExceeded {
                kind: mtef::LimitKind::InputBytes,
                ..
            }))
        ));

        let tiny_omml_output =
            omml::Limits::new(4096, 32, 256, 256, 1024, 1).expect("finite OMML output policy");
        assert!(matches!(
            latex_to_omml("x", &tiny_omml_output),
            Err(FormulaError::Omml(OmmlError::LimitExceeded {
                kind: omml::LimitKind::OutputBytes,
                ..
            }))
        ));

        let tiny_mtef_output =
            mtef::Limits::new(4096, 32, 256, 256, 1024, 1).expect("finite MTEF output policy");
        assert!(matches!(
            latex_to_mtef("x", &tiny_mtef_output),
            Err(FormulaError::MtefWrite(MtefWriteError::LimitExceeded {
                kind: mtef::LimitKind::OutputBytes,
                ..
            }))
        ));
    }

    #[test]
    fn raw_parser_depth_is_caller_selected() {
        let xml = r#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        let limits =
            omml::Limits::new(4096, 1, 256, 256, 1024, 4096).expect("finite OMML depth policy");
        assert!(matches!(
            omml_to_latex(xml, &limits),
            Err(FormulaError::Omml(OmmlError::LimitExceeded {
                kind: omml::LimitKind::Depth,
                ..
            }))
        ));
    }

    #[test]
    fn raw_writers_charge_work_items_and_strings() {
        let nodes = [MathNode::Text(Cow::Borrowed("xx"))];
        let mtef_string_limits =
            mtef::Limits::new(4096, 32, 256, 256, 1, 4096).expect("finite MTEF string policy");
        assert!(matches!(
            MtefWriter::new(mtef_string_limits).write_nodes(&nodes),
            Err(MtefWriteError::LimitExceeded {
                kind: mtef::LimitKind::StringBytes,
                ..
            })
        ));

        let omml_string_limits =
            omml::Limits::new(4096, 32, 256, 256, 1, 4096).expect("finite OMML string policy");
        assert!(matches!(
            OmmlWriter::new(omml_string_limits).write_nodes(&nodes),
            Err(OmmlError::LimitExceeded {
                kind: omml::LimitKind::StringBytes,
                ..
            })
        ));

        let mtef_item_limits =
            mtef::Limits::new(4096, 32, 256, 1, 1024, 4096).expect("finite MTEF item policy");
        let two_nodes = [
            MathNode::Text(Cow::Borrowed("x")),
            MathNode::Text(Cow::Borrowed("y")),
        ];
        assert!(matches!(
            MtefWriter::new(mtef_item_limits).write_nodes(&two_nodes),
            Err(MtefWriteError::LimitExceeded {
                kind: mtef::LimitKind::Items,
                ..
            })
        ));

        let omml_item_limits =
            omml::Limits::new(4096, 32, 256, 1, 1024, 4096).expect("finite OMML item policy");
        assert!(matches!(
            OmmlWriter::new(omml_item_limits).write_nodes(&two_nodes),
            Err(OmmlError::LimitExceeded {
                kind: omml::LimitKind::Items,
                ..
            })
        ));

        let mtef_work_limits =
            mtef::Limits::new(4096, 32, 1, 256, 1024, 4096).expect("finite MTEF work policy");
        assert!(matches!(
            MtefWriter::new(mtef_work_limits).write_nodes(&[MathNode::Text(Cow::Borrowed("x"))]),
            Err(MtefWriteError::LimitExceeded {
                kind: mtef::LimitKind::Work,
                ..
            })
        ));

        let omml_work_limits =
            omml::Limits::new(4096, 32, 1, 256, 1024, 4096).expect("finite OMML work policy");
        assert!(matches!(
            OmmlWriter::new(omml_work_limits).write_nodes(&[MathNode::Text(Cow::Borrowed("x"))]),
            Err(OmmlError::LimitExceeded {
                kind: omml::LimitKind::Work,
                ..
            })
        ));
    }
}
