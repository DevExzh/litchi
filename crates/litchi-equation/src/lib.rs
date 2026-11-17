//! Math equations in office documents.
//!
//! Office files store equations in three notations: **OMML** (Office Math
//! Markup Language) in `.docx`, `.pptx` and `.xlsx`, binary **MTEF** (the
//! MathType and Equation Editor format) in the `Equation Native` stream of
//! OLE objects in `.doc`, `.ppt` and `.xls`, and **MathML** in OpenDocument
//! formula objects. This crate converts them to and from **LaTeX** through
//! the [`mathema`] converter, and adds what Litchi needs on top of it: one
//! typed [`Error`], caller-selected [`Limits`], and conversions shaped for
//! the hosts that embed equations.
//!
//! ```
//! use litchi_equation::{Limits, OmmlHost};
//!
//! let latex = litchi_equation::omml_to_latex(
//!     "<m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>\
//!      <m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>",
//!     &Limits::DEFAULT,
//! )?;
//! assert_eq!(latex, r"\frac{a}{b}");
//!
//! let omml = litchi_equation::latex_to_omml("x^2", OmmlHost::WordprocessingMl, &Limits::DEFAULT)?;
//! assert!(omml.xml().starts_with("<m:oMath"));
//! assert!(omml.losses().is_empty());
//! # Ok::<(), litchi_equation::Error>(())
//! ```

use std::fmt;

#[cfg(feature = "render")]
mod render;
#[cfg(feature = "render")]
pub use render::{Png, PngOptions, latex_to_png, omml_size_pt, omml_to_png};

mod mtef_write;
#[cfg(feature = "ole")]
mod ole_object;
#[cfg(feature = "render")]
pub use mtef_write::mtef_to_png;
pub use mtef_write::{
    EquationNative, MtefTarget, latex_to_equation_native, omml_to_equation_native,
};
#[cfg(feature = "ole")]
pub use ole_object::{
    MAX_EQUATION_NATIVE_BYTES, MAX_EQUATION_OBJECT_BYTES, MAX_EQUATION_OBJECT_DEPTH,
    MAX_EQUATION_OBJECT_ENTRIES, equation_object, equation_object_native,
    replace_equation_object_native,
};

/// Result type of every conversion in this crate.
pub type Result<T> = std::result::Result<T, Error>;

/// An equation notation, as named in an [`Error`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    /// Office Math Markup Language.
    Omml,
    /// MathType Equation Format (binary).
    Mtef,
    /// MathML.
    MathMl,
    /// LaTeX math markup.
    Latex,
}

impl fmt::Display for Format {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Omml => "OMML",
            Self::Mtef => "MTEF",
            Self::MathMl => "MathML",
            Self::Latex => "LaTeX",
        })
    }
}

/// Why an equation could not be converted.
#[derive(Debug, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The input is not a well-formed equation in `format`.
    #[error("malformed {format} equation: {message}")]
    Malformed {
        /// The notation of the rejected input.
        format: Format,
        /// What is wrong with it.
        message: String,
    },
    /// The equation uses a construct the target notation cannot represent.
    #[error("unsupported equation construct: {0}")]
    Unsupported(String),
    /// The input is larger than [`Limits::max_input_bytes`].
    #[error("equation input of {size} bytes exceeds the {limit}-byte limit")]
    InputTooLarge {
        /// Size of the rejected input, in bytes.
        size: usize,
        /// The limit it exceeded.
        limit: usize,
    },
    /// Parsing needed more of a bounded resource than [`Limits`] allows.
    #[error("equation {kind} limit of {limit} exceeded")]
    LimitExceeded {
        /// Which limit was exceeded.
        kind: LimitKind,
        /// The value of that limit.
        limit: usize,
    },
    /// A [`Limits`] value is out of range.
    #[error("invalid equation limit: {0}")]
    InvalidLimit(&'static str),
    /// The equation could not be rendered as a picture, for example because
    /// the picture would exceed [`PngOptions::max_pixels`].
    #[error("equation rendering failed: {0}")]
    Render(String),
    /// The compound file of an OLE equation object is invalid, lacks its
    /// `Equation Native` stream, or exceeds a size limit.
    #[error("invalid OLE equation object: {0}")]
    Object(String),
    /// The converter failed for a reason unrelated to the input.
    #[error("equation conversion failed: {0}")]
    Internal(String),
}

/// A bounded resource of one conversion, as named in
/// [`Error::LimitExceeded`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The number of parsed nodes; see [`Limits::max_nodes`].
    Nodes,
    /// The nesting depth; see [`Limits::max_depth`].
    Depth,
    /// A fixed structural bound of the converter (the children of one node,
    /// or the size of its node store); not configurable.
    Structure,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Nodes => "node count",
            Self::Depth => "nesting depth",
            Self::Structure => "structure size",
        })
    }
}

impl LimitKind {
    fn from_mathema(kind: mathema::LimitKind) -> Self {
        match kind {
            mathema::LimitKind::Nodes => Self::Nodes,
            mathema::LimitKind::Depth => Self::Depth,
            _ => Self::Structure,
        }
    }
}

impl Error {
    fn from_mathema(error: mathema::Error, format: Format) -> Self {
        use mathema::Error as E;
        let malformed = |format, message| Self::Malformed { format, message };
        match error {
            E::LatexParse { message, .. } => malformed(Format::Latex, message),
            E::MathmlParse { message } => malformed(Format::MathMl, message),
            E::OmmlParse { message } => malformed(Format::Omml, message),
            E::MtefParse { message } => malformed(Format::Mtef, message),
            E::Xml { message } | E::UnicodeMathParse { message } => malformed(format, message),
            E::Unsupported(message) => Self::Unsupported(message),
            E::Render { message } => Self::Render(message),
            // The writer refused its own output; a converter defect, not
            // a property of the input.
            E::MtefWrite { message } => Self::Internal(format!(
                "the MTEF writer produced unreadable data: {message}"
            )),
            E::InputTooLarge { size, limit } => Self::InputTooLarge { size, limit },
            E::LimitExceeded { kind, limit } => Self::LimitExceeded {
                kind: LimitKind::from_mathema(kind),
                limit,
            },
            other => Self::Internal(other.to_string()),
        }
    }
}

/// Bounds on the work one equation conversion may do.
///
/// Equations are kilobytes at most; the bounds keep a hostile document from
/// exhausting memory or stack through one embedded equation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Limits {
    max_input_bytes: usize,
    max_nodes: usize,
    max_depth: usize,
}

impl Limits {
    /// The deepest nesting any conversion accepts; [`Limits::new`] rejects
    /// larger depths.
    pub const MAX_DEPTH: usize = mathema::error::DEFAULT_MAX_DEPTH;

    /// 4 MiB of input, one million parsed nodes and [`Limits::MAX_DEPTH`]
    /// levels of nesting.
    pub const DEFAULT: Self = Self {
        max_input_bytes: mathema::error::DEFAULT_MAX_INPUT_BYTES,
        max_nodes: mathema::error::DEFAULT_MAX_NODES,
        max_depth: Self::MAX_DEPTH,
    };

    /// Limits with the given input size in bytes, parsed node count and
    /// nesting depth.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when a value is zero or the depth
    /// exceeds [`Limits::MAX_DEPTH`].
    pub const fn new(max_input_bytes: usize, max_nodes: usize, max_depth: usize) -> Result<Self> {
        if max_input_bytes == 0 {
            return Err(Error::InvalidLimit("the input byte limit must be positive"));
        }
        if max_nodes == 0 {
            return Err(Error::InvalidLimit("the node limit must be positive"));
        }
        if max_depth == 0 || max_depth > Self::MAX_DEPTH {
            return Err(Error::InvalidLimit(
                "the depth limit must be between 1 and Limits::MAX_DEPTH",
            ));
        }
        Ok(Self {
            max_input_bytes,
            max_nodes,
            max_depth,
        })
    }

    /// The largest accepted input, in bytes.
    #[must_use]
    pub const fn max_input_bytes(&self) -> usize {
        self.max_input_bytes
    }

    /// The most nodes one parse may create.
    #[must_use]
    pub const fn max_nodes(&self) -> usize {
        self.max_nodes
    }

    /// The deepest accepted nesting.
    #[must_use]
    pub const fn max_depth(&self) -> usize {
        self.max_depth
    }

    fn config(&self) -> mathema::Config {
        mathema::Config::default()
            .with_max_input_bytes(Some(self.max_input_bytes))
            .with_max_nodes(Some(self.max_nodes))
            .with_max_depth(self.max_depth)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// The document part an OMML equation is written into, which decides how
/// run formatting (colour, bold or italic normal text) is expressed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum OmmlHost {
    /// Pure OMML without host run properties; run formatting is dropped.
    #[default]
    Bare,
    /// A WordprocessingML (Word) paragraph: `w:rPr` run properties.
    WordprocessingMl,
    /// A DrawingML (PowerPoint, Excel) text body: `a:rPr` run properties.
    DrawingMl,
}

impl OmmlHost {
    fn run_properties(self) -> mathema::OmmlRunProperties {
        match self {
            Self::Bare => mathema::OmmlRunProperties::None,
            Self::WordprocessingMl => mathema::OmmlRunProperties::WordprocessingMl,
            Self::DrawingMl => mathema::OmmlRunProperties::DrawingMl,
        }
    }
}

/// The product that wrote an MTEF equation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MtefProduct {
    /// MathType.
    MathType,
    /// Equation Editor (Microsoft Equation 3.0).
    EquationEditor,
    /// Another product code.
    Other,
}

/// The `EQNOLEFILEHDR` that precedes MTEF data in an OLE `Equation Native`
/// stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct OleHeader {
    /// The header's version field (`0x0002_0000` for every known writer).
    pub version: u32,
    /// The registered clipboard format of the equation data.
    pub clipboard_format: u16,
    /// The declared length of the MTEF data that follows the header.
    pub object_size: u32,
}

/// What MTEF data declared about itself. A caller that rewrites an
/// embedded equation uses it to keep the object's kind (Equation Editor 3
/// versus MathType).
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct MtefInfo {
    /// The MTEF format version: 3 (Equation Editor 3) or 5 (MathType 4 and
    /// later).
    pub version: u8,
    /// The product that wrote the equation.
    pub product: MtefProduct,
    /// The product's major version.
    pub product_version: u8,
    /// The product's minor version.
    pub product_subversion: u8,
    /// The application key of an MTEF 5 equation (for example `DSMT6`).
    pub application_key: Option<String>,
    /// Whether an MTEF 5 equation is inline (`Some(true)`) or display
    /// (`Some(false)`); MTEF 3 does not record it.
    pub inline: Option<bool>,
    /// The OLE header, when the input was an `Equation Native` stream.
    pub ole_header: Option<OleHeader>,
}

impl MtefInfo {
    fn from_mathema(info: &mathema::MtefInfo) -> Self {
        Self {
            version: info.version.number(),
            product: match info.product {
                mathema::MtefProduct::MathType => MtefProduct::MathType,
                mathema::MtefProduct::EquationEditor => MtefProduct::EquationEditor,
                _ => MtefProduct::Other,
            },
            product_version: info.product_version,
            product_subversion: info.product_subversion,
            application_key: info.application_key.clone(),
            inline: info.inline,
            ole_header: info.ole_header.map(|header| OleHeader {
                version: header.version,
                clipboard_format: header.clipboard_format,
                object_size: header.object_size,
            }),
        }
    }
}

/// An MTEF equation read from an `Equation Native` stream or bare MTEF data.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MtefEquation {
    latex: String,
    info: Option<MtefInfo>,
}

impl MtefEquation {
    /// The equation as LaTeX; empty for an equation without content.
    #[must_use]
    pub fn latex(&self) -> &str {
        &self.latex
    }

    /// The equation as LaTeX, by value.
    #[must_use]
    pub fn into_latex(self) -> String {
        self.latex
    }

    /// What the data declared about itself: MTEF version, generating
    /// product and platform, and the OLE header of an `Equation Native`
    /// stream. `None` for empty input.
    #[must_use]
    pub fn info(&self) -> Option<&MtefInfo> {
        self.info.as_ref()
    }
}

/// Convert an OMML equation (`m:oMath`, `m:oMathPara` or a fragment of
/// them) to LaTeX.
///
/// The OMML may omit its namespace declaration when it uses the
/// conventional `m` prefix, as a fragment cut out of a document does.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for malformed XML or OMML, and the limit
/// errors when the input exceeds `limits`.
pub fn omml_to_latex(omml: &str, limits: &Limits) -> Result<String> {
    let parsed = mathema::from_omml_with(omml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    parsed
        .to_latex()
        .map_err(|error| Error::from_mathema(error, Format::Latex))
}

/// Convert an OMML equation to UnicodeMath, the linear form Office types
/// and shows for an equation in plain text (Unicode Technical Note 28).
///
/// # Errors
///
/// As [`omml_to_latex`].
pub fn omml_to_unicode_math(omml: &str, limits: &Limits) -> Result<String> {
    let parsed = mathema::from_omml_with(omml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    parsed
        .to_unicode_math()
        .map_err(|error| Error::Internal(error.to_string()))
}

/// Convert a MathML equation (`math`) to LaTeX.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for malformed XML or MathML, and the limit
/// errors when the input exceeds `limits`.
pub fn mathml_to_latex(mathml: &str, limits: &Limits) -> Result<String> {
    let parsed = mathema::from_mathml_with(mathml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::MathMl))?;
    parsed
        .to_latex()
        .map_err(|error| Error::from_mathema(error, Format::Latex))
}

/// Read binary MTEF: the contents of an OLE `Equation Native` stream (its
/// 28-byte header is detected and validated) or bare MTEF 3 or MTEF 5 data.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for data that is not valid MTEF, and the
/// limit errors when the input exceeds `limits`.
pub fn read_mtef(bytes: &[u8], limits: &Limits) -> Result<MtefEquation> {
    let parsed = mathema::from_mtef_bytes_with(bytes, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Mtef))?;
    let latex = parsed
        .to_latex()
        .map_err(|error| Error::from_mathema(error, Format::Latex))?;
    Ok(MtefEquation {
        latex,
        info: parsed.mtef_info().map(MtefInfo::from_mathema),
    })
}

/// Convert binary MTEF to LaTeX; see [`read_mtef`].
///
/// # Errors
///
/// As [`read_mtef`].
pub fn mtef_to_latex(bytes: &[u8], limits: &Limits) -> Result<String> {
    read_mtef(bytes, limits).map(MtefEquation::into_latex)
}

/// An equation written as OMML, with what the conversion could not carry
/// over.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Omml {
    xml: String,
    losses: Vec<String>,
}

impl Omml {
    /// The `m:oMath` element.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// The `m:oMath` element, by value.
    #[must_use]
    pub fn into_xml(self) -> String {
        self.xml
    }

    /// One message per piece of the input the OMML does not represent (for
    /// example run colour with [`OmmlHost::Bare`]); empty when nothing was
    /// lost.
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// Convert LaTeX math to an `m:oMath` element for `host`, with the math
/// namespace (and the host's run-property namespace) declared on it.
///
/// Formatting the host cannot express is dropped and reported in
/// [`Omml::losses`]; the caller decides whether a lossy result is
/// acceptable.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for LaTeX that does not parse,
/// [`Error::Unsupported`] for a construct the writer refuses, and the limit
/// errors when the input exceeds `limits`.
pub fn latex_to_omml(latex: &str, host: OmmlHost, limits: &Limits) -> Result<Omml> {
    latex_to_omml_with(
        latex,
        mathema::OmmlOptions::default().with_run_properties(host.run_properties()),
        limits,
    )
}

/// How the equations of a display-math paragraph line up.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Justification {
    /// Left-aligned.
    Left,
    /// Right-aligned.
    Right,
    /// Each equation centred on its own.
    Center,
    /// Left-aligned to each other and centred as a group; what Office
    /// assumes when the paragraph names none.
    #[default]
    CenterGroup,
}

impl Justification {
    const fn to_mathema(self) -> mathema::OmmlJustification {
        match self {
            Self::Left => mathema::OmmlJustification::Left,
            Self::Right => mathema::OmmlJustification::Right,
            Self::Center => mathema::OmmlJustification::Center,
            Self::CenterGroup => mathema::OmmlJustification::CenterGroup,
        }
    }
}

/// Convert LaTeX math to a display-math paragraph (`m:oMathPara` around one
/// `m:oMath`) for `host`, with `justification` written as its `m:jc` when
/// given; otherwise as [`latex_to_omml`].
///
/// # Errors
///
/// As [`latex_to_omml`].
pub fn latex_to_omml_paragraph(
    latex: &str,
    host: OmmlHost,
    justification: Option<Justification>,
    limits: &Limits,
) -> Result<Omml> {
    latex_to_omml_with(
        latex,
        mathema::OmmlOptions::default()
            .with_run_properties(host.run_properties())
            .with_display_block(justification.map(Justification::to_mathema)),
        limits,
    )
}

fn latex_to_omml_with(latex: &str, options: mathema::OmmlOptions, limits: &Limits) -> Result<Omml> {
    let parsed = mathema::from_latex_with(latex, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Latex))?;
    let (xml, diagnostics) = parsed
        .to_omml_with_options(&options)
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    let losses = diagnostics
        .into_iter()
        .filter(|diagnostic| diagnostic.kind == mathema::DiagnosticKind::Warning)
        .map(|diagnostic| diagnostic.message)
        .collect();
    Ok(Omml { xml, losses })
}
