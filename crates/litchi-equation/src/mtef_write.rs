//! Writing equations as the `Equation Native` stream of an OLE equation
//! object: MathType (`Equation.DSMT4`, MTEF 5) or Equation Editor 3.0
//! (`Equation.3`, MTEF 3).

use crate::{Error, Format, Limits, MtefInfo, MtefProduct, Result};

/// Which kind of OLE equation object to write.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MtefTarget {
    /// MathType: ProgID `Equation.DSMT4`, CLSID
    /// `{0002CE03-0000-0000-C000-000000000046}`, MTEF 5.
    MathType,
    /// Equation Editor 3.0: ProgID `Equation.3`, CLSID
    /// `{0002CE02-0000-0000-C000-000000000046}`, MTEF 3.
    EquationEditor3,
    /// The MTEF version and product of an existing equation, so that an
    /// updated object keeps its kind.
    SameAs(MtefInfo),
}

/// `{0002CE02-0000-0000-C000-000000000046}` in on-disk byte order.
const EQUATION_3_CLSID: [u8; 16] = [
    0x02, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
/// `{0002CE03-0000-0000-C000-000000000046}` in on-disk byte order.
const MATHTYPE_CLSID: [u8; 16] = [
    0x03, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

impl MtefTarget {
    /// Whether the target writes MTEF 3 (Equation Editor 3.0).
    fn is_equation_editor(&self) -> bool {
        match self {
            Self::MathType => false,
            Self::EquationEditor3 => true,
            Self::SameAs(info) => info.version < 5,
        }
    }

    /// The OLE ProgID of the object: `Equation.DSMT4` or `Equation.3`.
    #[must_use]
    pub fn prog_id(&self) -> &'static str {
        if self.is_equation_editor() {
            "Equation.3"
        } else {
            "Equation.DSMT4"
        }
    }

    /// The object's CLSID in on-disk (little-endian GUID) byte order.
    #[must_use]
    pub fn clsid(&self) -> [u8; 16] {
        if self.is_equation_editor() {
            EQUATION_3_CLSID
        } else {
            MATHTYPE_CLSID
        }
    }

    /// The user type an application records in the object's `CompObj`
    /// stream.
    #[must_use]
    pub fn user_type(&self) -> &'static str {
        if self.is_equation_editor() {
            "Microsoft Equation 3.0"
        } else {
            "MathType 6.0 Equation"
        }
    }

    fn write_options(&self) -> mathema::MtefWriteOptions {
        let version = if self.is_equation_editor() {
            mathema::MtefVersion::V3
        } else {
            mathema::MtefVersion::V5
        };
        let mut options = mathema::MtefWriteOptions::new(version).with_ole_header(true);
        if let Self::SameAs(info) = self {
            match info.product {
                MtefProduct::MathType => options.product = mathema::MtefProduct::MathType,
                MtefProduct::EquationEditor => {
                    options.product = mathema::MtefProduct::EquationEditor;
                },
                MtefProduct::Other => {},
            }
            options.product_version = info.product_version;
            options.product_subversion = info.product_subversion;
            if let Some(key) = info
                .application_key
                .as_ref()
                .filter(|_| version == mathema::MtefVersion::V5)
            {
                options.application_key.clone_from(key);
            }
        }
        options
    }
}

/// A complete `Equation Native` stream: the 28-byte `EQNOLEFILEHDR`
/// followed by MTEF data, with what the conversion could not carry over.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EquationNative {
    bytes: Vec<u8>,
    info: MtefInfo,
    losses: Vec<String>,
}

impl EquationNative {
    /// The stream bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The stream bytes, by value.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// What the written data declares about itself.
    #[must_use]
    pub fn info(&self) -> &MtefInfo {
        &self.info
    }

    /// One message per piece of the input MTEF does not represent (for
    /// example a colour); empty when nothing was lost.
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// Convert LaTeX math to the `Equation Native` stream of a `target`
/// object.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for LaTeX that does not parse,
/// [`Error::Unsupported`] for a construct the writer refuses, and the limit
/// errors when the input exceeds `limits`.
pub fn latex_to_equation_native(
    latex: &str,
    target: &MtefTarget,
    limits: &Limits,
) -> Result<EquationNative> {
    let parsed = mathema::from_latex_with(latex, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Latex))?;
    write(&parsed, target, limits)
}

/// Convert an OMML equation (`m:oMath`, `m:oMathPara` or a fragment of
/// them) to the `Equation Native` stream of a `target` object.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for malformed XML or OMML,
/// [`Error::Unsupported`] for a construct the writer refuses, and the limit
/// errors when the input exceeds `limits`.
pub fn omml_to_equation_native(
    omml: &str,
    target: &MtefTarget,
    limits: &Limits,
) -> Result<EquationNative> {
    let parsed = mathema::from_omml_with(omml, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Omml))?;
    write(&parsed, target, limits)
}

fn write(
    parsed: &mathema::ConvertResult,
    target: &MtefTarget,
    limits: &Limits,
) -> Result<EquationNative> {
    let (bytes, diagnostics) = parsed
        .to_mtef_bytes_with_diagnostics(&target.write_options())
        .map_err(|error| Error::from_mathema(error, Format::Mtef))?;
    let losses = diagnostics
        .into_iter()
        .filter(|diagnostic| diagnostic.kind == mathema::DiagnosticKind::Warning)
        .map(|diagnostic| diagnostic.message)
        .collect();
    // Read the written stream back: it proves the output is MTEF the reader
    // accepts and yields the header facts the caller keeps.
    let written = mathema::from_mtef_bytes_with(&bytes, &limits.config()).map_err(|error| {
        Error::Internal(format!("the written MTEF does not read back: {error}"))
    })?;
    let info = written
        .mtef_info()
        .map(MtefInfo::from_mathema)
        .ok_or_else(|| Error::Internal("the written MTEF has no header".to_owned()))?;
    Ok(EquationNative {
        bytes,
        info,
        losses,
    })
}

/// Render an equation stored as MTEF (an `Equation Native` stream or bare
/// MTEF data) as a PNG picture with a transparent background: the preview
/// an office file keeps beside an OLE equation.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for data that is not valid MTEF, the limit
/// errors when the input exceeds `limits`, and [`Error::Render`] when the
/// picture would exceed [`crate::PngOptions::max_pixels`] or cannot be
/// drawn.
#[cfg(feature = "render")]
pub fn mtef_to_png(
    equation_native_or_mtef: &[u8],
    options: &crate::PngOptions,
    limits: &Limits,
) -> Result<crate::Png> {
    let parsed = mathema::from_mtef_bytes_with(equation_native_or_mtef, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Mtef))?;
    crate::render::render(&parsed, options)
}
