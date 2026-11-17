//! Writing equations as the `Equation Native` stream of an OLE equation
//! object: MathType (`Equation.DSMT4`, MTEF 5) or Equation Editor 3.0
//! (`Equation.3`, MTEF 3).

use crate::{Error, Format, Limits, MtefInfo, MtefProduct, Result};

/// `{0002CE03-0000-0000-C000-000000000046}`, MathType's `Equation.DSMT4`, in
/// on-disk (little-endian GUID) byte order.
const MATH_TYPE_CLSID: [u8; 16] = [
    0x03, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];
/// `{0002CE02-0000-0000-C000-000000000046}`, Equation Editor 3.0's
/// `Equation.3`, in on-disk byte order.
const EQUATION_EDITOR_3_CLSID: [u8; 16] = [
    0x02, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// Which kind of OLE equation object to write.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum MtefTarget {
    /// MathType: ProgID `Equation.DSMT4`, CLSID
    /// `{0002CE03-0000-0000-C000-000000000046}`, MTEF 5 as MathType 6
    /// (`DSMT6`) writes it.
    MathType,
    /// Equation Editor 3.0: ProgID `Equation.3`, CLSID
    /// `{0002CE02-0000-0000-C000-000000000046}`, MTEF 3.
    EquationEditor3,
    /// The MTEF version, product and application key of an existing
    /// equation, so that an updated object keeps its kind. An MTEF 3
    /// equation is an `Equation.3` object, any other an `Equation.DSMT4`
    /// one. A product code other than MathType and Equation Editor is
    /// written as the version's usual product.
    SameAs(MtefInfo),
}

impl MtefTarget {
    fn is_mtef_3(&self) -> bool {
        match self {
            Self::MathType => false,
            Self::EquationEditor3 => true,
            Self::SameAs(info) => info.version == 3,
        }
    }

    /// The kind of equation object a ProgID names: [`Self::MathType`] for
    /// `Equation.DSMT4` and the other `Equation.DSMT*` ProgIDs,
    /// [`Self::EquationEditor3`] for `Equation.3`, `None` for any other
    /// class. Case is ignored, as COM ignores it.
    #[must_use]
    pub fn from_prog_id(prog_id: &str) -> Option<Self> {
        if prog_id.eq_ignore_ascii_case("Equation.3") {
            Some(Self::EquationEditor3)
        } else if prog_id
            .get(..13)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Equation.DSMT"))
        {
            Some(Self::MathType)
        } else {
            None
        }
    }

    /// The OLE ProgID of the object: `Equation.DSMT4` or `Equation.3`.
    #[must_use]
    pub fn prog_id(&self) -> &'static str {
        if self.is_mtef_3() {
            "Equation.3"
        } else {
            "Equation.DSMT4"
        }
    }

    /// The object's CLSID in on-disk (little-endian GUID) byte order.
    #[must_use]
    pub fn clsid(&self) -> [u8; 16] {
        if self.is_mtef_3() {
            EQUATION_EDITOR_3_CLSID
        } else {
            MATH_TYPE_CLSID
        }
    }

    /// The user type the applications record in the object's `CompObj`
    /// stream: `Microsoft Equation 3.0` or `MathType 6.0 Equation`.
    #[must_use]
    pub fn user_type(&self) -> &'static str {
        if self.is_mtef_3() {
            "Microsoft Equation 3.0"
        } else {
            "MathType 6.0 Equation"
        }
    }

    fn options(&self) -> mathema::MtefWriteOptions {
        let options = match self {
            Self::MathType => mathema::MtefWriteOptions::new(mathema::MtefVersion::V5),
            Self::EquationEditor3 => mathema::MtefWriteOptions::new(mathema::MtefVersion::V3),
            Self::SameAs(info) => {
                let version = if info.version == 3 {
                    mathema::MtefVersion::V3
                } else {
                    mathema::MtefVersion::V5
                };
                let mut options = mathema::MtefWriteOptions::new(version);
                match info.product {
                    MtefProduct::MathType => {
                        options.product = mathema::MtefProduct::MathType;
                    },
                    MtefProduct::EquationEditor => {
                        options.product = mathema::MtefProduct::EquationEditor;
                    },
                    MtefProduct::Other => {},
                }
                if info.product != MtefProduct::Other {
                    options.product_version = info.product_version;
                    options.product_subversion = info.product_subversion;
                }
                if version == mathema::MtefVersion::V5
                    && let Some(key) = &info.application_key
                {
                    options.application_key.clone_from(key);
                }
                options
            },
        };
        options.with_ole_header(true)
    }
}

/// A complete `Equation Native` stream: the 28-byte `EQNOLEFILEHDR` and
/// the MTEF data, ready to store in an OLE equation object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EquationNative {
    bytes: Vec<u8>,
    info: MtefInfo,
    losses: Vec<String>,
}

impl EquationNative {
    /// The stream's bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The stream's bytes, by value.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }

    /// What the written stream declares about itself: its MTEF version,
    /// product and OLE header.
    #[must_use]
    pub fn info(&self) -> &MtefInfo {
        &self.info
    }

    /// One message per piece of the input MTEF does not represent (for
    /// example colour); empty when nothing was lost.
    #[must_use]
    pub fn losses(&self) -> &[String] {
        &self.losses
    }
}

/// Convert LaTeX math to the `Equation Native` stream of an equation object
/// of `target`'s kind.
///
/// # Errors
///
/// Returns [`Error::Malformed`] for LaTeX that does not parse,
/// [`Error::Unsupported`] for a construct MTEF cannot hold (an array of more
/// than 255 rows or columns), the limit errors when the input or the written
/// equation exceeds `limits`, and [`Error::Internal`] when the writer
/// produces data its reader rejects.
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
/// them) to the `Equation Native` stream of an equation object of
/// `target`'s kind.
///
/// # Errors
///
/// As [`latex_to_equation_native`], with [`Error::Malformed`] for malformed
/// XML or OMML.
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
        .to_mtef_bytes_with_diagnostics(&target.options())
        .map_err(|error| Error::from_mathema(error, Format::Mtef))?;
    let info = mathema::from_mtef_bytes_with(&bytes, &limits.config())
        .map_err(|error| Error::from_mathema(error, Format::Mtef))?
        .mtef_info()
        .map(MtefInfo::from_mathema)
        .ok_or_else(|| Error::Internal("the written MTEF declares no header".to_owned()))?;
    let losses = diagnostics
        .into_iter()
        .filter(|diagnostic| diagnostic.kind == mathema::DiagnosticKind::Warning)
        .map(|diagnostic| diagnostic.message)
        .collect();
    Ok(EquationNative {
        bytes,
        info,
        losses,
    })
}

/// Render binary MTEF (an `Equation Native` stream or bare MTEF data) as a
/// PNG picture with a transparent background, for the preview an OLE
/// equation object shows.
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_name_their_object_kind() {
        assert_eq!(MtefTarget::MathType.prog_id(), "Equation.DSMT4");
        assert_eq!(MtefTarget::EquationEditor3.prog_id(), "Equation.3");
        assert_eq!(MtefTarget::MathType.clsid()[0], 0x03);
        assert_eq!(MtefTarget::EquationEditor3.clsid()[0], 0x02);
    }

    #[test]
    fn writes_equation_native_streams_of_both_kinds() {
        let limits = Limits::DEFAULT;
        let v5 = latex_to_equation_native(r"\frac{1}{2}", &MtefTarget::MathType, &limits).unwrap();
        assert_eq!(&v5.bytes()[..2], &[28, 0]);
        assert_eq!(v5.bytes()[28], 5);
        assert_eq!(v5.info().version, 5);
        assert!(v5.info().ole_header.is_some());
        assert_eq!(
            crate::mtef_to_latex(v5.bytes(), &limits).unwrap(),
            r"\frac{1}{2}"
        );

        let v3 = latex_to_equation_native("x^2", &MtefTarget::EquationEditor3, &limits).unwrap();
        assert_eq!(v3.bytes()[28], 3);
        let same = MtefTarget::SameAs(v3.info().clone());
        assert_eq!(same.prog_id(), "Equation.3");
        let again = latex_to_equation_native("y", &same, &limits).unwrap();
        assert_eq!(&again.bytes()[28..33], &v3.bytes()[28..33]);

        let omml = omml_to_equation_native(
            "<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>",
            &MtefTarget::MathType,
            &limits,
        )
        .unwrap();
        assert_eq!(crate::mtef_to_latex(omml.bytes(), &limits).unwrap(), "a");
        assert!(latex_to_equation_native(r"\frac{", &MtefTarget::MathType, &limits).is_err());
    }
}
