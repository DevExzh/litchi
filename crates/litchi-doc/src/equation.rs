//! MathType and Equation Editor objects of legacy Word documents.
//!
//! An equation is an OLE object: a storage `ObjectPool/_<id>` holding its
//! MTEF in an `Equation Native` stream, shown by an `EMBED` field whose
//! picture character carries a preview picture. [`EquationContent`] is the
//! content one object holds (its `Equation Native` stream and preview), and
//! [`EquationKind`] the class of a new object. With the `equation` feature,
//! [`EquationContent`] is made from LaTeX, OMML or MTEF through
//! `litchi-equation`, which also renders the preview.

use super::package::{Error as PackageError, Result};
use super::writer::images::Picture;
use litchi_cfb::OleWriter;
use std::io::Cursor;

const EQNOLE_HEADER_SIZE: usize = 28;
const MAX_MTEF_PAYLOAD_SIZE: usize = 16 * 1024 * 1024;
const EQNOLE_VERSION: u32 = 0x0002_0000;
const MTEF_CLIPBOARD_FORMAT: u16 = 0xC1C6;
const MIN_MTEF_CLIPBOARD_FORMAT: u16 = 0xC100;
const MAX_MTEF_CLIPBOARD_FORMAT: u16 = 0xC2FF;
const MTEF_END: u8 = 0;
const EQUATION_NATIVE_STREAM: &str = "Equation Native";
const COMP_OBJ_STREAM: &str = "\u{1}CompObj";
const OLE_STREAM: &str = "\u{1}Ole";
const OBJ_INFO_STREAM: &str = "\u{3}ObjInfo";

/// Equation Editor 3.0 CLSID `{0002CE02-0000-0000-C000-000000000046}` in
/// CFB byte order.
pub const EQUATION_3_CLSID: [u8; 16] = [
    0x02, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// A bounded, envelope-validated Equation Native stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MtefEquationObject {
    equation_native: Vec<u8>,
}

impl MtefEquationObject {
    /// Construct an Equation Native stream from a headerless MTEF payload.
    pub fn from_mtef_payload(payload: Vec<u8>) -> Result<Self> {
        validate_mtef_payload(&payload)?;
        let payload_len = u32::try_from(payload.len())
            .map_err(|_| invalid("MTEF payload exceeds the Equation Native size field"))?;
        let mut bytes = Vec::with_capacity(EQNOLE_HEADER_SIZE + payload.len());
        bytes.extend_from_slice(&(EQNOLE_HEADER_SIZE as u16).to_le_bytes());
        bytes.extend_from_slice(&EQNOLE_VERSION.to_le_bytes());
        bytes.extend_from_slice(&MTEF_CLIPBOARD_FORMAT.to_le_bytes());
        bytes.extend_from_slice(&payload_len.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0x0014_F690u32.to_le_bytes());
        bytes.extend_from_slice(&0x0014_EBB4u32.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&payload);
        Ok(Self {
            equation_native: bytes,
        })
    }

    /// Validate and retain a complete 28-byte-header Equation Native stream.
    pub fn from_equation_native(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() < EQNOLE_HEADER_SIZE {
            return Err(invalid("Equation Native header is truncated"));
        }
        if u16::from_le_bytes(bytes[0..2].try_into().expect("fixed slice"))
            != EQNOLE_HEADER_SIZE as u16
        {
            return Err(invalid("Equation Native header length is not 28"));
        }
        let version = u32::from_le_bytes(bytes[2..6].try_into().expect("fixed slice"));
        if version != EQNOLE_VERSION && version != 0x0000_0200 {
            return Err(invalid("Equation Native version is unsupported"));
        }
        let format = u16::from_le_bytes(bytes[6..8].try_into().expect("fixed slice"));
        if !(MIN_MTEF_CLIPBOARD_FORMAT..=MAX_MTEF_CLIPBOARD_FORMAT).contains(&format) {
            return Err(invalid("Equation Native clipboard format is invalid"));
        }
        let declared = u32::from_le_bytes(bytes[8..12].try_into().expect("fixed slice")) as usize;
        if declared > MAX_MTEF_PAYLOAD_SIZE
            || declared
                .checked_add(EQNOLE_HEADER_SIZE)
                .is_none_or(|length| length != bytes.len())
        {
            return Err(invalid("Equation Native object size is invalid"));
        }
        validate_mtef_payload(&bytes[EQNOLE_HEADER_SIZE..])?;
        Ok(Self {
            equation_native: bytes,
        })
    }

    /// Complete bytes stored in the `Equation Native` stream.
    #[must_use]
    pub fn equation_native(&self) -> &[u8] {
        &self.equation_native
    }

    /// Headerless MTEF record stream.
    #[must_use]
    pub fn mtef_payload(&self) -> &[u8] {
        &self.equation_native[EQNOLE_HEADER_SIZE..]
    }

    /// Consume this value and return the complete Equation Native stream.
    #[must_use]
    pub fn into_equation_native(self) -> Vec<u8> {
        self.equation_native
    }
}

/// The class of an OLE equation object.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum EquationKind {
    /// MathType (`Equation.DSMT4`), holding MTEF 5.
    MathType,
    /// Microsoft Equation Editor 3.0 (`Equation.3`), holding MTEF 3.
    EquationEditor3,
}

impl EquationKind {
    /// The kind whose objects hold MTEF of `version`: 3 for Equation Editor
    /// 3.0, anything else for MathType.
    #[must_use]
    pub const fn for_mtef_version(version: u8) -> Self {
        if version == 3 {
            Self::EquationEditor3
        } else {
            Self::MathType
        }
    }

    /// The OLE ProgID, which the `EMBED` field instruction names.
    #[must_use]
    pub const fn prog_id(self) -> &'static str {
        match self {
            Self::MathType => "Equation.DSMT4",
            Self::EquationEditor3 => "Equation.3",
        }
    }

    /// The CLSID of the object storage, in CFB byte order.
    #[must_use]
    pub const fn clsid(self) -> [u8; 16] {
        match self {
            Self::MathType => MATH_TYPE_CLSID,
            Self::EquationEditor3 => EQUATION_3_CLSID,
        }
    }

    /// The user type the application records in the `CompObj` stream.
    const fn user_type(self) -> &'static str {
        match self {
            Self::MathType => "MathType 6.0 Equation",
            Self::EquationEditor3 => "Microsoft Equation 3.0",
        }
    }

    /// The clipboard format name recorded in the `CompObj` stream.
    const fn clipboard_format(self) -> &'static str {
        match self {
            Self::MathType => "MathType EF",
            Self::EquationEditor3 => "DS Equation",
        }
    }
}

/// MathType `Equation.DSMT4` CLSID `{0002CE03-0000-0000-C000-000000000046}`
/// in CFB byte order.
pub const MATH_TYPE_CLSID: [u8; 16] = [
    0x03, 0xCE, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// The content of one equation object: its `Equation Native` stream and the
/// picture Word shows for it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EquationContent {
    equation: MtefEquationObject,
    preview: Picture,
    lowering: u16,
}

impl EquationContent {
    /// Content from a validated `Equation Native` stream, its preview, and
    /// how far the preview reaches below the equation's baseline in
    /// half-points. The preview is lowered by that much
    /// (`sprmCPicLocation`'s character gets `sprmCHpsPos`), as Word lowers
    /// each equation so that its baseline sits on the text line.
    #[must_use]
    pub fn new(equation: MtefEquationObject, preview: Picture, lowering: u16) -> Self {
        Self {
            equation,
            preview,
            lowering,
        }
    }

    /// The `Equation Native` stream.
    #[must_use]
    pub fn equation(&self) -> &MtefEquationObject {
        &self.equation
    }

    /// The preview picture.
    #[must_use]
    pub fn preview(&self) -> &Picture {
        &self.preview
    }

    /// How far the preview is lowered below the baseline, in half-points.
    #[must_use]
    pub fn lowering(&self) -> u16 {
        self.lowering
    }

    /// The MTEF version of the stream: 3 or 5.
    #[must_use]
    pub fn mtef_version(&self) -> u8 {
        self.equation
            .mtef_payload()
            .first()
            .copied()
            .unwrap_or_default()
    }
}

#[cfg(feature = "equation")]
impl EquationContent {
    /// Content from an `Equation Native` stream (or bare MTEF, which gets
    /// the stream header), with its preview rendered from the MTEF at
    /// `font_size_pt`, the size of the text the equation stands in (Word
    /// draws an equation at its text's size).
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Equation`] when the MTEF does not read or
    /// render within `limits` or the font size is out of range, and
    /// [`PackageError::InvalidFormat`] for a stream whose envelope is
    /// invalid.
    pub fn from_mtef(
        bytes: Vec<u8>,
        font_size_pt: f64,
        limits: &litchi_equation::Limits,
    ) -> Result<Self> {
        litchi_equation::read_mtef(&bytes, limits).map_err(PackageError::Equation)?;
        let header = (EQNOLE_HEADER_SIZE as u16).to_le_bytes();
        let equation = if bytes.len() >= EQNOLE_HEADER_SIZE && bytes[..2] == header {
            MtefEquationObject::from_equation_native(bytes)?
        } else {
            MtefEquationObject::from_mtef_payload(bytes)?
        };
        let (preview, lowering) = render_preview(equation.equation_native(), font_size_pt, limits)?;
        Ok(Self {
            equation,
            preview,
            lowering,
        })
    }

    /// Content converted from LaTeX math into an object of `target`'s
    /// kind, its preview drawn at `font_size_pt`, and the messages of what
    /// the conversion lost.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Equation`] when the LaTeX does not convert
    /// or render within `limits` or the font size is out of range.
    pub fn from_latex(
        latex: &str,
        target: &litchi_equation::MtefTarget,
        font_size_pt: f64,
        limits: &litchi_equation::Limits,
    ) -> Result<(Self, Vec<String>)> {
        let native = litchi_equation::latex_to_equation_native(latex, target, limits)
            .map_err(PackageError::Equation)?;
        Self::from_written(native, font_size_pt, limits)
    }

    /// Content converted from an OMML equation into an object of
    /// `target`'s kind, and the messages of what the conversion lost.
    ///
    /// # Errors
    ///
    /// As [`Self::from_latex`].
    pub fn from_omml(
        omml: &str,
        target: &litchi_equation::MtefTarget,
        font_size_pt: f64,
        limits: &litchi_equation::Limits,
    ) -> Result<(Self, Vec<String>)> {
        let native = litchi_equation::omml_to_equation_native(omml, target, limits)
            .map_err(PackageError::Equation)?;
        Self::from_written(native, font_size_pt, limits)
    }

    fn from_written(
        native: litchi_equation::EquationNative,
        font_size_pt: f64,
        limits: &litchi_equation::Limits,
    ) -> Result<(Self, Vec<String>)> {
        let losses = native.losses().to_vec();
        let equation = MtefEquationObject::from_equation_native(native.into_bytes())?;
        let (preview, lowering) = render_preview(equation.equation_native(), font_size_pt, limits)?;
        Ok((
            Self {
                equation,
                preview,
                lowering,
            },
            losses,
        ))
    }
}

/// Render the preview of an equation at `font_size_pt`, drawn at twice the
/// screen resolution and shown at its size in points, and how far the same
/// picture reaches below the baseline in half-points.
#[cfg(feature = "equation")]
fn render_preview(
    equation_native: &[u8],
    font_size_pt: f64,
    limits: &litchi_equation::Limits,
) -> Result<(Picture, u16)> {
    let options = litchi_equation::PngOptions::new(font_size_pt)
        .and_then(|options| options.with_scale(2.0))
        .map_err(PackageError::Equation)?;
    let png = litchi_equation::mtef_to_png(equation_native, &options, limits)
        .map_err(PackageError::Equation)?;
    let (width, height) = png.size_pt();
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "the value is clamped to the PICF goal range 1..=32767"
    )]
    let twips = |points: f64| (points * 20.0).round().clamp(1.0, f64::from(i16::MAX)) as u32;
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "the value is clamped to the sprmCHpsPos range 0..=3168"
    )]
    let lowering = (png.depth_pt() * 2.0).round().clamp(0.0, 3168.0) as u16;
    let picture = Picture::from_parts(png.into_bytes(), twips(width), twips(height))
        .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
    Ok((picture, lowering))
}

/// The compound file of a new equation object of `kind` holding
/// `equation_native`, with the streams Word writes for one: `CompObj`,
/// `Ole`, `ObjInfo` and `Equation Native`.
pub(crate) fn equation_object_compound_file(
    kind: EquationKind,
    equation_native: &[u8],
) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.set_root_clsid(kind.clsid());
    writer.create_stream(&[COMP_OBJ_STREAM], &comp_obj(kind))?;
    writer.create_stream(&[OLE_STREAM], WORD_OLE)?;
    writer.create_stream(&[OBJ_INFO_STREAM], WORD_OBJ_INFO)?;
    writer.create_stream(&[EQUATION_NATIVE_STREAM], equation_native)?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

/// `CompObjStream` ([MS-OLEDS] 2.3.8) for `kind`, as Word writes it: the
/// header, the class, and the user type, clipboard format and ProgID as
/// ANSI strings, then the Unicode marker without Unicode strings.
fn comp_obj(kind: EquationKind) -> Vec<u8> {
    let mut bytes = vec![0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00];
    bytes.extend_from_slice(&[0xFF; 4]);
    bytes.extend_from_slice(&kind.clsid());
    for text in [kind.user_type(), kind.clipboard_format(), kind.prog_id()] {
        let length = u32::try_from(text.len() + 1).unwrap_or(u32::MAX);
        bytes.extend_from_slice(&length.to_le_bytes());
        bytes.extend_from_slice(text.as_bytes());
        bytes.push(0);
    }
    bytes.extend_from_slice(&0x71B2_39F4u32.to_le_bytes());
    bytes.extend_from_slice(&[0; 12]);
    bytes
}

/// The `Ole` stream Word writes for an embedded equation ([MS-OLEDS]
/// 2.3.3).
const WORD_OLE: &[u8] = &[
    0x01, 0x00, 0x00, 0x02, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
];
/// The `ObjInfo` stream Word writes for an embedded equation (an `ODT`,
/// [MS-DOC] 2.9): clipboard format 3 (metafile picture).
const WORD_OBJ_INFO: &[u8] = &[0x00, 0x00, 0x03, 0x00, 0x04, 0x00];

fn validate_mtef_payload(payload: &[u8]) -> Result<()> {
    if payload.len() > MAX_MTEF_PAYLOAD_SIZE {
        return Err(invalid("MTEF payload exceeds 16 MiB"));
    }
    if payload.last() != Some(&MTEF_END) {
        return Err(invalid("MTEF payload must end with an END record"));
    }
    let version = *payload
        .first()
        .ok_or_else(|| invalid("MTEF payload is empty"))?;
    let header_len = match version {
        1 | 101 => 1,
        2..=4 => 5,
        5 => {
            if payload.len() < 7 {
                return Err(invalid("MTEF 5 header is truncated"));
            }
            let application_key_end = payload[5..]
                .iter()
                .position(|byte| *byte == 0)
                .map(|relative| relative + 5)
                .ok_or_else(|| invalid("MTEF 5 application key is unterminated"))?;
            application_key_end
                .checked_add(2)
                .ok_or_else(|| invalid("MTEF header length overflow"))?
        },
        _ => return Err(invalid("MTEF version is unsupported")),
    };
    if payload.len() <= header_len {
        return Err(invalid("MTEF payload contains no record stream"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> PackageError {
    PackageError::InvalidFormat(message.into())
}
