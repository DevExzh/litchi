//! MathType/Equation Editor object authoring for legacy Word documents.
//!
//! The writer accepts an already encoded MTEF payload, validates its bounded
//! envelope, and creates the complete Equation.3 compound object plus the DOC
//! EMBED field and bitmap presentation. Formula-AST-to-MTEF conversion is a
//! separate concern and is intentionally not approximated here.

use super::embedded_object::{Editor, Reference, WriteOptions};
use super::package::{Error as PackageError, Result};
use super::writer::images::{DocPicture, write_picture_block};
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

// Canonical Equation.3 registration streams emitted by LibreOffice.
const EQUATION_COMP_OBJ: &[u8] = &[
    0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0x02, 0xCE, 0x02, 0x00,
    0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46, 0x17, 0x00, 0x00, 0x00,
    b'M', b'i', b'c', b'r', b'o', b's', b'o', b'f', b't', b' ', b'E', b'q', b'u', b'a', b't', b'i',
    b'o', b'n', b' ', b'3', b'.', b'0', 0x00, 0x0C, 0x00, 0x00, 0x00, b'D', b'S', b' ', b'E', b'q',
    b'u', b'a', b't', b'i', b'o', b'n', 0x00, 0x0B, 0x00, 0x00, 0x00, b'E', b'q', b'u', b'a', b't',
    b'i', b'o', b'n', b'.', b'3', 0x00, 0xF4, 0x39, 0xB2, 0x71, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
];
const EQUATION_OLE: &[u8] = &[
    0x01, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
];
// ODTPersist1: fRecomposeOnResize | fViewObject; clipboard format 3
// (metafile); ODTPersist2: no EMF flags.
const EQUATION_OBJ_INFO: &[u8] = &[0x00, 0x82, 0x03, 0x00, 0x00, 0x00];

/// A bounded, envelope-validated Equation Native stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MtefEquation {
    equation_native: Vec<u8>,
}

impl MtefEquation {
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
    pub fn equation_native(&self) -> &[u8] {
        &self.equation_native
    }

    /// Headerless MTEF record stream.
    pub fn mtef_payload(&self) -> &[u8] {
        &self.equation_native[EQNOLE_HEADER_SIZE..]
    }

    /// Consume this value and return the complete Equation Native stream.
    pub fn into_equation_native(self) -> Vec<u8> {
        self.equation_native
    }
}

/// Inputs for adding one native MathType/Equation Editor object to a DOC.
#[derive(Clone, Debug)]
pub struct DocMtefEquationWriteOptions {
    pub storage_id: u32,
    pub equation: MtefEquation,
    pub preview: DocPicture,
}

impl DocMtefEquationWriteOptions {
    pub fn new(storage_id: u32, equation: MtefEquation, preview: DocPicture) -> Self {
        Self {
            storage_id,
            equation,
            preview,
        }
    }
}

impl Editor {
    /// Append a native Equation.3 EMBED field with a real PICF bitmap preview.
    pub fn add_mtef_equation(&mut self, options: DocMtefEquationWriteOptions) -> Result<Reference> {
        let mut picture_data = Vec::new();
        write_picture_block(&options.preview, options.storage_id, &mut picture_data)
            .map_err(|error| PackageError::Corrupted(error.to_string()))?;
        let compound_file = equation_compound_file(options.equation)?;
        self.add(WriteOptions {
            storage_id: options.storage_id,
            instruction: format!(" EMBED Equation.3 _{} ", options.storage_id),
            picture_data,
            compound_file,
        })
    }
}

fn equation_compound_file(equation: MtefEquation) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.set_root_clsid(EQUATION_3_CLSID);
    writer.create_stream(&[COMP_OBJ_STREAM], EQUATION_COMP_OBJ)?;
    writer.create_stream(&[OLE_STREAM], EQUATION_OLE)?;
    writer.create_stream(&[OBJ_INFO_STREAM], EQUATION_OBJ_INFO)?;
    writer.create_stream(&[EQUATION_NATIVE_STREAM], equation.equation_native())?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

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
