//! BIFF8 `BkHim` record (0x00E9, MS-XLS 2.4.19) of the worksheet substream
//! (MS-XLS 2.1): image data for a sheet background.
//!
//! Everything in this module is INERT: the image bytes are stored verbatim
//! and never decoded, rendered, or applied as a background. In particular,
//! the 0x000E "native format" payload cannot be directly processed (MS-XLS
//! 2.4.19) and is preserved as opaque bytes in all cases.
//!
//! # References
//!
//! - MS-XLS 2.4.19 (BkHim)

use super::{XlsError, XlsResult};

/// Record type of the `BkHim` record (MS-XLS 2.4.19).
pub(crate) const BK_HIM_RECORD_TYPE: u16 = 0x00E9;

/// Byte length of the fixed `BkHim` prefix: `cf` + `reserved` + `lcb`.
const HEADER_LEN: usize = 8;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: BK_HIM_RECORD_TYPE,
        message: message.into(),
    }
}

/// The `cf` image format of a `BkHim` record (MS-XLS 2.4.19).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u16)]
pub enum XlsBackgroundImageFormat {
    /// 0x0009: bitmap format as described in MSDN-BMP.
    Bitmap = 0x0009,
    /// 0x000E: native format of another application; the bytes cannot be
    /// directly processed.
    Native = 0x000E,
}

impl XlsBackgroundImageFormat {
    fn parse(value: u16) -> XlsResult<Self> {
        match value {
            0x0009 => Ok(Self::Bitmap),
            0x000E => Ok(Self::Native),
            other => Err(invalid(format!(
                "BkHim cf {other:#06X} is not a defined image format"
            ))),
        }
    }
}

/// Typed `BkHim` record content (MS-XLS 2.4.19): sheet background image data.
///
/// The `reserved` field (MUST be 0x0001 and MUST be ignored) is preserved
/// verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsBackgroundImage {
    /// Image format (`cf`).
    format: XlsBackgroundImageFormat,
    /// Raw `reserved` field, preserved verbatim.
    reserved: u16,
    /// Opaque `imageBlob` bytes. Guaranteed non-empty (`lcb` >= 1).
    image: Vec<u8>,
}

impl XlsBackgroundImage {
    /// Parse a `BkHim` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < HEADER_LEN {
            return Err(XlsError::InvalidLength {
                expected: HEADER_LEN,
                found: data.len(),
            });
        }
        let format = XlsBackgroundImageFormat::parse(u16::from_le_bytes([data[0], data[1]]))?;
        let reserved = u16::from_le_bytes([data[2], data[3]]);
        let declared = i32::from_le_bytes(data[4..8].try_into().expect("length checked"));
        // MS-XLS 2.4.19: lcb MUST be greater than or equal to 1.
        if declared < 1 {
            return Err(invalid(format!(
                "BkHim lcb {declared} is not greater than or equal to 1"
            )));
        }
        let declared = usize::try_from(declared).expect("positive i32 fits usize");
        let image_len = data.len() - HEADER_LEN;
        if image_len != declared {
            return Err(invalid(format!(
                "BkHim lcb {declared} does not match its imageBlob size {image_len}"
            )));
        }
        Ok(Self {
            format,
            reserved,
            image: data[HEADER_LEN..].to_vec(),
        })
    }

    /// Serialize back to a complete `BkHim` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(HEADER_LEN + self.image.len());
        payload.extend_from_slice(&(self.format as u16).to_le_bytes());
        payload.extend_from_slice(&self.reserved.to_le_bytes());
        payload.extend_from_slice(&(self.image.len() as i32).to_le_bytes());
        payload.extend_from_slice(&self.image);
        payload
    }

    /// Image format (`cf`).
    pub fn format(&self) -> XlsBackgroundImageFormat {
        self.format
    }

    /// Raw `reserved` field value.
    pub fn reserved(&self) -> u16 {
        self.reserved
    }

    /// The opaque `imageBlob` bytes.
    pub fn image(&self) -> &[u8] {
        &self.image
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(cf: u16, reserved: u16, image: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&cf.to_le_bytes());
        data.extend_from_slice(&reserved.to_le_bytes());
        data.extend_from_slice(&(image.len() as i32).to_le_bytes());
        data.extend_from_slice(image);
        data
    }

    #[test]
    fn round_trip_both_formats() {
        for (cf, expected) in [
            (0x0009, XlsBackgroundImageFormat::Bitmap),
            (0x000E, XlsBackgroundImageFormat::Native),
        ] {
            let bytes = record(cf, 0x0001, b"\x42\x4Dimage-bytes");
            let parsed = XlsBackgroundImage::parse(&bytes).unwrap();
            assert_eq!(parsed.format(), expected);
            assert_eq!(parsed.reserved(), 0x0001);
            assert_eq!(parsed.image(), b"\x42\x4Dimage-bytes");
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn preserves_nonstandard_reserved_value() {
        // reserved MUST be 0x0001 and MUST be ignored; it round-trips verbatim.
        let bytes = record(0x0009, 0x7F7F, b"x");
        let parsed = XlsBackgroundImage::parse(&bytes).unwrap();
        assert_eq!(parsed.reserved(), 0x7F7F);
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x0009, 0x0001, b"image");
        // Truncated header.
        assert!(XlsBackgroundImage::parse(&bytes[..7]).is_err());
        // Undefined cf value.
        assert!(XlsBackgroundImage::parse(&record(0x0002, 0x0001, b"image")).is_err());
        // lcb smaller than 1.
        assert!(XlsBackgroundImage::parse(&record(0x0009, 0x0001, b"")).is_err());
        let mut negative = record(0x0009, 0x0001, b"image");
        negative[4..8].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(XlsBackgroundImage::parse(&negative).is_err());
        // lcb does not match the imageBlob size.
        let mut mismatch = record(0x0009, 0x0001, b"image");
        mismatch[4..8].copy_from_slice(&4i32.to_le_bytes());
        assert!(XlsBackgroundImage::parse(&mismatch).is_err());
    }
}
