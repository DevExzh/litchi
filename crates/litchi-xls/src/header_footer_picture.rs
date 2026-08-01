//! BIFF8 `HFPicture` record (0x0866, MS-XLS 2.4.138) of the worksheet
//! substream (MS-XLS 2.1): a picture used by a sheet header or footer.
//!
//! Everything in this module is INERT: the `OfficeArtDgContainer`/
//! `OfficeArtDggContainer` bytes (\[MS-ODRAW\]) are stored verbatim and never
//! parsed, rendered, or anchored. The picture may be continued across
//! multiple `HFPicture` records (the `fContinue` bit); each record is read
//! individually here and the raw bytes are preserved for byte-exact round
//! trips.
//!
//! # References
//!
//! - MS-XLS 2.4.138 (HFPicture), 2.5.134 (FrtFlags), 2.5.135 (FrtHeader)

use super::{XlsError, XlsResult};

/// Record type of the `HFPicture` record (MS-XLS 2.4.138); also the required
/// `frtHeader.rt` value.
pub(crate) const HF_PICTURE_RECORD_TYPE: u16 = 0x0866;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Byte length of the flags + reserved prefix after the `FrtHeader`.
const FLAGS_LEN: usize = 2;

/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Flags bit: `rgDrawing` is an `OfficeArtDgContainer` record.
const FLAG_IS_DRAWING: u8 = 0x01;
/// Flags bit: `rgDrawing` is an `OfficeArtDggContainer` record.
const FLAG_IS_DRAWING_GROUP: u8 = 0x02;
/// Flags bit: this record continues the previous `HFPicture` record.
const FLAG_CONTINUE: u8 = 0x04;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: HF_PICTURE_RECORD_TYPE,
        message: message.into(),
    }
}

/// Typed `HFPicture` record content (MS-XLS 2.4.138): a picture used by a
/// sheet header or footer.
///
/// The `FrtHeader` reserved bytes, the five undefined flags bits, and the
/// `reserved` byte (MUST be ignored) are preserved verbatim so the record
/// round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsHeaderFooterPicture {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// Raw flags byte: `fIsDrawing`, `fIsDrawingGroup`, `fContinue`, and the
    /// five undefined bits, preserved verbatim.
    flags: u8,
    /// `reserved` byte, preserved verbatim.
    reserved: u8,
    /// Opaque `rgDrawing` bytes: an `OfficeArtDgContainer` or
    /// `OfficeArtDggContainer` record as specified in \[MS-ODRAW\], or a
    /// continuation of the previous record's bytes.
    drawing: Vec<u8>,
}

impl XlsHeaderFooterPicture {
    /// Parse an `HFPicture` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < FRT_HEADER_LEN + FLAGS_LEN {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_LEN + FLAGS_LEN,
                found: data.len(),
            });
        }
        if u16::from_le_bytes([data[0], data[1]]) != HF_PICTURE_RECORD_TYPE {
            return Err(invalid("HFPicture FrtHeader.rt mismatch"));
        }
        let frt_flags = u16::from_le_bytes([data[2], data[3]]);
        if frt_flags & FRT_FLAGS_FORBIDDEN != 0 {
            return Err(invalid(format!(
                "HFPicture FrtHeader.grbitFrt {frt_flags:#06X} sets fFrtRef or fFrtAlert"
            )));
        }
        let flags = data[FRT_HEADER_LEN];
        // MS-XLS 2.4.138: exactly one of fIsDrawing / fIsDrawingGroup is set.
        let is_drawing = flags & FLAG_IS_DRAWING != 0;
        let is_drawing_group = flags & FLAG_IS_DRAWING_GROUP != 0;
        if is_drawing == is_drawing_group {
            return Err(invalid(
                "HFPicture must set exactly one of fIsDrawing or fIsDrawingGroup",
            ));
        }
        Ok(Self {
            frt_flags,
            frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
            flags,
            reserved: data[FRT_HEADER_LEN + 1],
            drawing: data[FRT_HEADER_LEN + FLAGS_LEN..].to_vec(),
        })
    }

    /// Serialize back to a complete `HFPicture` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(FRT_HEADER_LEN + FLAGS_LEN + self.drawing.len());
        payload.extend_from_slice(&HF_PICTURE_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.push(self.flags);
        payload.push(self.reserved);
        payload.extend_from_slice(&self.drawing);
        payload
    }

    /// Whether `rgDrawing` is an `OfficeArtDgContainer` record (`fIsDrawing`).
    pub fn is_drawing(&self) -> bool {
        self.flags & FLAG_IS_DRAWING != 0
    }

    /// Whether `rgDrawing` is an `OfficeArtDggContainer` record
    /// (`fIsDrawingGroup`).
    pub fn is_drawing_group(&self) -> bool {
        self.flags & FLAG_IS_DRAWING_GROUP != 0
    }

    /// Whether this record continues the previous `HFPicture` record
    /// (`fContinue`).
    pub fn is_continuation(&self) -> bool {
        self.flags & FLAG_CONTINUE != 0
    }

    /// Raw flags byte, including the five undefined bits.
    pub fn flags(&self) -> u8 {
        self.flags
    }

    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }

    /// The opaque `rgDrawing` bytes.
    pub fn drawing(&self) -> &[u8] {
        &self.drawing
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(flags: u8, drawing: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&HF_PICTURE_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.push(flags);
        data.push(0);
        data.extend_from_slice(drawing);
        data
    }

    #[test]
    fn round_trip_drawing_and_drawing_group() {
        for flags in [FLAG_IS_DRAWING, FLAG_IS_DRAWING_GROUP | FLAG_CONTINUE] {
            let bytes = record(flags, b"drawing-bytes");
            let parsed = XlsHeaderFooterPicture::parse(&bytes).unwrap();
            assert_eq!(parsed.is_drawing(), flags & FLAG_IS_DRAWING != 0);
            assert_eq!(
                parsed.is_drawing_group(),
                flags & FLAG_IS_DRAWING_GROUP != 0
            );
            assert_eq!(parsed.is_continuation(), flags & FLAG_CONTINUE != 0);
            assert_eq!(parsed.drawing(), b"drawing-bytes");
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn preserves_undefined_and_reserved_bytes() {
        // The five undefined flags bits, the reserved byte, and the FrtHeader
        // reserved bytes MUST be ignored but round-trip verbatim.
        let mut bytes = record(FLAG_IS_DRAWING | 0xF8, b"x");
        bytes[13] = 0xAA;
        bytes[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
        let parsed = XlsHeaderFooterPicture::parse(&bytes).unwrap();
        assert_eq!(parsed.flags(), FLAG_IS_DRAWING | 0xF8);
        assert!(parsed.is_drawing());
        assert!(!parsed.is_continuation());
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(FLAG_IS_DRAWING, b"drawing");
        // Truncated.
        assert!(XlsHeaderFooterPicture::parse(&bytes[..13]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x0867u16.to_le_bytes());
        assert!(XlsHeaderFooterPicture::parse(&wrong_rt).is_err());
        // fFrtRef / fFrtAlert set.
        for frt_flags in [0x0001u16, 0x0002] {
            let mut bad = bytes.clone();
            bad[2..4].copy_from_slice(&frt_flags.to_le_bytes());
            assert!(XlsHeaderFooterPicture::parse(&bad).is_err());
        }
        // Neither or both of fIsDrawing / fIsDrawingGroup.
        assert!(XlsHeaderFooterPicture::parse(&record(0x00, b"drawing")).is_err());
        assert!(XlsHeaderFooterPicture::parse(&record(0x03, b"drawing")).is_err());
    }
}
