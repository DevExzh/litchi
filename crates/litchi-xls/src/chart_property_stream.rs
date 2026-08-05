//! BIFF8 chart property-stream future records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **ShapePropsStream** (0x08A4): shape formatting properties for chart
//!   elements, as an XML stream (MS-XLS 2.4.258).
//! - **TextPropsStream** (0x08A5): additional text properties for chart text,
//!   as an XML stream (MS-XLS 2.4.325).
//! - **RichTextStream** (0x08A6): additional rich text properties for chart
//!   text, as an XML stream (MS-XLS 2.4.218).
//!
//! All three records may be ignored without loss of functionality (except
//! their additional properties), so everything in this module is INERT: the
//! XML stream bytes and `dwChecksum` values are stored verbatim and the XML
//! is never parsed, validated, or applied. The checksum inputs include data
//! owned by other records (Text/Font/LineFormat/...), so checksums are not
//! recomputed here.
//!
//! # References
//!
//! - MS-XLS 2.4.218 (RichTextStream), 2.4.258 (ShapePropsStream), 2.4.325
//!   (TextPropsStream), 2.5.134 (FrtFlags), 2.5.135 (FrtHeader)

use super::{Error, Result};

/// Record type of the `ShapePropsStream` record (MS-XLS 2.4.258); also the
/// required `frtHeader.rt` value.
pub(crate) const SHAPE_PROPS_STREAM_RECORD_TYPE: u16 = 0x08A4;

/// Record type of the `TextPropsStream` record (MS-XLS 2.4.325); also the
/// required `frtHeader.rt` value.
pub(crate) const TEXT_PROPS_STREAM_RECORD_TYPE: u16 = 0x08A5;

/// Record type of the `RichTextStream` record (MS-XLS 2.4.218); also the
/// required `frtHeader.rt` value.
pub(crate) const RICH_TEXT_STREAM_RECORD_TYPE: u16 = 0x08A6;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// Byte length of the `cb` field.
const CB_LEN: usize = 4;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Validate an `FrtHeader` (MS-XLS 2.5.135): the `rt` field and the
/// `fFrtRef`/`fFrtAlert` bits that MUST be zero.
fn validate_frt_header(data: &[u8], record_type: u16, name: &str) -> Result<u16> {
    if u16::from_le_bytes([data[0], data[1]]) != record_type {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.rt mismatch"),
        ));
    }
    let flags = u16::from_le_bytes([data[2], data[3]]);
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        ));
    }
    Ok(flags)
}

/// Read the trailing `cb` + `rgb` pair shared by the three property-stream
/// records, validating that `cb` is the exact `rgb` length.
fn read_stream(data: &[u8], offset: usize, record_type: u16, name: &str) -> Result<Vec<u8>> {
    let declared = u32::from_le_bytes(
        data[offset..offset + CB_LEN]
            .try_into()
            .expect("length checked"),
    );
    let stream = data[offset + CB_LEN..].to_vec();
    if stream.len() as u64 != u64::from(declared) {
        return Err(invalid(
            record_type,
            format!(
                "{name} cb {declared} does not match its rgb size {}",
                stream.len()
            ),
        ));
    }
    Ok(stream)
}

fn write_stream(checksum: u32, stream: &[u8], output: &mut Vec<u8>) {
    output.extend_from_slice(&checksum.to_le_bytes());
    output.extend_from_slice(&(stream.len() as u32).to_le_bytes());
    output.extend_from_slice(stream);
}

/// Typed `ShapePropsStream` record content (MS-XLS 2.4.258): shape formatting
/// properties for chart elements, as an opaque XML stream.
///
/// The `frtHeader` reserved bytes and the `unused` field (MUST be ignored)
/// are preserved verbatim so the record round-trips unchanged.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapePropsStream {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// `wObjContext`: the chart element the properties apply to. The meaning
    /// is defined by the containing record rule (AXS, CRT, SS, FRAME, or
    /// DROPBAR; MS-XLS 2.4.258).
    object_context: u16,
    /// `unused` field, preserved verbatim.
    unused: u16,
    /// Raw `dwChecksum` of the shape formatting properties, preserved verbatim.
    checksum: u32,
    /// Opaque `rgb` bytes: the XML representation of the shape formatting
    /// properties (ECMA-376 Part 1, 21.2.2.197).
    stream: Vec<u8>,
}

impl ShapePropsStream {
    /// Byte length of the fixed prefix: `FrtHeader` (12) + `wObjContext` (2) +
    /// `unused` (2) + `dwChecksum` (4) + `cb` (4).
    const HEADER_LEN: usize = FRT_HEADER_LEN + 2 + 2 + 4 + CB_LEN;

    /// Parse a `ShapePropsStream` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < Self::HEADER_LEN {
            return Err(Error::InvalidLength {
                expected: Self::HEADER_LEN,
                found: data.len(),
            });
        }
        let frt_flags =
            validate_frt_header(data, SHAPE_PROPS_STREAM_RECORD_TYPE, "ShapePropsStream")?;
        let stream = read_stream(
            data,
            Self::HEADER_LEN - CB_LEN,
            SHAPE_PROPS_STREAM_RECORD_TYPE,
            "ShapePropsStream",
        )?;
        Ok(Self {
            frt_flags,
            frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
            object_context: u16::from_le_bytes([data[12], data[13]]),
            unused: u16::from_le_bytes([data[14], data[15]]),
            checksum: u32::from_le_bytes(data[16..20].try_into().expect("length checked")),
            stream,
        })
    }

    /// Serialize back to a complete `ShapePropsStream` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(Self::HEADER_LEN + self.stream.len());
        payload.extend_from_slice(&SHAPE_PROPS_STREAM_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.extend_from_slice(&self.object_context.to_le_bytes());
        payload.extend_from_slice(&self.unused.to_le_bytes());
        write_stream(self.checksum, &self.stream, &mut payload);
        payload
    }

    /// The chart element the properties apply to (`wObjContext`); the meaning
    /// depends on the containing record rule (MS-XLS 2.4.258).
    pub fn object_context(&self) -> u16 {
        self.object_context
    }

    /// Raw `unused` field value.
    pub fn unused(&self) -> u16 {
        self.unused
    }

    /// Raw `dwChecksum` value, preserved verbatim.
    pub fn checksum(&self) -> u32 {
        self.checksum
    }

    /// The opaque XML property-stream bytes (`rgb`).
    pub fn stream(&self) -> &[u8] {
        &self.stream
    }
}

macro_rules! text_property_stream {
    (
        $(#[$meta:meta])*
        $name:ident, $record_type:expr, $spec:literal
    ) => {
        $(#[$meta])*
        ///
        /// The `frtHeader` reserved bytes (MUST be ignored) are preserved
        /// verbatim so the record round-trips unchanged.
        #[derive(Debug, Clone, PartialEq, Eq)]
        pub struct $name {
            /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
            frt_flags: u16,
            /// `frtHeader.reserved` bytes, preserved verbatim.
            frt_reserved: [u8; 8],
            /// Raw `dwChecksum` of the text properties, preserved verbatim.
            checksum: u32,
            /// Opaque `rgb` bytes: the XML representation of the text
            /// properties (ECMA-376 Part 1, 21.2.2.216).
            stream: Vec<u8>,
        }

        impl $name {
            /// Byte length of the fixed prefix: `FrtHeader` (12) +
            /// `dwChecksum` (4) + `cb` (4).
            const HEADER_LEN: usize = FRT_HEADER_LEN + 4 + CB_LEN;

            /// Parse the record payload.
            pub fn parse(data: &[u8]) -> Result<Self> {
                if data.len() < Self::HEADER_LEN {
                    return Err(Error::InvalidLength {
                        expected: Self::HEADER_LEN,
                        found: data.len(),
                    });
                }
                let frt_flags = validate_frt_header(data, $record_type, $spec)?;
                let stream = read_stream(
                    data,
                    Self::HEADER_LEN - CB_LEN,
                    $record_type,
                    $spec,
                )?;
                Ok(Self {
                    frt_flags,
                    frt_reserved: data[4..FRT_HEADER_LEN].try_into().expect("length checked"),
                    checksum: u32::from_le_bytes(
                        data[FRT_HEADER_LEN..FRT_HEADER_LEN + 4]
                            .try_into()
                            .expect("length checked"),
                    ),
                    stream,
                })
            }

            /// Serialize back to a complete record payload.
            pub fn to_payload(&self) -> Vec<u8> {
                let mut payload = Vec::with_capacity(Self::HEADER_LEN + self.stream.len());
                payload.extend_from_slice(&$record_type.to_le_bytes());
                payload.extend_from_slice(&self.frt_flags.to_le_bytes());
                payload.extend_from_slice(&self.frt_reserved);
                write_stream(self.checksum, &self.stream, &mut payload);
                payload
            }

            /// Raw `dwChecksum` value, preserved verbatim.
            pub fn checksum(&self) -> u32 {
                self.checksum
            }

            /// The opaque XML property-stream bytes (`rgb`).
            pub fn stream(&self) -> &[u8] {
                &self.stream
            }
        }
    };
}

text_property_stream! {
    /// Typed `TextPropsStream` record content (MS-XLS 2.4.325): additional
    /// text properties for chart text, as an opaque XML stream.
    TextPropsStream, TEXT_PROPS_STREAM_RECORD_TYPE, "TextPropsStream"
}

text_property_stream! {
    /// Typed `RichTextStream` record content (MS-XLS 2.4.218): additional
    /// rich text properties for chart text, as an opaque XML stream.
    RichTextStream, RICH_TEXT_STREAM_RECORD_TYPE, "RichTextStream"
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(record_type: u16, middle: &[u8], checksum: u32, stream: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&record_type.to_le_bytes());
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data.extend_from_slice(middle);
        data.extend_from_slice(&checksum.to_le_bytes());
        data.extend_from_slice(&(stream.len() as u32).to_le_bytes());
        data.extend_from_slice(stream);
        data
    }

    #[test]
    fn shape_props_stream_round_trip() {
        let bytes = record(
            SHAPE_PROPS_STREAM_RECORD_TYPE,
            &[0x02, 0x00, 0xAA, 0xBB],
            0x1234_5678,
            b"<c:spPr/>",
        );
        let parsed = ShapePropsStream::parse(&bytes).unwrap();
        assert_eq!(parsed.object_context(), 0x0002);
        assert_eq!(parsed.unused(), 0xBBAA);
        assert_eq!(parsed.checksum(), 0x1234_5678);
        assert_eq!(parsed.stream(), b"<c:spPr/>");
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn text_and_rich_text_streams_round_trip() {
        let text = record(TEXT_PROPS_STREAM_RECORD_TYPE, &[], 7, b"<c:txPr/>");
        let parsed = TextPropsStream::parse(&text).unwrap();
        assert_eq!(parsed.checksum(), 7);
        assert_eq!(parsed.stream(), b"<c:txPr/>");
        assert_eq!(parsed.to_payload(), text);

        let rich = record(RICH_TEXT_STREAM_RECORD_TYPE, &[], 0, b"<c:rich/>");
        let parsed = RichTextStream::parse(&rich).unwrap();
        assert_eq!(parsed.stream(), b"<c:rich/>");
        assert_eq!(parsed.to_payload(), rich);

        // Empty streams are legal.
        let empty = record(TEXT_PROPS_STREAM_RECORD_TYPE, &[], 0, b"");
        assert_eq!(TextPropsStream::parse(&empty).unwrap().to_payload(), empty);
    }

    #[test]
    fn rejects_malformed_streams() {
        let bytes = record(SHAPE_PROPS_STREAM_RECORD_TYPE, &[0; 4], 0, b"xml");
        // Truncated headers.
        assert!(ShapePropsStream::parse(&bytes[..23]).is_err());
        assert!(TextPropsStream::parse(&bytes[..19]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&0x08A5u16.to_le_bytes());
        assert!(ShapePropsStream::parse(&wrong_rt).is_err());
        // fFrtRef / fFrtAlert set.
        let mut bad_flags = bytes.clone();
        bad_flags[2..4].copy_from_slice(&0x0003u16.to_le_bytes());
        assert!(ShapePropsStream::parse(&bad_flags).is_err());
        // cb does not match the rgb size.
        let mut mismatch = bytes.clone();
        let cb_offset = bytes.len() - 4 - 3;
        mismatch[cb_offset..cb_offset + 4].copy_from_slice(&99u32.to_le_bytes());
        assert!(ShapePropsStream::parse(&mismatch).is_err());
        let text = record(TEXT_PROPS_STREAM_RECORD_TYPE, &[], 0, b"xml");
        let mut text_mismatch = text.clone();
        let text_cb_offset = text.len() - 4 - 3;
        text_mismatch[text_cb_offset..text_cb_offset + 4].copy_from_slice(&1u32.to_le_bytes());
        assert!(TextPropsStream::parse(&text_mismatch).is_err());
    }
}
