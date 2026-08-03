//! BIFF8 extended data label records of the Chart Sheet substream
//! (MS-XLS 2.1):
//!
//! - **DataLabExt** (0x086A): the beginning of an extended data label record
//!   collection (MS-XLS 2.4.75).
//! - **DataLabExtContents** (0x086B): the contents of an extended data label
//!   (MS-XLS 2.4.76).
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! data label is composed or rendered.
//!
//! # References
//!
//! - MS-XLS 2.4.75 (DataLabExt), 2.4.76 (DataLabExtContents), 2.5.134
//!   (FrtFlags), 2.5.135 (FrtHeader), 2.5.295 (XLUnicodeStringMin2), 2.5.296
//!   (XLUnicodeStringNoCch)

use super::{XlsError, XlsResult};

/// Record type of the `DataLabExt` record (MS-XLS 2.4.75); also the required
/// `frtHeader.rt` value.
pub(crate) const DATA_LAB_EXT_RECORD_TYPE: u16 = 0x086A;

/// Record type of the `DataLabExtContents` record (MS-XLS 2.4.76); also the
/// required `frtHeader.rt` value.
pub(crate) const DATA_LAB_EXT_CONTENTS_RECORD_TYPE: u16 = 0x086B;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// `FrtFlags` bits that MUST be zero in an `FrtHeader` (MS-XLS 2.5.135):
/// `fFrtRef` and `fFrtAlert`.
const FRT_FLAGS_FORBIDDEN: u16 = 0x0003;
/// `XLUnicodeStringNoCch` flag: `fHighByte` (double-byte characters).
const HIGH_BYTE: u8 = 0x01;

/// Checked cursor for the fixed-width and length-prefixed fields in the
/// extended data-label records.
struct DataLabReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> DataLabReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn read_bytes<const N: usize>(&mut self) -> XlsResult<[u8; N]> {
        let end = self.offset.checked_add(N).ok_or_else(|| {
            invalid(
                DATA_LAB_EXT_CONTENTS_RECORD_TYPE,
                "field offset overflows usize",
            )
        })?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        let value: [u8; N] = bytes.try_into().map_err(|_| XlsError::InvalidLength {
            expected: N,
            found: bytes.len(),
        })?;
        self.offset = end;
        Ok(value)
    }

    fn read_u8(&mut self) -> XlsResult<u8> {
        Ok(u8::from_le_bytes(self.read_bytes()?))
    }

    fn read_u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.read_bytes()?))
    }

    fn read_vec(&mut self, len: usize) -> XlsResult<Vec<u8>> {
        let end = self.offset.checked_add(len).ok_or_else(|| {
            invalid(
                DATA_LAB_EXT_CONTENTS_RECORD_TYPE,
                "field offset overflows usize",
            )
        })?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        self.offset = end;
        Ok(bytes.to_vec())
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.offset)
    }
}

fn invalid(record_type: u16, message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Validate an `FrtHeader` (MS-XLS 2.5.135): the `rt` field and the
/// `fFrtRef`/`fFrtAlert` bits that MUST be zero, returning the raw flags
/// word and reserved bytes.
fn validate_frt_header(
    reader: &mut DataLabReader<'_>,
    record_type: u16,
    name: &str,
) -> XlsResult<(u16, [u8; 8])> {
    let found_type = reader.read_u16()?;
    if found_type != record_type {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.rt mismatch"),
        ));
    }
    let flags = reader.read_u16()?;
    if flags & FRT_FLAGS_FORBIDDEN != 0 {
        return Err(invalid(
            record_type,
            format!("{name} FrtHeader.grbitFrt {flags:#06X} sets fFrtRef or fFrtAlert"),
        ));
    }
    Ok((flags, reader.read_bytes()?))
}

/// Typed `DataLabExt` record content (MS-XLS 2.4.75): the beginning of an
/// extended data label record collection.
///
/// The `frtHeader` reserved bytes (MUST be ignored) are preserved verbatim so
/// the record round-trips unchanged.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsDataLabExt {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
}

impl XlsDataLabExt {
    /// Parse a `DataLabExt` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        // MS-XLS 2.4.75: the record contains only the FrtHeader.
        if data.len() != FRT_HEADER_LEN {
            return Err(XlsError::InvalidLength {
                expected: FRT_HEADER_LEN,
                found: data.len(),
            });
        }
        let mut reader = DataLabReader::new(data);
        let (frt_flags, frt_reserved) =
            validate_frt_header(&mut reader, DATA_LAB_EXT_RECORD_TYPE, "DataLabExt")?;
        Ok(Self {
            frt_flags,
            frt_reserved,
        })
    }

    /// Serialize back to a complete `DataLabExt` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(FRT_HEADER_LEN);
        payload.extend_from_slice(&DATA_LAB_EXT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload
    }

    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    pub fn frt_flags(&self) -> u16 {
        self.frt_flags
    }
}

/// Typed `DataLabExtContents` record content (MS-XLS 2.4.76): the contents
/// of an extended data label.
///
/// The 11 `reserved` flags bits (MUST be ignored) and the raw `rgchSep`
/// string bytes are preserved verbatim so the record round-trips unchanged.
/// The chart-group constraints on `fPercent` and `fBubSizes` (MS-XLS 2.4.76)
/// are cross-record constraints the caller validates.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsDataLabExtContents {
    /// Raw `frtHeader.grbitFrt` bitfield (`fFrtRef`/`fFrtAlert` are zero).
    frt_flags: u16,
    /// `frtHeader.reserved` bytes, preserved verbatim.
    frt_reserved: [u8; 8],
    /// Raw flags word: `fSerName`, `fCatName`, `fValue`, `fPercent`,
    /// `fBubSizes`, and the 11 reserved bits.
    flags: u16,
    /// `rgchSep` character count (`cch`).
    separator_len: u16,
    /// Raw `rgchSep` option flags byte (`fHighByte` and 7 reserved bits).
    separator_flags: u8,
    /// Raw `rgchSep` character bytes, preserved verbatim.
    separator_bytes: Vec<u8>,
}

impl XlsDataLabExtContents {
    /// Flags bit: `fSerName` (series name displayed).
    const FLAG_SERIES_NAME: u16 = 0x0001;
    /// Flags bit: `fCatName` (category name displayed).
    const FLAG_CATEGORY_NAME: u16 = 0x0002;
    /// Flags bit: `fValue` (data value displayed).
    const FLAG_VALUE: u16 = 0x0004;
    /// Flags bit: `fPercent` (percentage displayed).
    const FLAG_PERCENT: u16 = 0x0008;
    /// Flags bit: `fBubSizes` (bubble size displayed).
    const FLAG_BUBBLE_SIZES: u16 = 0x0010;

    /// Parse a `DataLabExtContents` record payload.
    pub fn parse(data: &[u8]) -> XlsResult<Self> {
        // FrtHeader (12) + flags (2) + cch (2); the string option flags and
        // characters follow when cch is greater than zero.
        const MIN_LEN: usize = FRT_HEADER_LEN + 4;
        if data.len() < MIN_LEN {
            return Err(XlsError::InvalidLength {
                expected: MIN_LEN,
                found: data.len(),
            });
        }
        let mut reader = DataLabReader::new(data);
        let (frt_flags, frt_reserved) = validate_frt_header(
            &mut reader,
            DATA_LAB_EXT_CONTENTS_RECORD_TYPE,
            "DataLabExtContents",
        )?;
        let flags = reader.read_u16()?;
        let separator_len = reader.read_u16()?;
        // MS-XLS 2.5.295: st MUST exist if and only if cch is greater than
        // zero; MS-XLS 2.5.296: rgb holds cch or 2*cch bytes.
        if separator_len == 0 {
            if reader.remaining() != 0 {
                return Err(XlsError::InvalidLength {
                    expected: MIN_LEN,
                    found: data.len(),
                });
            }
            return Ok(Self {
                frt_flags,
                frt_reserved,
                flags,
                separator_len,
                separator_flags: 0,
                separator_bytes: Vec::new(),
            });
        }
        if reader.remaining() < 1 {
            return Err(XlsError::InvalidLength {
                expected: MIN_LEN + 1,
                found: data.len(),
            });
        }
        let separator_flags = reader.read_u8()?;
        let byte_len = if separator_flags & HIGH_BYTE != 0 {
            usize::from(separator_len).checked_mul(2).ok_or_else(|| {
                invalid(
                    DATA_LAB_EXT_CONTENTS_RECORD_TYPE,
                    "separator byte length overflows usize",
                )
            })?
        } else {
            usize::from(separator_len)
        };
        if reader.remaining() != byte_len {
            return Err(invalid(
                DATA_LAB_EXT_CONTENTS_RECORD_TYPE,
                format!(
                    "DataLabExtContents rgchSep holds {} bytes for cch {separator_len}",
                    reader.remaining()
                ),
            ));
        }
        let separator_bytes = reader.read_vec(byte_len)?;
        Ok(Self {
            frt_flags,
            frt_reserved,
            flags,
            separator_len,
            separator_flags,
            separator_bytes,
        })
    }

    /// Serialize back to a complete `DataLabExtContents` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::new();
        payload.extend_from_slice(&DATA_LAB_EXT_CONTENTS_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&self.frt_flags.to_le_bytes());
        payload.extend_from_slice(&self.frt_reserved);
        payload.extend_from_slice(&self.flags.to_le_bytes());
        payload.extend_from_slice(&self.separator_len.to_le_bytes());
        if self.separator_len > 0 {
            payload.push(self.separator_flags);
            payload.extend_from_slice(&self.separator_bytes);
        }
        payload
    }

    /// Whether the series name is displayed (`fSerName`).
    pub fn show_series_name(&self) -> bool {
        self.flags & Self::FLAG_SERIES_NAME != 0
    }

    /// Whether the category name or horizontal value is displayed
    /// (`fCatName`).
    pub fn show_category_name(&self) -> bool {
        self.flags & Self::FLAG_CATEGORY_NAME != 0
    }

    /// Whether the data value or vertical value is displayed (`fValue`).
    pub fn show_value(&self) -> bool {
        self.flags & Self::FLAG_VALUE != 0
    }

    /// Whether the percentage of the series sum is displayed (`fPercent`).
    pub fn show_percent(&self) -> bool {
        self.flags & Self::FLAG_PERCENT != 0
    }

    /// Whether the bubble size is displayed (`fBubSizes`).
    pub fn show_bubble_sizes(&self) -> bool {
        self.flags & Self::FLAG_BUBBLE_SIZES != 0
    }

    /// Raw flags word, including the 11 reserved bits.
    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// The decoded `rgchSep` separator string. Compressed characters
    /// (`fHighByte` 0) are single low bytes of UTF-16 code units (MS-XLS
    /// 2.5.296).
    pub fn separator(&self) -> String {
        let units: Vec<u16> = if self.separator_flags & HIGH_BYTE != 0 {
            self.separator_bytes
                .chunks_exact(2)
                .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
                .collect()
        } else {
            self.separator_bytes
                .iter()
                .map(|byte| u16::from(*byte))
                .collect()
        };
        String::from_utf16(&units).unwrap_or_default()
    }

    /// The raw `rgchSep` option flags byte (`fHighByte` and 7 reserved bits).
    pub fn separator_flags(&self) -> u8 {
        self.separator_flags
    }

    /// The raw `rgchSep` character bytes, preserved verbatim.
    pub fn separator_bytes(&self) -> &[u8] {
        &self.separator_bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn frt_header(record_type: u16) -> Vec<u8> {
        let mut data = record_type.to_le_bytes().to_vec();
        data.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        data
    }

    #[test]
    fn data_lab_ext_round_trip() {
        let bytes = frt_header(DATA_LAB_EXT_RECORD_TYPE);
        let parsed = XlsDataLabExt::parse(&bytes).unwrap();
        assert_eq!(parsed.frt_flags(), 0);
        assert_eq!(parsed.to_payload(), bytes);
        // Reserved header bytes round-trip verbatim.
        let mut reserved = bytes.clone();
        reserved[4..FRT_HEADER_LEN].copy_from_slice(&[1, 2, 3, 4, 5, 6, 7, 8]);
        assert_eq!(
            XlsDataLabExt::parse(&reserved).unwrap().to_payload(),
            reserved
        );
    }

    #[test]
    fn data_lab_ext_rejects_malformed_records() {
        let bytes = frt_header(DATA_LAB_EXT_RECORD_TYPE);
        assert!(XlsDataLabExt::parse(&bytes[..11]).is_err());
        assert!(XlsDataLabExt::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&DATA_LAB_EXT_CONTENTS_RECORD_TYPE.to_le_bytes());
        assert!(XlsDataLabExt::parse(&wrong_rt).is_err());
        let mut bad_flags = bytes.clone();
        bad_flags[2..4].copy_from_slice(&0x0001u16.to_le_bytes());
        assert!(XlsDataLabExt::parse(&bad_flags).is_err());
    }

    fn contents_record(flags: u16, separator: &[u8], high_byte: bool) -> Vec<u8> {
        let mut data = frt_header(DATA_LAB_EXT_CONTENTS_RECORD_TYPE);
        data.extend_from_slice(&flags.to_le_bytes());
        let cch = if high_byte {
            separator.len() / 2
        } else {
            separator.len()
        };
        data.extend_from_slice(&(cch as u16).to_le_bytes());
        if cch > 0 {
            data.push(u8::from(high_byte));
            data.extend_from_slice(separator);
        }
        data
    }

    #[test]
    fn contents_round_trip_compressed_and_double_byte() {
        let bytes = contents_record(0x001F, b" | ", false);
        let parsed = XlsDataLabExtContents::parse(&bytes).unwrap();
        assert!(parsed.show_series_name());
        assert!(parsed.show_category_name());
        assert!(parsed.show_value());
        assert!(parsed.show_percent());
        assert!(parsed.show_bubble_sizes());
        assert_eq!(parsed.separator(), " | ");
        assert_eq!(parsed.to_payload(), bytes);

        let wide: Vec<u8> = "; ".encode_utf16().flat_map(u16::to_le_bytes).collect();
        let bytes = contents_record(0x0005, &wide, true);
        let parsed = XlsDataLabExtContents::parse(&bytes).unwrap();
        assert_eq!(parsed.separator(), "; ");
        assert_eq!(parsed.separator_flags(), 0x01);
        assert_eq!(parsed.to_payload(), bytes);

        // Empty separator: cch 0 and no string bytes.
        let bytes = contents_record(0x0000, b"", false);
        let parsed = XlsDataLabExtContents::parse(&bytes).unwrap();
        assert_eq!(parsed.separator(), "");
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn contents_rejects_malformed_records() {
        let bytes = contents_record(0x0001, b"ab", false);
        // Truncated header.
        assert!(XlsDataLabExtContents::parse(&bytes[..15]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = bytes.clone();
        wrong_rt[0..2].copy_from_slice(&DATA_LAB_EXT_RECORD_TYPE.to_le_bytes());
        assert!(XlsDataLabExtContents::parse(&wrong_rt).is_err());
        // cch does not match the string byte count.
        let mut mismatch = contents_record(0, b"ab", false);
        mismatch[14..16].copy_from_slice(&3u16.to_le_bytes());
        assert!(XlsDataLabExtContents::parse(&mismatch).is_err());
        // String bytes present with cch 0.
        let mut extra = contents_record(0, b"", false);
        extra.push(0);
        assert!(XlsDataLabExtContents::parse(&extra).is_err());
        // Double-byte length mismatch.
        let mut wide_mismatch = contents_record(0, &[0x41, 0x00], true);
        wide_mismatch[14..16].copy_from_slice(&2u16.to_le_bytes());
        assert!(XlsDataLabExtContents::parse(&wide_mismatch).is_err());
    }

    #[test]
    fn rejects_every_truncated_payload() {
        let header = frt_header(DATA_LAB_EXT_RECORD_TYPE);
        for length in 0..FRT_HEADER_LEN {
            assert!(
                XlsDataLabExt::parse(&header[..length]).is_err(),
                "DataLabExt length {length}"
            );
        }

        let contents = contents_record(0x0001, b"ab", false);
        for length in 0..contents.len() {
            assert!(
                XlsDataLabExtContents::parse(&contents[..length]).is_err(),
                "DataLabExtContents length {length}"
            );
        }
    }

    #[test]
    fn reader_rejects_offset_overflow() {
        let mut reader = DataLabReader {
            data: &[],
            offset: usize::MAX,
        };

        assert!(reader.read_bytes::<1>().is_err());
        assert!(reader.read_vec(1).is_err());
    }
}
