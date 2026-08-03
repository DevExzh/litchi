//! BIFF8 `SheetExt` record (MS-XLS 2.4.259): sheet tab color and publish
//! state.
//!
//! The record carries a legacy palette tab color (`icvPlain`) and, when the
//! record declares the extended size, a `SheetExtOptional` structure
//! (MS-XLS 2.5.238) with a refreshed tab color, the conditional-formatting
//! calculation flag, the published flag, and a full `CFColor`.

use super::{XlsError, XlsResult};

/// Record type of the `SheetExt` record.
pub(crate) const SHEET_EXT_RECORD_TYPE: u16 = 0x0862;

/// Size in bytes of an `FrtHeader` (MS-XLS 2.5.135).
const FRT_HEADER_LEN: usize = 12;
/// Size in bytes of the record-size field following the `FrtHeader`.
const CB_LEN: usize = 4;
/// Size in bytes of the `icvPlain`/reserved bitfield.
const FLAGS_LEN: usize = 4;
/// Size in bytes of a `SheetExtOptional` structure (MS-XLS 2.5.238).
const OPTIONAL_LEN: usize = 20;
/// Size in bytes of the embedded `CFColor` inside `SheetExtOptional`.
const CF_COLOR_LEN: usize = 16;

/// `cb` value of a `SheetExt` without a `SheetExtOptional`.
const CB_BASE: u32 = 0x14;
/// `cb` value of a `SheetExt` carrying a `SheetExtOptional`.
const CB_WITH_OPTIONAL: u32 = 0x28;
/// Minimum payload length, in bytes, through the base flags.
const CB_BASE_LEN: usize = FRT_HEADER_LEN + CB_LEN + FLAGS_LEN;
/// Full payload length when the optional extension is present.
const CB_WITH_OPTIONAL_LEN: usize = CB_BASE_LEN + OPTIONAL_LEN;

/// `icvPlain` value meaning the sheet tab has no color assigned.
const ICV_NO_COLOR: u32 = 0x7F;
/// Lowest palette index usable as a tab color (MS-XLS 2.5.161 `Icv`).
const ICV_MIN: u8 = 0x08;
/// Highest palette index usable as a tab color (MS-XLS 2.5.161 `Icv`).
const ICV_MAX: u8 = 0x3F;
/// Bitmask selecting the 7-bit `icvPlain`/`icvPlain12` field.
const ICV_MASK: u32 = 0x7F;
/// `SheetExtOptional` bit: conditional-formatting formulas are evaluated.
const COND_FMT_CALC: u32 = 1 << 7;
/// `SheetExtOptional` bit: the sheet is not published.
const NOT_PUBLISHED: u32 = 1 << 8;

fn invalid(message: impl Into<String>) -> XlsError {
    XlsError::InvalidRecord {
        record_type: SHEET_EXT_RECORD_TYPE,
        message: message.into(),
    }
}

/// Checked cursor for the fixed-width fields in a `SheetExt` payload.
///
/// Keeping the offset in one place prevents a future layout change from
/// turning a field boundary into a panic or an overflowing slice range.
struct SheetExtReader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> SheetExtReader<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn read_bytes<const N: usize>(&mut self) -> XlsResult<[u8; N]> {
        let end = self.offset.checked_add(N).ok_or_else(|| {
            invalid(format!(
                "SheetExt field offset {} overflows usize",
                self.offset
            ))
        })?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or(XlsError::InvalidLength {
                expected: end,
                found: self.data.len(),
            })?;
        let value = bytes.try_into().map_err(|_| XlsError::InvalidLength {
            expected: N,
            found: bytes.len(),
        })?;
        self.offset = end;
        Ok(value)
    }

    fn read_u16(&mut self) -> XlsResult<u16> {
        Ok(u16::from_le_bytes(self.read_bytes()?))
    }

    fn read_u32(&mut self) -> XlsResult<u32> {
        Ok(u32::from_le_bytes(self.read_bytes()?))
    }
}

/// Decode a 7-bit `Icv` field into a validated palette index.
fn decode_icv(raw: u32, record_type: u16) -> XlsResult<Option<u8>> {
    let value = raw & ICV_MASK;
    if value == ICV_NO_COLOR {
        return Ok(None);
    }
    let index = u8::try_from(value).map_err(|_| XlsError::InvalidRecord {
        record_type,
        message: format!("sheet tab color index {value:#04X} does not fit in an Icv"),
    })?;
    if (ICV_MIN..=ICV_MAX).contains(&index) {
        Ok(Some(index))
    } else {
        Err(XlsError::InvalidRecord {
            record_type,
            message: format!("sheet tab color index {index:#04X} is outside the Icv palette"),
        })
    }
}

/// The optional extension of a `SheetExt` record (MS-XLS 2.5.238).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSheetExtOptional {
    /// Refreshed tab color palette index (`icvPlain12`); the base record's
    /// `icvPlain` takes precedence when the two disagree.
    tab_color_12: Option<u8>,
    /// Whether conditional-formatting formulas are evaluated (`fCondFmtCalc`).
    cond_fmt_calc: bool,
    /// Whether the sheet is excluded from publishing (`fNotPublished`).
    not_published: bool,
    /// Full `CFColor` tab color, preserved verbatim.
    color: [u8; CF_COLOR_LEN],
}

impl XlsSheetExtOptional {
    /// Refreshed tab color palette index, when assigned.
    pub fn tab_color_12(&self) -> Option<u8> {
        self.tab_color_12
    }

    /// Whether conditional-formatting formulas are evaluated.
    pub fn cond_fmt_calc(&self) -> bool {
        self.cond_fmt_calc
    }

    /// Whether the sheet is excluded from publishing.
    pub fn not_published(&self) -> bool {
        self.not_published
    }

    /// Full `CFColor` tab color bytes.
    pub fn color(&self) -> &[u8; CF_COLOR_LEN] {
        &self.color
    }
}

/// Typed `SheetExt` record content (MS-XLS 2.4.259).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSheetExt {
    /// Sheet tab color palette index (`icvPlain`); `None` when unassigned.
    tab_color: Option<u8>,
    /// Optional extension, present iff the record declares the extended size.
    optional: Option<XlsSheetExtOptional>,
}

impl XlsSheetExt {
    /// Sheet tab color as a palette index, when assigned.
    pub fn tab_color(&self) -> Option<u8> {
        self.tab_color
    }

    /// The optional extension, when present.
    pub fn optional(&self) -> Option<&XlsSheetExtOptional> {
        self.optional.as_ref()
    }

    /// Parse a `SheetExt` record payload.
    pub(crate) fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < CB_BASE_LEN {
            return Err(XlsError::InvalidLength {
                expected: CB_BASE_LEN,
                found: data.len(),
            });
        }
        let mut reader = SheetExtReader::new(data);
        if reader.read_u16()? != SHEET_EXT_RECORD_TYPE {
            return Err(invalid("SheetExt FrtHeader.rt mismatch"));
        }
        let _ = reader.read_bytes::<{ FRT_HEADER_LEN - 2 }>()?;
        let cb = reader.read_u32()?;
        let expected_len = match cb {
            CB_BASE => CB_BASE_LEN,
            CB_WITH_OPTIONAL => CB_WITH_OPTIONAL_LEN,
            _ => return Err(invalid("SheetExt declares an unsupported record size")),
        };
        if data.len() != expected_len {
            return Err(XlsError::InvalidLength {
                expected: expected_len,
                found: data.len(),
            });
        }
        let flags = reader.read_u32()?;
        let tab_color = decode_icv(flags, SHEET_EXT_RECORD_TYPE)?;
        let optional = if cb == CB_WITH_OPTIONAL {
            let ext_flags = reader.read_u32()?;
            Some(XlsSheetExtOptional {
                tab_color_12: decode_icv(ext_flags, SHEET_EXT_RECORD_TYPE)?,
                cond_fmt_calc: ext_flags & COND_FMT_CALC != 0,
                not_published: ext_flags & NOT_PUBLISHED != 0,
                color: reader.read_bytes()?,
            })
        } else {
            None
        };
        Ok(Self {
            tab_color,
            optional,
        })
    }

    /// Serialize back to a complete `SheetExt` record payload.
    pub(crate) fn to_payload(&self) -> Vec<u8> {
        let (cb, capacity) = if self.optional.is_some() {
            (CB_WITH_OPTIONAL, CB_WITH_OPTIONAL_LEN)
        } else {
            (CB_BASE, CB_BASE_LEN)
        };
        let mut payload = Vec::with_capacity(capacity);
        // FrtHeader: rt, grbitFrt (0), reserved (0).
        payload.extend_from_slice(&SHEET_EXT_RECORD_TYPE.to_le_bytes());
        payload.extend_from_slice(&[0; FRT_HEADER_LEN - 2]);
        payload.extend_from_slice(&cb.to_le_bytes());
        payload.extend_from_slice(&self.tab_color.map_or(ICV_NO_COLOR, u32::from).to_le_bytes());
        if let Some(optional) = &self.optional {
            let mut flags = optional.tab_color_12.map_or(ICV_NO_COLOR, u32::from);
            if optional.cond_fmt_calc {
                flags |= COND_FMT_CALC;
            }
            if optional.not_published {
                flags |= NOT_PUBLISHED;
            }
            payload.extend_from_slice(&flags.to_le_bytes());
            payload.extend_from_slice(&optional.color);
        }
        payload
    }

    /// Construct a record carrying only a base tab color (writer path).
    pub(crate) fn from_tab_color(tab_color: Option<u8>) -> Self {
        Self {
            tab_color,
            optional: None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn base_record(icv: u32) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&SHEET_EXT_RECORD_TYPE.to_le_bytes());
        data.extend_from_slice(&[0; 10]);
        data.extend_from_slice(&CB_BASE.to_le_bytes());
        data.extend_from_slice(&icv.to_le_bytes());
        data
    }

    #[test]
    fn parses_base_record_with_and_without_tab_color() {
        let colored = XlsSheetExt::parse(&base_record(0x0A)).unwrap();
        assert_eq!(colored.tab_color(), Some(0x0A));
        assert!(colored.optional().is_none());
        assert_eq!(colored.to_payload(), base_record(0x0A));

        let plain = XlsSheetExt::parse(&base_record(ICV_NO_COLOR)).unwrap();
        assert_eq!(plain.tab_color(), None);
        assert_eq!(plain.to_payload(), base_record(ICV_NO_COLOR));
    }

    #[test]
    fn parses_optional_extension() {
        let mut data = base_record(0x0A);
        data[12..16].copy_from_slice(&CB_WITH_OPTIONAL.to_le_bytes());
        let ext_flags = 0x0B | COND_FMT_CALC | NOT_PUBLISHED;
        data.extend_from_slice(&ext_flags.to_le_bytes());
        data.extend_from_slice(&[0x5A; CF_COLOR_LEN]);

        let parsed = XlsSheetExt::parse(&data).unwrap();
        let optional = parsed.optional().unwrap();
        assert_eq!(optional.tab_color_12(), Some(0x0B));
        assert!(optional.cond_fmt_calc());
        assert!(optional.not_published());
        assert_eq!(optional.color(), &[0x5A; CF_COLOR_LEN]);
        assert_eq!(parsed.to_payload(), data);
    }

    #[test]
    fn rejects_malformed_records() {
        // Truncated.
        assert!(XlsSheetExt::parse(&base_record(0x0A)[..10]).is_err());
        // Wrong FrtHeader.rt.
        let mut wrong_rt = base_record(0x0A);
        wrong_rt[0..2].copy_from_slice(&0x0863u16.to_le_bytes());
        assert!(XlsSheetExt::parse(&wrong_rt).is_err());
        // Unsupported cb.
        let mut wrong_cb = base_record(0x0A);
        wrong_cb[12..16].copy_from_slice(&24u32.to_le_bytes());
        assert!(XlsSheetExt::parse(&wrong_cb).is_err());
        // Out-of-palette color index.
        assert!(XlsSheetExt::parse(&base_record(0x05)).is_err());
        // Declared size disagreeing with the payload length.
        let mut padded = base_record(0x0A);
        padded.push(0);
        assert!(XlsSheetExt::parse(&padded).is_err());
    }

    #[test]
    fn rejects_truncation_at_every_fixed_width_boundary() {
        let mut data = base_record(0x0A);
        data[12..16].copy_from_slice(&CB_WITH_OPTIONAL.to_le_bytes());
        data.extend_from_slice(&0x0Bu32.to_le_bytes());
        data.extend_from_slice(&[0x5A; CF_COLOR_LEN]);

        for length in 0..data.len() {
            assert!(
                XlsSheetExt::parse(&data[..length]).is_err(),
                "truncated SheetExt payload of length {length} was accepted"
            );
        }
    }

    #[test]
    fn checked_reader_rejects_offset_overflow() {
        let mut reader = SheetExtReader {
            data: &[],
            offset: usize::MAX,
        };

        let error = reader.read_bytes::<2>().unwrap_err();
        assert!(matches!(
            error,
            XlsError::InvalidRecord { message, .. }
                if message.contains("offset") && message.contains("overflows")
        ));
        assert_eq!(reader.offset, usize::MAX);
    }
}
