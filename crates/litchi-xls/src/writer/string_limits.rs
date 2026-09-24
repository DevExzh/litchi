//! Checked lengths of the BIFF8 string fields the writer emits.
//!
//! Each limit counts UTF-16 code units, which is what a BIFF8 `cch` counts. A
//! string that does not fit is refused with [`Error::StringTooLong`] before any
//! output is produced; the writer never truncates one.

use crate::{Error, Result};

/// An SST entry, `XLUnicodeRichExtendedString`, has a 16-bit `cch`
/// ([MS-XLS] 2.5.293). The workbook editor refuses longer shared strings too.
pub(crate) const SHARED_STRING_UNITS: usize = 0xFFFF;

/// `BoundSheet8.stName.cch` is 1 through 31 ([MS-XLS] 2.4.28).
pub(crate) const WORKSHEET_NAME_UNITS: usize = 31;

/// `Lbl.cch`, the length of a defined name, is one byte ([MS-XLS] 2.4.150).
pub(crate) const DEFINED_NAME_UNITS: usize = 255;

/// `NameCmt.cchComment`, the length of a defined name's comment, is at most
/// 0xFF ([MS-XLS] 2.4.176).
pub(crate) const NAME_COMMENT_UNITS: usize = 255;

/// `PtgStr.string` is a `ShortXLUnicodeString` with a one-byte `cch`
/// ([MS-XLS] 2.5.198.89, 2.5.240).
pub(crate) const FORMULA_STRING_UNITS: usize = 255;

/// `Format.stFormat` holds 1 through 255 characters ([MS-XLS] 2.4.126); the
/// reader refuses longer format strings.
pub(crate) const NUMBER_FORMAT_UNITS: usize = 255;

/// `Font.fontName` holds 1 through 31 characters ([MS-XLS] 2.4.122).
pub(crate) const FONT_NAME_UNITS: usize = 31;

/// `AFDOperStr.cch`, the length of an AutoFilter string operand, is one byte
/// and at least 1 ([MS-XLS] 2.5.8).
pub(crate) const AUTOFILTER_STRING_UNITS: usize = 255;

/// The payload of one BIFF8 record ([MS-XLS] 2.1.4). `HLink`, `XFExt`,
/// `StyleExt`, `CRN` and `SXString` have no continuation records, so their
/// whole payload has to fit.
pub(crate) const RECORD_PAYLOAD_BYTES: usize = litchi_biff::MAX_RECORD_BYTES;

/// `HLink` for an internal link: `Ref8U`, the CLSID, the stream version, the
/// flags and the location's `HyperlinkString` length, before its characters.
pub(crate) const INTERNAL_HYPERLINK_FIXED_BYTES: usize = 8 + 16 + 4 + 4 + 4;

/// `HLink` for a URL: as for an internal link, plus the URL moniker CLSID.
pub(crate) const WEB_HYPERLINK_FIXED_BYTES: usize = INTERNAL_HYPERLINK_FIXED_BYTES + 16;

/// The longest internal hyperlink target: its UTF-16 code units and the NUL
/// that terminates them fill the rest of one `HLink` record.
pub(crate) const INTERNAL_HYPERLINK_UNITS: usize =
    (RECORD_PAYLOAD_BYTES - INTERNAL_HYPERLINK_FIXED_BYTES) / 2 - 1;

/// The longest URL hyperlink target, counted as for an internal one.
pub(crate) const WEB_HYPERLINK_UNITS: usize =
    (RECORD_PAYLOAD_BYTES - WEB_HYPERLINK_FIXED_BYTES) / 2 - 1;

/// `SxView.cchTableName` is at most 0xFF ([MS-XLS] 2.4.313).
pub(crate) const PIVOT_TABLE_NAME_UNITS: usize = 255;

/// `SxView.cchDataName` is 1 through 0xFE ([MS-XLS] 2.4.313).
pub(crate) const PIVOT_DATA_FIELD_NAME_UNITS: usize = 254;

/// `Sxvd.cchName` is 1 through 255 when present ([MS-XLS] 2.4.309).
pub(crate) const PIVOT_FIELD_NAME_UNITS: usize = 255;

/// `SXVI.cchName` is at most 254 when present ([MS-XLS] 2.4.312).
pub(crate) const PIVOT_ITEM_NAME_UNITS: usize = 254;

/// `SXDI.cchName` is 1 through 0xFF when present ([MS-XLS] 2.4.278).
pub(crate) const PIVOT_DATA_ITEM_NAME_UNITS: usize = 255;

/// `SXFDB.stFieldName` is at most 255 characters ([MS-XLS] 2.4.283).
pub(crate) const PIVOT_CACHE_FIELD_NAME_UNITS: usize = 255;

/// `SXString`: the cache item's `cch`, its option byte and its characters
/// fill one record ([MS-XLS] 2.4.304). A string is stored one byte per
/// character only when it is ASCII.
pub(crate) const fn pivot_cache_string_units(ascii: bool) -> usize {
    let characters = RECORD_PAYLOAD_BYTES - 3;
    if ascii { characters } else { characters / 2 }
}

/// Converts a record payload length to the record header's `u16`, refusing
/// with [`Error::RecordTooLong`] a payload no BIFF8 record can hold.
pub(crate) fn record_len(record: &'static str, bytes: usize) -> Result<u16> {
    if bytes > RECORD_PAYLOAD_BYTES {
        return Err(Error::RecordTooLong {
            record,
            bytes,
            limit: RECORD_PAYLOAD_BYTES,
        });
    }
    u16::try_from(bytes).map_err(|_error| Error::RecordTooLong {
        record,
        bytes,
        limit: RECORD_PAYLOAD_BYTES,
    })
}

/// The number of UTF-16 code units in `value`.
pub(crate) fn utf16_len(value: &str) -> usize {
    if value.is_ascii() {
        value.len()
    } else {
        value.encode_utf16().count()
    }
}

/// Returns the UTF-16 length of `value`, or [`Error::StringTooLong`] naming
/// `field` when it is longer than `limit`.
pub(crate) fn checked_utf16_len(value: &str, limit: usize, field: &'static str) -> Result<usize> {
    let utf16_units = utf16_len(value);
    if utf16_units > limit {
        return Err(Error::StringTooLong {
            field,
            utf16_units,
            limit,
        });
    }
    Ok(utf16_units)
}

/// Refuses `value` with [`Error::StringTooLong`] when it is longer than
/// `limit` UTF-16 code units.
///
/// No UTF-8 byte encodes more than one UTF-16 code unit, so a string of at
/// most `limit` bytes fits without being counted.
pub(crate) fn ensure_utf16_len_within(
    value: &str,
    limit: usize,
    field: &'static str,
) -> Result<()> {
    if value.len() <= limit {
        return Ok(());
    }
    checked_utf16_len(value, limit, field).map(drop)
}

/// Converts a length already checked against a one-byte field to that byte.
pub(crate) fn u8_len(units: usize, field: &'static str) -> Result<u8> {
    u8::try_from(units).map_err(|_error| Error::StringTooLong {
        field,
        utf16_units: units,
        limit: usize::from(u8::MAX),
    })
}

/// Converts a length already checked against a two-byte field to those bytes.
pub(crate) fn u16_len(units: usize, field: &'static str) -> Result<u16> {
    u16::try_from(units).map_err(|_error| Error::StringTooLong {
        field,
        utf16_units: units,
        limit: usize::from(u16::MAX),
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn too_long(result: Result<impl std::fmt::Debug>) -> (&'static str, usize, usize) {
        match result {
            Err(Error::StringTooLong {
                field,
                utf16_units,
                limit,
            }) => (field, utf16_units, limit),
            other => panic!("expected StringTooLong, got {other:?}"),
        }
    }

    #[test]
    fn lengths_count_utf16_code_units() {
        assert_eq!(utf16_len(""), 0);
        assert_eq!(utf16_len("ascii"), 5);
        // Two-, three- and four-byte scalars: one, one and two code units.
        assert_eq!(utf16_len("é漢😀"), 4);
    }

    #[test]
    fn a_string_at_the_limit_fits_and_one_unit_more_is_refused() {
        for (value, limit) in [
            ("a".repeat(31), 31),
            ("é".repeat(31), 31),
            (format!("{}😀", "a".repeat(29)), 31),
        ] {
            assert_eq!(checked_utf16_len(&value, limit, "field").unwrap(), limit);
            ensure_utf16_len_within(&value, limit, "field").unwrap();
            assert_eq!(
                too_long(checked_utf16_len(&value, limit - 1, "field")),
                ("field", limit, limit - 1)
            );
            assert_eq!(
                too_long(ensure_utf16_len_within(&value, limit - 1, "field")),
                ("field", limit, limit - 1)
            );
        }
    }

    #[test]
    fn a_long_utf8_string_is_counted_before_it_is_refused() {
        // 96 bytes but 32 code units: the byte length alone cannot refuse it.
        let value = "漢".repeat(32);
        ensure_utf16_len_within(&value, 32, "field").unwrap();
        assert_eq!(
            too_long(ensure_utf16_len_within(&value, 31, "field")),
            ("field", 32, 31)
        );
    }

    #[test]
    fn narrowing_a_checked_length_refuses_what_does_not_fit() {
        assert_eq!(u8_len(255, "byte").unwrap(), 255);
        assert_eq!(too_long(u8_len(256, "byte")), ("byte", 256, 255));
        assert_eq!(u16_len(65_535, "word").unwrap(), 65_535);
        assert_eq!(too_long(u16_len(65_536, "word")), ("word", 65_536, 65_535));
    }
}
