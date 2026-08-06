//! Bounds and ownership validation for one inert XLSB external-link part.

use super::codec::is_modeled_record;
use super::{
    Error, Kind, Link, MAX_LINK_PART_BYTES, MAX_UNKNOWN_BYTES, MAX_UNKNOWN_RECORDS,
    MAX_WIDE_STRING_UNITS, Result, UnknownRecord,
};

/// Validate the semantic external-link graph without resolving its source.
pub fn validate_link(link: &Link) -> Result<()> {
    link.validate()?;
    if link.kind() == Kind::Workbook {
        for name in link.sheet_names() {
            let units = name.encode_utf16().count();
            if units == 0
                || units > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(Error::InvalidFormula(format!(
                    "external sheet name {name:?} does not follow sheet-name grammar"
                )));
            }
        }
    }
    Ok(())
}

/// Validate the relationship cardinality required by `[MS-XLSB]`.
pub(crate) fn validate_relationship(link: &Link, relationship_id: Option<&str>) -> Result<()> {
    match (link.kind(), relationship_id) {
        (Kind::Dde, None) => Ok(()),
        (Kind::Dde, Some(_)) => Err(Error::InvalidFormula(
            "DDE external link cannot own a package relationship".to_string(),
        )),
        (Kind::Workbook | Kind::Ole, Some(value)) if !value.is_empty() => {
            validate_relationship_id(value)
        },
        (Kind::Workbook, None) => Err(Error::InvalidFormula(
            "external workbook link must own an external-link relationship".to_string(),
        )),
        (Kind::Ole, None) => Err(Error::InvalidFormula(
            "OLE external link must own an OLE relationship".to_string(),
        )),
        (Kind::Workbook | Kind::Ole, Some(_)) => Err(Error::InvalidFormula(
            "external-link relationship identifier must not be empty".to_string(),
        )),
    }
}

/// Validate one relationship identifier without following it.
pub(crate) fn validate_relationship_id(value: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if units == 0 || units > MAX_WIDE_STRING_UNITS || value.contains('\0') {
        return Err(Error::InvalidFormula(
            "external-link relationship identifier is empty, too long, or contains NUL".to_string(),
        ));
    }
    Ok(())
}

/// Validate the source-owned opaque record set before it is merged into a
/// newly authored stream.
pub(crate) fn validate_unknown_records(records: &[UnknownRecord]) -> Result<()> {
    if records.len() > MAX_UNKNOWN_RECORDS {
        return Err(Error::InvalidLength {
            expected: MAX_UNKNOWN_RECORDS,
            found: records.len(),
        });
    }
    let mut total = 0usize;
    let limits = crate::raw::Limits::new(MAX_LINK_PART_BYTES, MAX_WIDE_STRING_UNITS);
    let mut previous_anchor = 0usize;
    for record in records {
        if record.bytes().is_empty() {
            return Err(Error::InvalidLength {
                expected: 1,
                found: 0,
            });
        }
        if record.after_known() < previous_anchor {
            return Err(Error::InvalidFormula(
                "opaque external-link records are not source ordered".to_string(),
            ));
        }
        previous_anchor = record.after_known();
        total = total.checked_add(record.bytes().len()).ok_or_else(|| {
            Error::InvalidFormula("opaque external-link byte count overflow".to_string())
        })?;
        if total > MAX_UNKNOWN_BYTES {
            return Err(Error::InvalidLength {
                expected: MAX_UNKNOWN_BYTES,
                found: total,
            });
        }
        let (header, header_len) = crate::raw::Header::parse(record.bytes(), limits)?;
        let expected = header_len
            .checked_add(header.len())
            .ok_or_else(|| Error::InvalidFormula("opaque record length overflow".to_string()))?;
        if expected != record.bytes().len()
            || u16::from(header.kind()) != record.kind()
            || is_modeled_record(header.kind())
        {
            return Err(Error::InvalidFormula(
                "opaque record does not retain a complete unknown BIFF12 record".to_string(),
            ));
        }
    }
    Ok(())
}
