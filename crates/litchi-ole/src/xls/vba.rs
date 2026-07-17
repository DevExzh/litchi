//! Inert BIFF8 VBA project metadata.
//!
//! This module reports markers, object code names, and OLE storage presence.
//! It never opens, decompresses, parses, or executes VBA streams.

use super::{XlsError, XlsResult};

pub(crate) const OB_PROJ_RECORD_TYPE: u16 = 0x00D3;
pub(crate) const CODE_NAME_RECORD_TYPE: u16 = 0x01BA;
pub(crate) const OB_NO_MACROS_RECORD_TYPE: u16 = 0x01BF;
const DIMENSIONS_RECORD_TYPE: u16 = 0x0200;

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct XlsVbaMetadata {
    project_marker: bool,
    no_macros_marker: bool,
    project_storage_present: bool,
    workbook_code_name: Option<String>,
}

impl XlsVbaMetadata {
    pub fn has_project_marker(&self) -> bool {
        self.project_marker
    }
    pub fn has_no_macros_marker(&self) -> bool {
        self.no_macros_marker
    }
    pub fn has_project_storage(&self) -> bool {
        self.project_storage_present
    }
    pub fn workbook_code_name(&self) -> Option<&str> {
        self.workbook_code_name.as_deref()
    }
    pub fn may_contain_executable_code(&self) -> bool {
        self.project_marker && self.project_storage_present && !self.no_macros_marker
    }
    pub fn markers_are_consistent(&self) -> bool {
        !self.no_macros_marker || self.project_marker
    }
    pub(crate) fn set_project_storage_present(&mut self, present: bool) {
        self.project_storage_present = present;
    }
}

pub(crate) struct WorkbookVbaCollector {
    metadata: XlsVbaMetadata,
    last_rank: Option<u8>,
}

impl WorkbookVbaCollector {
    pub(crate) fn new() -> Self {
        Self {
            metadata: XlsVbaMetadata::default(),
            last_rank: None,
        }
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        let rank = match record_type {
            OB_PROJ_RECORD_TYPE => 0,
            OB_NO_MACROS_RECORD_TYPE => 1,
            CODE_NAME_RECORD_TYPE => 2,
            _ => return Ok(()),
        };
        if self.last_rank.is_some_and(|previous| rank < previous) {
            return invalid(
                record_type,
                "VBA metadata record is out of workbook-global order",
            );
        }
        self.last_rank = Some(rank);
        match record_type {
            OB_PROJ_RECORD_TYPE => {
                if self.metadata.project_marker {
                    return invalid(record_type, "duplicate ObProj record");
                }
                require_empty(record_type, data)?;
                self.metadata.project_marker = true;
            },
            OB_NO_MACROS_RECORD_TYPE => {
                if self.metadata.no_macros_marker {
                    return invalid(record_type, "duplicate ObNoMacros record");
                }
                require_empty(record_type, data)?;
                if !self.metadata.project_marker {
                    return invalid(record_type, "ObNoMacros requires a preceding ObProj record");
                }
                self.metadata.no_macros_marker = true;
            },
            CODE_NAME_RECORD_TYPE => {
                if self.metadata.workbook_code_name.is_some() {
                    return invalid(record_type, "duplicate workbook CodeName record");
                }
                self.metadata.workbook_code_name = Some(parse_code_name(data)?);
            },
            _ => unreachable!(),
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> XlsVbaMetadata {
        self.metadata
    }
}

pub(crate) struct WorksheetVbaCollector {
    code_name: Option<String>,
    dimensions_seen: bool,
}

impl WorksheetVbaCollector {
    pub(crate) fn new() -> Self {
        Self {
            code_name: None,
            dimensions_seen: false,
        }
    }
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            DIMENSIONS_RECORD_TYPE => self.dimensions_seen = true,
            OB_PROJ_RECORD_TYPE | OB_NO_MACROS_RECORD_TYPE => {
                return invalid(
                    record_type,
                    "workbook VBA marker appears in worksheet scope",
                );
            },
            CODE_NAME_RECORD_TYPE => {
                if !self.dimensions_seen {
                    return invalid(
                        record_type,
                        "worksheet CodeName appears before worksheet content",
                    );
                }
                if self.code_name.is_some() {
                    return invalid(record_type, "duplicate worksheet CodeName record");
                }
                self.code_name = Some(parse_code_name(data)?);
            },
            _ => {},
        }
        Ok(())
    }
    pub(crate) fn finish(self) -> Option<String> {
        self.code_name
    }
}

pub(crate) fn validate_code_name(value: &str) -> XlsResult<()> {
    if value.encode_utf16().count() > 31 {
        return invalid_data("VBA object code name exceeds 31 UTF-16 code units");
    }
    if value.is_empty() {
        return Ok(());
    }
    let mut characters = value.chars();
    let first = characters.next().unwrap();
    if first.is_ascii() && !first.is_ascii_alphabetic() {
        return invalid_data("VBA object code name must begin with a letter");
    }
    if characters.any(|character| {
        character.is_ascii() && !(character.is_ascii_alphanumeric() || character == '_')
    }) {
        return invalid_data("VBA object code name contains an invalid ASCII character");
    }
    if value.chars().any(|character| character == '\u{FFE3}') {
        return invalid_data("VBA object code name contains forbidden U+FFE3");
    }
    Ok(())
}

pub(crate) fn parse_code_name(data: &[u8]) -> XlsResult<String> {
    if data.len() < 3 {
        return invalid(CODE_NAME_RECORD_TYPE, "truncated CodeName record");
    }
    let count = u16::from_le_bytes([data[0], data[1]]) as usize;
    if count > 31 {
        return invalid(CODE_NAME_RECORD_TYPE, "CodeName exceeds 31 characters");
    }
    let options = data[2];
    if options & 0xFE != 0 {
        return invalid(
            CODE_NAME_RECORD_TYPE,
            "CodeName contains reserved string option bits",
        );
    }
    let width = if options & 1 == 0 { 1 } else { 2 };
    let expected = 3usize
        .checked_add(
            count
                .checked_mul(width)
                .ok_or_else(|| XlsError::InvalidData("CodeName size overflow".to_string()))?,
        )
        .ok_or_else(|| XlsError::InvalidData("CodeName size overflow".to_string()))?;
    if data.len() != expected {
        return invalid(
            CODE_NAME_RECORD_TYPE,
            "CodeName character count does not match payload length",
        );
    }
    let value = if width == 1 {
        data[3..].iter().map(|byte| char::from(*byte)).collect()
    } else {
        let units: Vec<u16> = data[3..]
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect();
        String::from_utf16(&units).map_err(|_| XlsError::InvalidRecord {
            record_type: CODE_NAME_RECORD_TYPE,
            message: "CodeName contains invalid UTF-16".to_string(),
        })?
    };
    validate_code_name(&value).map_err(|error| match error {
        XlsError::InvalidData(message) => XlsError::InvalidRecord {
            record_type: CODE_NAME_RECORD_TYPE,
            message,
        },
        other => other,
    })?;
    Ok(value)
}

fn require_empty(record_type: u16, data: &[u8]) -> XlsResult<()> {
    if !data.is_empty() {
        return invalid(record_type, "marker record payload must be empty");
    }
    Ok(())
}
fn invalid<T>(record_type: u16, message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidRecord {
        record_type,
        message: message.into(),
    })
}
fn invalid_data<T>(message: impl Into<String>) -> XlsResult<T> {
    Err(XlsError::InvalidData(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn parses_strict_code_names() {
        assert_eq!(
            parse_code_name(&[6, 0, 0, b'S', b'h', b'e', b'e', b't', b'1']).unwrap(),
            "Sheet1"
        );
        assert!(parse_code_name(&[1, 0, 2, b'A']).is_err());
        assert!(parse_code_name(&[2, 0, 0, b'A']).is_err());
        assert!(parse_code_name(&[1, 0, 0, b'1']).is_err());
    }
    #[test]
    fn rejects_bad_marker_lengths_order_and_scope() {
        let mut globals = WorkbookVbaCollector::new();
        assert!(globals.feed_record(OB_PROJ_RECORD_TYPE, &[0]).is_err());
        let mut globals = WorkbookVbaCollector::new();
        assert!(globals.feed_record(OB_NO_MACROS_RECORD_TYPE, &[]).is_err());
        let mut sheet = WorksheetVbaCollector::new();
        assert!(sheet.feed_record(OB_PROJ_RECORD_TYPE, &[]).is_err());
        assert!(
            sheet
                .feed_record(CODE_NAME_RECORD_TYPE, &[1, 0, 0, b'A'])
                .is_err()
        );
    }
    #[test]
    fn no_macros_marker_is_inert() {
        let mut globals = WorkbookVbaCollector::new();
        globals.feed_record(OB_PROJ_RECORD_TYPE, &[]).unwrap();
        globals.feed_record(OB_NO_MACROS_RECORD_TYPE, &[]).unwrap();
        let metadata = globals.finish();
        assert!(metadata.has_project_marker());
        assert!(metadata.has_no_macros_marker());
        assert!(!metadata.may_contain_executable_code());
    }
}
