//! BIFF8 workbook and worksheet protection records.
//!
//! Verifiers exposed here are inert legacy metadata, not encryption keys.

use crate::xls::{XlsError, XlsResult};

pub const PROTECT_TYPE: u16 = 0x0012;
pub const PASSWORD_TYPE: u16 = 0x0013;
pub const WINPROTECT_TYPE: u16 = 0x0019;
pub const FILESHARING_TYPE: u16 = 0x005B;
pub const OBJECTPROTECT_TYPE: u16 = 0x0063;
pub const WRITEPROTECT_TYPE: u16 = 0x0086;
pub const SCENPROTECT_TYPE: u16 = 0x00DD;
pub const PROT4REV_TYPE: u16 = 0x01AF;
pub const PROT4REVPASS_TYPE: u16 = 0x01BC;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct PasswordVerifier(u16);

impl PasswordVerifier {
    pub const fn from_raw(value: u16) -> Self { Self(value) }
    pub const fn raw(self) -> u16 { self.0 }
    pub const fn is_set(self) -> bool { self.0 != 0 }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSharing {
    read_only_recommended: bool,
    write_password: PasswordVerifier,
    user_name: String,
}

impl FileSharing {
    pub const fn read_only_recommended(&self) -> bool { self.read_only_recommended }
    pub const fn write_password(&self) -> PasswordVerifier { self.write_password }
    pub fn user_name(&self) -> &str { &self.user_name }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct WorkbookProtection {
    structure_protected: bool,
    windows_protected: bool,
    password: PasswordVerifier,
    revisions_protected: bool,
    revision_password: PasswordVerifier,
    write_protected: bool,
    file_sharing: Option<FileSharing>,
}

impl WorkbookProtection {
    pub const fn structure_protected(&self) -> bool { self.structure_protected }
    pub const fn windows_protected(&self) -> bool { self.windows_protected }
    pub const fn password(&self) -> PasswordVerifier { self.password }
    pub const fn revisions_protected(&self) -> bool { self.revisions_protected }
    pub const fn revision_password(&self) -> PasswordVerifier { self.revision_password }
    pub const fn write_protected(&self) -> bool { self.write_protected }
    pub fn file_sharing(&self) -> Option<&FileSharing> { self.file_sharing.as_ref() }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct SheetProtection {
    sheet_protected: bool,
    objects_protected: bool,
    scenarios_protected: bool,
    password: Option<PasswordVerifier>,
}

impl SheetProtection {
    pub const fn is_protected(&self) -> bool { self.sheet_protected }
    pub const fn objects_protected(&self) -> bool { self.objects_protected }
    pub const fn scenarios_protected(&self) -> bool { self.scenarios_protected }
    pub const fn password(&self) -> Option<PasswordVerifier> { self.password }
    pub const fn has_password(&self) -> bool { self.password.is_some() }
}

fn parse_bool(record_type: u16, data: &[u8]) -> XlsResult<bool> {
    if data.len() != 2 {
        return Err(XlsError::InvalidLength { expected: 2, found: data.len() });
    }
    match u16::from_le_bytes([data[0], data[1]]) {
        0 => Ok(false),
        1 => Ok(true),
        value => Err(XlsError::InvalidRecord {
            record_type,
            message: format!("Boolean must be 0 or 1, found 0x{value:04X}"),
        }),
    }
}

fn parse_verifier(data: &[u8]) -> XlsResult<PasswordVerifier> {
    if data.len() != 2 {
        return Err(XlsError::InvalidLength { expected: 2, found: data.len() });
    }
    Ok(PasswordVerifier::from_raw(u16::from_le_bytes([data[0], data[1]])))
}

fn duplicate(record_type: u16) -> XlsError {
    XlsError::InvalidRecord { record_type, message: "duplicate protection record".into() }
}

fn parse_file_sharing(data: &[u8]) -> XlsResult<FileSharing> {
    if data.len() < 6 {
        return Err(XlsError::InvalidLength { expected: 6, found: data.len() });
    }
    let read_only_recommended = parse_bool(FILESHARING_TYPE, &data[..2])?;
    let write_password = PasswordVerifier::from_raw(u16::from_le_bytes([data[2], data[3]]));
    let cch_or_marker = u16::from_le_bytes([data[4], data[5]]);
    if !write_password.is_set() {
        if cch_or_marker != 0 {
            return Err(XlsError::InvalidRecord {
                record_type: FILESHARING_TYPE,
                message: "iNoResPass must be zero".into(),
            });
        }
        if data.len() != 6 {
            return Err(XlsError::InvalidLength { expected: 6, found: data.len() });
        }
        return Ok(FileSharing { read_only_recommended, write_password, user_name: String::new() });
    }
    let cch = usize::from(cch_or_marker);
    if cch > 54 {
        return Err(XlsError::InvalidRecord {
            record_type: FILESHARING_TYPE,
            message: "username exceeds 54 characters".into(),
        });
    }
    if data.len() < 7 {
        return Err(XlsError::InvalidLength { expected: 7, found: data.len() });
    }
    let flags = data[6];
    if flags & !1 != 0 {
        return Err(XlsError::InvalidRecord {
            record_type: FILESHARING_TYPE,
            message: format!("reserved string flags set: 0x{flags:02X}"),
        });
    }
    let wide = flags == 1;
    let expected = 7 + cch * if wide { 2 } else { 1 };
    if data.len() != expected {
        return Err(XlsError::InvalidLength { expected, found: data.len() });
    }
    let user_name = if wide {
        let units = data[7..].chunks_exact(2)
            .map(|v| u16::from_le_bytes([v[0], v[1]])).collect::<Vec<_>>();
        String::from_utf16(&units).map_err(|error| XlsError::InvalidRecord {
            record_type: FILESHARING_TYPE,
            message: format!("invalid UTF-16 username: {error}"),
        })?
    } else {
        data[7..].iter().map(|byte| char::from(*byte)).collect()
    };
    Ok(FileSharing { read_only_recommended, write_password, user_name })
}

#[derive(Debug, Default)]
pub(crate) struct WorkbookProtectionCollector {
    value: WorkbookProtection,
    protect: bool,
    password: bool,
    window: bool,
    rev: bool,
    rev_pass: bool,
    write: bool,
    sharing: bool,
    previous: Option<u16>,
}

impl WorkbookProtectionCollector {
    pub(crate) fn new() -> Self { Self::default() }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            PROTECT_TYPE => {
                if self.protect { return Err(duplicate(record_type)); }
                self.protect = true;
                self.value.structure_protected = parse_bool(record_type, data)?;
            },
            PASSWORD_TYPE => {
                if self.password { return Err(duplicate(record_type)); }
                self.password = true;
                self.value.password = parse_verifier(data)?;
            },
            WINPROTECT_TYPE => {
                if self.window { return Err(duplicate(record_type)); }
                self.window = true;
                self.value.windows_protected = parse_bool(record_type, data)?;
            },
            PROT4REV_TYPE => {
                if self.rev { return Err(duplicate(record_type)); }
                self.rev = true;
                self.value.revisions_protected = parse_bool(record_type, data)?;
            },
            PROT4REVPASS_TYPE => {
                if self.rev_pass { return Err(duplicate(record_type)); }
                if self.previous != Some(PROT4REV_TYPE) {
                    return Err(XlsError::InvalidRecord {
                        record_type,
                        message: "PROT4REVPASS must immediately follow PROT4REV".into(),
                    });
                }
                self.rev_pass = true;
                self.value.revision_password = parse_verifier(data)?;
                if !self.value.revisions_protected && self.value.revision_password.is_set() {
                    return Err(XlsError::InvalidRecord {
                        record_type,
                        message: "revision verifier must be zero when protection is disabled".into(),
                    });
                }
            },
            WRITEPROTECT_TYPE => {
                if self.write { return Err(duplicate(record_type)); }
                if !data.is_empty() {
                    return Err(XlsError::InvalidLength { expected: 0, found: data.len() });
                }
                self.write = true;
                self.value.write_protected = true;
            },
            FILESHARING_TYPE => {
                if self.sharing { return Err(duplicate(record_type)); }
                self.sharing = true;
                self.value.file_sharing = Some(parse_file_sharing(data)?);
            },
            _ => {},
        }
        self.previous = Some(record_type);
        Ok(())
    }

    pub(crate) fn finish(&self) -> XlsResult<WorkbookProtection> {
        if self.rev != self.rev_pass {
            return Err(XlsError::InvalidRecord {
                record_type: PROT4REVPASS_TYPE,
                message: "PROT4REV and PROT4REVPASS must occur as a pair".into(),
            });
        }
        Ok(self.value.clone())
    }
}

#[derive(Debug, Default)]
pub(crate) struct SheetProtectionCollector {
    value: SheetProtection,
    protect: bool,
    object: bool,
    scenario: bool,
    password: bool,
}

impl SheetProtectionCollector {
    pub(crate) fn new() -> Self { Self::default() }

    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> XlsResult<()> {
        match record_type {
            PROTECT_TYPE => {
                if self.protect { return Err(duplicate(record_type)); }
                self.protect = true;
                if !parse_bool(record_type, data)? {
                    return Err(XlsError::InvalidRecord { record_type, message: "sheet PROTECT must be 1".into() });
                }
                self.value.sheet_protected = true;
            },
            OBJECTPROTECT_TYPE => {
                if self.object { return Err(duplicate(record_type)); }
                self.object = true;
                if !parse_bool(record_type, data)? {
                    return Err(XlsError::InvalidRecord { record_type, message: "OBJPROTECT must be 1".into() });
                }
                self.value.objects_protected = true;
            },
            SCENPROTECT_TYPE => {
                if self.scenario { return Err(duplicate(record_type)); }
                self.scenario = true;
                self.value.scenarios_protected = parse_bool(record_type, data)?;
            },
            PASSWORD_TYPE => {
                if self.password { return Err(duplicate(record_type)); }
                self.password = true;
                let verifier = parse_verifier(data)?;
                if !verifier.is_set() {
                    return Err(XlsError::InvalidRecord { record_type, message: "sheet verifier must not be zero".into() });
                }
                self.value.password = Some(verifier);
            },
            _ => {},
        }
        Ok(())
    }

    pub(crate) fn finish(&self) -> XlsResult<SheetProtection> {
        if !self.protect && (self.object || self.scenario || self.password) {
            return Err(XlsError::InvalidRecord {
                record_type: PROTECT_TYPE,
                message: "sheet protection metadata exists without PROTECT".into(),
            });
        }
        Ok(self.value.clone())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_boolean_and_lengths() {
        assert!(parse_bool(PROTECT_TYPE, &[1, 0]).unwrap());
        assert!(parse_bool(PROTECT_TYPE, &[2, 0]).is_err());
        assert!(parse_bool(PROTECT_TYPE, &[1, 0, 0]).is_err());
    }

    #[test]
    fn file_sharing_forms_are_strict() {
        assert!(parse_file_sharing(&[1, 0, 0, 0, 0, 0]).unwrap().read_only_recommended());
        let wide = parse_file_sharing(&[0, 0, 1, 0, 2, 0, 1, b'A', 0, 0x2D, 0x4E]).unwrap();
        assert_eq!(wide.user_name(), "A中");
        assert!(parse_file_sharing(&[0, 0, 0, 0, 1, 0]).is_err());
        assert!(parse_file_sharing(&[0, 0, 1, 0, 0, 0, 0x80]).is_err());
    }

    #[test]
    fn revision_pair_must_be_immediate() {
        let mut value = WorkbookProtectionCollector::new();
        value.feed_record(PROTECT_TYPE, &[0, 0]).unwrap();
        value.feed_record(PASSWORD_TYPE, &[0, 0]).unwrap();
        value.feed_record(PROT4REV_TYPE, &[1, 0]).unwrap();
        value.feed_record(0x1234, &[]).unwrap();
        assert!(value.feed_record(PROT4REVPASS_TYPE, &[1, 0]).is_err());
    }

    #[test]
    fn absent_workbook_records_default_to_unprotected() {
        let value = WorkbookProtectionCollector::new().finish().unwrap();
        assert!(!value.structure_protected());
        assert!(!value.windows_protected());
        assert!(!value.password().is_set());
        assert!(!value.revisions_protected());
        assert!(!value.write_protected());
        assert!(value.file_sharing().is_none());
    }

    #[test]
    fn sheet_password_is_nonzero_and_scoped() {
        let mut value = SheetProtectionCollector::new();
        value.feed_record(PASSWORD_TYPE, &[1, 0]).unwrap();
        assert!(value.finish().is_err());
        let mut value = SheetProtectionCollector::new();
        value.feed_record(PROTECT_TYPE, &[1, 0]).unwrap();
        assert!(value.feed_record(PASSWORD_TYPE, &[0, 0]).is_err());
    }
}
