use super::super::{FileSharing, SheetProtection, Writer};
use crate::error::{Error, Result};

impl Writer {
    fn hash_password(password: &str) -> u16 {
        let bytes = password.as_bytes();
        if bytes.is_empty() {
            return 0;
        }

        let mut hash: u16 = 0;
        for &b in bytes.iter().rev() {
            let high_bit = (hash >> 14) & 0x0001;
            hash = ((hash << 1) & 0x7FFF) | high_bit;
            hash ^= u16::from(b);
        }

        let high_bit = (hash >> 14) & 0x0001;
        hash = ((hash << 1) & 0x7FFF) | high_bit;
        hash ^= crate::utils::truncate_usize_to_u16(bytes.len());
        hash ^= 0xCE4B;
        hash
    }

    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        protect_structure: bool,
        protect_windows: bool,
    ) {
        if !protect_structure && !protect_windows && password.is_none() {
            self.workbook_protection = None;
            return;
        }

        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_structure = protect_structure;
        protection.protect_windows = protect_windows;
        protection.password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    pub fn unprotect_workbook(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_structure = false;
            protection.protect_windows = false;
            protection.password_hash = None;
            self.workbook_protection = protection.protect_revisions.then_some(protection);
        }
    }

    /// Configure legacy shared-workbook revision protection.
    pub fn protect_revisions(&mut self, password: Option<&str>) {
        let mut protection = self.workbook_protection.unwrap_or_default();
        protection.protect_revisions = true;
        protection.revision_password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(protection);
    }

    /// Remove shared-workbook revision protection.
    pub fn unprotect_revisions(&mut self) {
        if let Some(mut protection) = self.workbook_protection {
            protection.protect_revisions = false;
            protection.revision_password_hash = None;
            self.workbook_protection = (protection.protect_structure
                || protection.protect_windows
                || protection.password_hash.is_some())
            .then_some(protection);
        }
    }

    /// Configure read-only recommendation and an optional write-reservation password.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn set_file_sharing(
        &mut self,
        read_only_recommended: bool,
        password: Option<&str>,
        user_name: &str,
    ) -> Result<()> {
        if user_name.encode_utf16().count() > 54 {
            return Err(Error::InvalidData(
                "FILESHARING username exceeds 54 UTF-16 code units".to_string(),
            ));
        }
        self.file_sharing = Some(FileSharing {
            read_only_recommended,
            password_hash: password.map(Self::hash_password),
            user_name: user_name.to_string(),
        });
        Ok(())
    }

    pub fn clear_file_sharing(&mut self) {
        self.file_sharing = None;
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn protect_sheet(
        &mut self,
        sheet: usize,
        password: Option<&str>,
        protect_objects: bool,
        protect_scenarios: bool,
    ) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;

        let password_hash = password.map(Self::hash_password);
        worksheet.sheet_protection = Some(SheetProtection {
            protect_objects,
            protect_scenarios,
            password_hash,
        });

        Ok(())
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn unprotect_sheet(&mut self, sheet: usize) -> Result<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet.sheet_protection = None;
        Ok(())
    }
}
