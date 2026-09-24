use super::super::named_range;
use super::super::{DefinedName, DefinedNameRecordOptions, Writer};
use crate::error::{Error, Result};
use crate::writer::string_limits::{
    DEFINED_NAME_UNITS, NAME_COMMENT_UNITS, ensure_utf16_len_within,
};

impl Writer {
    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 UTF-16 code units (`Lbl.cch` is a
    ///   byte counting them), refused with [`Error::StringTooLong`].
    /// - Name MUST NOT contain NUL, which litchi's reader refuses in a
    ///   user-defined `Lbl` name.
    fn validate_defined_name(name: &str) -> Result<()> {
        if name.is_empty() {
            return Err(Error::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        ensure_utf16_len_within(name, DEFINED_NAME_UNITS, "defined name")?;
        if name.contains('\0') {
            return Err(Error::InvalidData(
                "Defined name must not contain a NUL character".to_string(),
            ));
        }
        Ok(())
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers; anything else is refused here, before the
    /// name is stored, so the writer never holds a name it cannot write.
    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] for a name longer than 255 UTF-16 code
    /// units, [`Error::InvalidData`] for an empty name or one containing NUL,
    /// [`Error::TooMany`] past the 65,535 names a workbook holds, the
    /// reference's parse error for an unsupported reference, and an error if
    /// any other validation fails; a refused name leaves the writer
    /// unchanged.
    pub fn define_name(&mut self, name: &str, reference: &str) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name: workbook must have at least one worksheet".to_string(),
            ));
        }

        // For now, workbook-scoped names that refer to cell ranges are
        // anchored to the first worksheet. Users who need explicit
        // sheet scoping can use `define_name_local`.
        let target_sheet = 0u16;

        self.push_defined_name(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        })
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    /// # Errors
    ///
    /// Refuses a name as [`Self::define_name`] does, and an unknown sheet;
    /// a refused name leaves the writer unchanged.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet: usize) -> Result<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;

        let itab = u16::try_from(sheet + 1).map_err(|_error| {
            Error::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.push_defined_name(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(crate::utils::truncate_usize_to_u16(sheet)),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        })
    }

    /// Define a workbook-scoped named range with a user-visible comment.
    /// # Errors
    ///
    /// Refuses a name as [`Self::define_name`] does, and returns
    /// [`Error::StringTooLong`] for a comment longer than the 255 UTF-16 code
    /// units a `NameCmt` record holds; a refused name leaves the writer
    /// unchanged.
    pub fn define_name_with_comment(
        &mut self,
        name: &str,
        reference: &str,
        comment: &str,
    ) -> Result<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(Error::InvalidData(
                "define_name_with_comment: workbook must have at least one worksheet".to_string(),
            ));
        }

        let target_sheet = 0u16;

        self.push_defined_name(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        })
    }

    /// Stores `name` once everything the write would check about it has been
    /// checked: its reference encodes and its comment fits `NameCmt`. A
    /// refused name is not stored, so the writer stays able to write.
    fn push_defined_name(&mut self, name: DefinedName) -> Result<()> {
        // The reader refuses a workbook with more than 65,535 `Lbl` records.
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::TooMany {
                collection: "defined names",
                limit: usize::from(u16::MAX),
            });
        }
        name.to_biff_formula()?;
        if let Some(comment) = &name.comment {
            ensure_utf16_len_within(comment, NAME_COMMENT_UNITS, "defined-name comment")?;
        }
        self.defined_names.push(name);
        Ok(())
    }

    /// Remove all defined names with the given name.
    ///
    /// Returns `true` if at least one name was removed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.defined_names.len();
        self.defined_names.retain(|n| n.name != name);
        self.defined_names.len() < initial_len
    }

    /// Get all defined names in this workbook.
    #[must_use]
    pub fn named_ranges(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Add complete inert BIFF8 defined-name metadata.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_defined_name_record(&mut self, options: DefinedNameRecordOptions) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records
            .push((options, Default::default()));
        Ok(index)
    }

    /// Add complete inert `Lbl` metadata and its ordered BIFF8 future records.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn add_defined_name_record_with_future_records(
        &mut self,
        options: DefinedNameRecordOptions,
        future: crate::DefinedNameFutureRecords,
    ) -> Result<usize> {
        options.validate(self.worksheets.len())?;
        named_range::validate_future_records(&future, options.serialized_name())?;
        if self.defined_names.len() + self.defined_name_records.len() >= usize::from(u16::MAX) {
            return Err(Error::InvalidData(
                "defined name count exceeds BIFF8 bound".to_string(),
            ));
        }
        let index = self.defined_name_records.len();
        self.defined_name_records.push((options, future));
        Ok(index)
    }
}
