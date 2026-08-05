use super::super::named_range;
use super::super::*;
use crate::error::{Error, Result};

impl Writer {
    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 characters (Lbl.cch is a byte).
    fn validate_defined_name(name: &str) -> Result<()> {
        if name.is_empty() {
            return Err(Error::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        let char_count = name.chars().count();
        if char_count > u8::MAX as usize {
            return Err(Error::InvalidData(
                "Defined name must be at most 255 characters".to_string(),
            ));
        }

        Ok(())
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers. More complex formulas will be rejected
    /// at serialization time to avoid emitting invalid BIFF payloads.
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

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet: usize) -> Result<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            Error::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a workbook-scoped named range with a user-visible comment.
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

        self.defined_names.push(DefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

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
    pub fn named_ranges(&self) -> &[DefinedName] {
        &self.defined_names
    }

    /// Add complete inert BIFF8 defined-name metadata.
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
