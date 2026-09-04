//! Private compatibility bridge for transferring an existing table between workbook sheets.
//!
//! The focused [`litchi_numbers::Package`] transaction owns the physical graph
//! rewrite.  This adapter remains crate-private solely because populated-sheet
//! duplication still builds its new sheet through the legacy editor before it
//! can hand the cloned table to the focused owner.  No migration-host move API
//! is exported from `litchi-iwa`.

use super::*;
use litchi_numbers::{Package as FocusedNumbersPackage, SheetSelector, TableSelector};

impl NumbersEditor {
    /// Move an existing table for the internal populated-sheet duplication path.
    ///
    /// The historical workbook-wide public route was retired; this narrow
    /// crate-private method keeps only the duplication implementation's
    /// selector adaptation and legacy readback validation.
    pub(crate) fn move_table(
        &mut self,
        selector: TableSelector,
        target: SheetSelector,
    ) -> Result<NumbersTableInfo> {
        let table_id = super::selectors::table_id(self, selector)?;
        let table = self
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers table {table_id} not found")))?;
        let target_sheet_id = super::selectors::sheet_id(self, target)?;
        let owner = find_table_owner(self.package(), table_id)?;
        if owner.sheet_id == target_sheet_id {
            return Ok(table);
        }

        let source_built = !self.package.source_is_exact();
        let (source_sheet, focused_table) =
            super::selectors::focused_table_location(self, table_id)?;
        let source_bytes = self.to_bytes()?;
        let bytes = if source_built {
            // A legacy builder snapshot is normalized to an exact byte owner
            // before the focused crate's private compatibility admission runs.
            // The migration host still performs no archive or wire mutation.
            FocusedNumbersPackage::__move_table_from_bytes_for_compatibility(
                &source_bytes,
                source_sheet,
                focused_table,
                target,
            )
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "focused Numbers table-move compatibility admission failed: {error}"
                ))
            })?
        } else {
            // Exact snapshots stay on the focused owner once ingress is
            // attempted.  A focused ingress/read/edit/commit failure is
            // terminal; the physical bridge is reserved for source-built
            // snapshots above.
            let source = FocusedNumbersPackage::from_bytes(&source_bytes).map_err(|error| {
                Error::InvalidFormat(format!(
                    "focused Numbers table-move source validation failed: {error}"
                ))
            })?;
            let commit = source
                .move_table(source_sheet, focused_table, target)
                .map_err(|error| {
                    Error::InvalidFormat(format!("focused Numbers table move failed: {error}"))
                })?;
            let mut bytes = Vec::new();
            commit
                .package()
                .write_to(&mut bytes)
                .map_err(|error| Error::Io(error.into_io_error()))?;
            bytes
        };
        let verified = Self::from_bytes(&bytes)?;
        let verified_owner = find_table_owner(verified.package(), table_id)?;
        let verified_table = verified
            .tables()?
            .into_iter()
            .find(|candidate| candidate.object_id == table_id)
            .ok_or_else(|| Error::InvalidFormat("Moved Numbers table disappeared".to_owned()))?;
        if verified_owner.sheet_id != target_sheet_id || verified_table != table {
            return Err(Error::InvalidFormat(
                "Numbers table move failed validation".to_owned(),
            ));
        }

        self.package = verified.package;
        Ok(verified_table)
    }
}
