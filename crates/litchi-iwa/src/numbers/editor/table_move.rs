//! Compatibility bridge for transferring an existing table between workbook sheets.

use super::*;
use litchi_numbers::{Package as FocusedNumbersPackage, SheetSelector, TableSelector};

impl NumbersEditor {
    /// Move an existing table to another sheet, preserving its object identity and contents.
    ///
    /// This compatibility method resolves the historical workbook-wide table selector to the
    /// focused sheet/table positions, delegates the package mutation to
    /// [`litchi_numbers::Package::move_table`], and validates the legacy readback before
    /// publishing it. The table is appended to the destination sheet's drawable order, matching
    /// the former native implementation.
    #[deprecated(
        since = "0.0.1",
        note = "legacy workbook-wide Numbers table move; use litchi_numbers::Package::move_table with source and destination SheetSelector values and a sheet-scoped TableSelector"
    )]
    pub fn move_table(
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

        let (source_sheet, focused_table) =
            super::selectors::focused_table_location(self, table_id)?;
        let source_bytes = self.to_bytes()?;
        let bytes = match FocusedNumbersPackage::from_bytes(&source_bytes) {
            Ok(source) => {
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
            },
            Err(_projection_error) => {
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
            },
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    #[test]
    fn moves_table_with_name_and_sheet_selectors() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let target = editor.add_empty_sheet("Archive").unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;

        let moved = editor
            .move_table(
                TableSelector::name("Revenue"),
                SheetSelector::name("Archive"),
            )
            .unwrap();
        assert_eq!(moved.object_id, table_id);
        assert_eq!(
            find_table_owner(editor.package(), table_id)
                .unwrap()
                .sheet_id,
            target.object_id
        );
        assert!(
            editor
                .move_table(TableSelector::index(1), SheetSelector::index(0))
                .is_err()
        );
    }
}
