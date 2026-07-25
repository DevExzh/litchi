//! Copy-on-write appearance CRUD for Numbers tables.

use super::*;
use crate::table_appearance::{
    TableAppearance, set_table_appearance as set_native_table_appearance,
    table_appearance as read_native_table_appearance,
};

impl NumbersEditor {
    /// Read the effective alternating-row and automatic-sizing settings.
    pub fn table_appearance(&self, table_id: u64) -> Result<TableAppearance> {
        read_native_table_appearance(&self.package, table_id)
    }

    /// Replace appearance settings without mutating styles shared by other tables.
    pub fn set_table_appearance(
        &mut self,
        table_id: u64,
        appearance: TableAppearance,
    ) -> Result<()> {
        if self.table_appearance(table_id)? == appearance {
            return Ok(());
        }
        let mut staged = self.package.clone();
        set_native_table_appearance(&mut staged, table_id, appearance)?;
        let verified = Self::from_package(staged)?;
        if verified.table_appearance(table_id)? != appearance {
            return Err(Error::InvalidFormat(
                "Numbers table appearance failed round-trip validation".to_owned(),
            ));
        }
        self.package = verified.package;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::table_appearance::{TableRowBanding, TableRowSizing};

    #[test]
    fn scratch_table_appearance_is_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_name("Source")
            .table_dimensions(3, 2)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].clone();
        let duplicate = editor.duplicate_table(source.object_id).unwrap();
        let appearance = TableAppearance {
            row_banding: TableRowBanding::Enabled,
            row_sizing: TableRowSizing::FitCellContents,
        };

        editor
            .set_table_appearance(duplicate.object_id, appearance)
            .unwrap();

        assert_eq!(
            editor.table_appearance(source.object_id).unwrap(),
            TableAppearance::default()
        );
        assert_eq!(
            editor.table_appearance(duplicate.object_id).unwrap(),
            appearance
        );
        assert_eq!(
            editor
                .tables()
                .unwrap()
                .into_iter()
                .find(|table| table.object_id == duplicate.object_id)
                .unwrap()
                .appearance,
            appearance
        );
    }
}
