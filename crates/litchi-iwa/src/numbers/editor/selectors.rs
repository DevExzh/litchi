//! Semantic Numbers selector resolution at the native archive boundary.
//!
//! The public editor accepts only archive-free selectors from
//! `litchi_numbers`. Native object identifiers are resolved once here and are
//! kept below the semantic API.

use super::{NumbersEditor, NumbersSheetInfo};
use crate::{Error, Result};
use litchi_numbers::{SheetSelector, TableSelector};

/// Resolve a semantic sheet selector to its native object identifier.
pub(super) fn sheet_id(editor: &NumbersEditor, selector: SheetSelector<'_>) -> Result<u64> {
    let sheets = editor.sheets()?;
    match selector {
        SheetSelector::Name(name) => unique_named_sheet(&sheets, name),
        SheetSelector::Index(index) => sheets
            .get(index)
            .map(NumbersSheetInfo::native_id)
            .ok_or_else(|| {
                Error::ParseError(format!(
                    "Numbers sheet catalog index {index} is out of bounds"
                ))
            }),
    }
}

/// Resolve a semantic table selector to its native model object identifier.
pub(super) fn table_id(editor: &NumbersEditor, selector: TableSelector<'_>) -> Result<u64> {
    let tables = super::table_models(&editor.package)?;
    match selector {
        TableSelector::Name(name) => unique_named_table(&tables, name),
        TableSelector::Index(index) => {
            tables
                .get(index)
                .map(|table| table.object_id)
                .ok_or_else(|| {
                    Error::ParseError(format!(
                        "Numbers table catalog index {index} is out of bounds"
                    ))
                })
        },
    }
}

/// Return the semantic catalog position of a native table identifier for
/// adapter-internal follow-up operations.
pub(super) fn table_index(editor: &NumbersEditor, native_id: u64) -> Result<usize> {
    super::table_models(&editor.package)?
        .iter()
        .position(|table| table.object_id == native_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table {native_id} not found")))
}

fn unique_named_sheet(sheets: &[NumbersSheetInfo], name: &str) -> Result<u64> {
    let mut matches = sheets.iter().filter(|sheet| sheet.name == name);
    let Some(sheet) = matches.next() else {
        return Err(Error::ParseError(format!(
            "Numbers sheet named {name:?} not found"
        )));
    };
    if matches.next().is_some() {
        return Err(Error::ParseError(format!(
            "Numbers sheet name {name:?} is ambiguous"
        )));
    }
    Ok(sheet.native_id())
}

fn unique_named_table(tables: &[super::model::TableDescriptor], name: &str) -> Result<u64> {
    let mut matches = tables.iter().filter(|table| table.model.table_name == name);
    let Some(table) = matches.next() else {
        return Err(Error::ParseError(format!(
            "Numbers table named {name:?} not found"
        )));
    };
    if matches.next().is_some() {
        return Err(Error::ParseError(format!(
            "Numbers table name {name:?} is ambiguous"
        )));
    }
    Ok(table.object_id)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::numbers::NumbersDocumentBuilder;

    #[test]
    fn selectors_share_the_editor_catalog_and_reject_invalid_entries() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let first_sheet = editor.sheets().unwrap().remove(0);
        let first_table = editor.tables().unwrap().remove(0);
        let duplicate = editor.duplicate_table(TableSelector::index(0)).unwrap();

        assert_eq!(
            sheet_id(&editor, SheetSelector::name("Summary")).unwrap(),
            first_sheet.native_id()
        );
        assert_eq!(
            sheet_id(&editor, SheetSelector::index(0)).unwrap(),
            first_sheet.native_id()
        );
        assert_eq!(
            table_id(&editor, TableSelector::name("Revenue")).unwrap(),
            first_table.native_id()
        );
        assert_eq!(
            table_id(&editor, TableSelector::index(0)).unwrap(),
            first_table.native_id()
        );
        assert_eq!(table_index(&editor, duplicate.native_id()).unwrap(), 1);

        assert!(sheet_id(&editor, SheetSelector::name("Missing")).is_err());
        assert!(sheet_id(&editor, SheetSelector::index(1)).is_err());
        assert!(table_id(&editor, TableSelector::name("Missing")).is_err());
        assert!(table_id(&editor, TableSelector::index(2)).is_err());
    }

    #[test]
    fn table_name_resolution_reports_cross_sheet_ambiguity() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        editor.add_empty_sheet("Archive").unwrap();
        editor
            .add_empty_table(SheetSelector::name("Archive"), "Revenue", 2, 2)
            .unwrap();

        assert!(table_id(&editor, TableSelector::name("Revenue")).is_err());
    }
}
