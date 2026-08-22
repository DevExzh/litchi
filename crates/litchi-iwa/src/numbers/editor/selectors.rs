//! Semantic Numbers selector resolution at the native archive boundary.
//!
//! The public editor accepts only archive-free selectors from
//! `litchi_numbers`. Native object identifiers are resolved once here and are
//! kept below the semantic API.

use std::collections::HashSet;

use super::NumbersEditor;
use crate::{Error, Result};
use litchi_numbers::{Dimensions, Document, Sheet, SheetSelector, Table, TableSelector};

/// A private semantic sheet catalog used only at the legacy archive boundary.
///
/// The focused Numbers selector is intentionally archive-free.  The host still
/// has to hand a selected native object to its existing writer, so this adapter
/// keeps the native IDs in a parallel private vector and delegates the actual
/// name/index match to `litchi_numbers::Document::sheet`.
struct SheetSelectorAdapter {
    semantic: Document,
    source_names: Vec<String>,
    native_ids: Vec<u64>,
}

impl SheetSelectorAdapter {
    fn from_editor(editor: &NumbersEditor) -> Result<Self> {
        let source = editor.sheets()?;
        let source_names = source
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let semantic_sheets = selector_safe_names(&source_names, "sheet")
            .into_iter()
            .enumerate()
            .map(|(index, name)| Sheet::new(name, index))
            .collect::<Vec<_>>();
        let semantic = Document::from_sheets(semantic_sheets).map_err(|error| {
            Error::InvalidFormat(format!(
                "Numbers sheet selector catalog is invalid: {error}"
            ))
        })?;
        let native_ids = source
            .into_iter()
            .map(|sheet| sheet.native_id())
            .collect::<Vec<_>>();
        Ok(Self {
            semantic,
            source_names,
            native_ids,
        })
    }

    fn native_id(&self, selector: SheetSelector<'_>) -> Result<u64> {
        ensure_unique_sheet_name(&self.source_names, selector)?;
        let index = self
            .semantic
            .sheet(selector)
            .map_err(|error| {
                Error::InvalidFormat(format!("Numbers sheet selector failed: {error}"))
            })?
            .map(Sheet::index)
            .ok_or_else(|| sheet_selector_error(selector))?;
        self.native_ids.get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers sheet selector catalog lost native entry at index {index}"
            ))
        })
    }
}

/// A private semantic table catalog used only at the legacy archive boundary.
///
/// `TableSelector` is scoped to one focused semantic sheet.  The historical
/// editor API predates that scope and exposes one workbook-wide table catalog,
/// so the adapter presents that catalog as one synthetic semantic sheet.  Its
/// source-name preflight retains the host's cross-sheet ambiguity rule.
struct TableSelectorAdapter {
    semantic: Sheet,
    source_names: Vec<String>,
    native_ids: Vec<u64>,
}

impl TableSelectorAdapter {
    fn from_editor(editor: &NumbersEditor) -> Result<Self> {
        let descriptors = super::table_models(&editor.package)?;
        let source_names = descriptors
            .iter()
            .map(|table| table.model.table_name.clone())
            .collect::<Vec<_>>();
        let semantic_tables = selector_safe_names(&source_names, "table")
            .into_iter()
            .zip(&descriptors)
            .map(|(name, descriptor)| {
                Table::new(
                    name,
                    Dimensions::new(
                        descriptor.model.number_of_rows,
                        descriptor.model.number_of_columns,
                    ),
                )
            })
            .collect::<Vec<_>>();
        let semantic = Sheet::try_from_tables("Numbers table selector catalog", 0, semantic_tables)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "Numbers table selector catalog is invalid: {error}"
                ))
            })?;
        let native_ids = descriptors
            .into_iter()
            .map(|table| table.object_id)
            .collect::<Vec<_>>();
        Ok(Self {
            semantic,
            source_names,
            native_ids,
        })
    }

    fn native_id(&self, selector: TableSelector<'_>) -> Result<u64> {
        ensure_unique_table_name(&self.source_names, selector)?;
        let selected = self
            .semantic
            .select(selector)
            .map_err(|error| {
                Error::InvalidFormat(format!("Numbers table selector failed: {error}"))
            })?
            .ok_or_else(|| table_selector_error(selector))?;
        let index = self
            .semantic
            .tables()
            .position(|table| std::ptr::eq(table, selected))
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers table selector catalog lost semantic entry".to_owned(),
                )
            })?;
        self.native_ids.get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table selector catalog lost native entry at index {index}"
            ))
        })
    }
}

/// Make a source catalog safe to pass through focused immutable constructors.
///
/// Native malformed packages can repeat a visible name.  The requested name
/// is checked against the original names before this list is used, while
/// duplicate entries that are irrelevant to an index lookup receive private
/// labels so they cannot make the focused catalog reject an otherwise valid
/// positional selection.  Matching remains exact and case-sensitive.
fn selector_safe_names(source_names: &[String], kind: &str) -> Vec<String> {
    let mut used = source_names.iter().cloned().collect::<HashSet<_>>();
    let mut seen = HashSet::new();
    source_names
        .iter()
        .enumerate()
        .map(|(index, name)| {
            if seen.insert(name.as_str()) {
                return name.clone();
            }
            let mut replacement = format!("\0litchi-selector-{kind}-{index}");
            while used.contains(&replacement) {
                replacement.push('_');
            }
            used.insert(replacement.clone());
            replacement
        })
        .collect()
}

fn ensure_unique_sheet_name(source_names: &[String], selector: SheetSelector<'_>) -> Result<()> {
    let SheetSelector::Name(name) = selector else {
        return Ok(());
    };
    let matches = source_names
        .iter()
        .filter(|candidate| candidate.as_str() == name)
        .count();
    match matches {
        0 => Err(sheet_selector_error(selector)),
        1 => Ok(()),
        _ => Err(Error::ParseError(format!(
            "Numbers sheet name {name:?} is ambiguous"
        ))),
    }
}

fn ensure_unique_table_name(source_names: &[String], selector: TableSelector<'_>) -> Result<()> {
    let TableSelector::Name(name) = selector else {
        return Ok(());
    };
    let matches = source_names
        .iter()
        .filter(|candidate| candidate.as_str() == name)
        .count();
    match matches {
        0 => Err(table_selector_error(selector)),
        1 => Ok(()),
        _ => Err(Error::ParseError(format!(
            "Numbers table name {name:?} is ambiguous"
        ))),
    }
}

fn sheet_selector_error(selector: SheetSelector<'_>) -> Error {
    match selector {
        SheetSelector::Name(name) => {
            Error::ParseError(format!("Numbers sheet named {name:?} not found"))
        },
        SheetSelector::Index(index) => Error::ParseError(format!(
            "Numbers sheet catalog index {index} is out of bounds"
        )),
    }
}

fn table_selector_error(selector: TableSelector<'_>) -> Error {
    match selector {
        TableSelector::Name(name) => {
            Error::ParseError(format!("Numbers table named {name:?} not found"))
        },
        TableSelector::Index(index) => Error::ParseError(format!(
            "Numbers table catalog index {index} is out of bounds"
        )),
    }
}

/// Resolve a semantic sheet selector to its native object identifier.
pub(super) fn sheet_id(editor: &NumbersEditor, selector: SheetSelector<'_>) -> Result<u64> {
    SheetSelectorAdapter::from_editor(editor)?.native_id(selector)
}

/// Resolve a semantic table selector to its native model object identifier.
pub(super) fn table_id(editor: &NumbersEditor, selector: TableSelector<'_>) -> Result<u64> {
    TableSelectorAdapter::from_editor(editor)?.native_id(selector)
}

/// Return the semantic catalog position of a native table identifier for
/// adapter-internal follow-up operations.
pub(super) fn table_index(editor: &NumbersEditor, native_id: u64) -> Result<usize> {
    super::table_models(&editor.package)?
        .iter()
        .position(|table| table.object_id == native_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers table {native_id} not found")))
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

    #[test]
    fn selectors_keep_exact_case_and_ambiguous_operations_atomic() {
        let mut editor = NumbersDocumentBuilder::new()
            .sheet_name("Summary")
            .table_name("Revenue")
            .table_dimensions(2, 2)
            .build()
            .unwrap();

        assert!(sheet_id(&editor, SheetSelector::name("summary")).is_err());
        assert!(table_id(&editor, TableSelector::name("revenue")).is_err());

        editor.add_empty_sheet("Archive").unwrap();
        editor
            .add_empty_table(SheetSelector::name("Archive"), "Revenue", 2, 2)
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .duplicate_table(TableSelector::name("Revenue"))
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
