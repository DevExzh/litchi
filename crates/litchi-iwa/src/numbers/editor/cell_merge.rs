//! Typed native table-cell merge storage shared by the iWork suite.

use std::num::NonZeroUsize;

use prost::Message;

use super::*;
use crate::numbers::formula_owner::{formula_owner_uuid_for_table, uuid_as_cfuuid};

mod axis;
mod formula;
mod wire;

#[cfg(test)]
use formula::parse_merge_formula;
use formula::{
    merge_formula, parse_regions, parse_table_uuid, rewrite_formula_region, validate_region_bounds,
};
use wire::{
    add_formula_store, add_merge_owner, append_formula, mutate_formulas, patch_table_model,
    remove_formula, remove_merge_owner, transform_formula_store, transform_merge_owner,
};

pub(crate) use axis::{MergeAnchorRelocation, MergeAxis};
use axis::{
    MergeDeletion, anchor_relocation_after_deletion, region_after_deletion, region_after_insertion,
};

#[derive(Debug)]
struct MergeFormulaMutation {
    formula_index: u32,
    previous: tsce::FormulaArchive,
    current: Option<tsce::FormulaArchive>,
}

/// A validated rectangular region containing at least two table cells.
///
/// Coordinates are zero-based and counts are non-zero. Construction also
/// rejects coordinate overflow, so all inclusive end-coordinate accessors are
/// infallible.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct IWorkTableCellRegion {
    row: usize,
    column: usize,
    row_count: NonZeroUsize,
    column_count: NonZeroUsize,
}

impl IWorkTableCellRegion {
    /// Construct a mergeable rectangular cell region.
    pub fn new(row: usize, column: usize, row_count: usize, column_count: usize) -> Result<Self> {
        let row_count = NonZeroUsize::new(row_count)
            .ok_or_else(|| Error::ParseError("Table-cell region has no rows".to_owned()))?;
        let column_count = NonZeroUsize::new(column_count)
            .ok_or_else(|| Error::ParseError("Table-cell region has no columns".to_owned()))?;
        if row_count.get() == 1 && column_count.get() == 1 {
            return Err(Error::ParseError(
                "A table-cell merge must contain at least two cells".to_owned(),
            ));
        }
        row.checked_add(row_count.get() - 1)
            .ok_or_else(|| Error::ParseError("Table-cell region row overflow".to_owned()))?;
        column
            .checked_add(column_count.get() - 1)
            .ok_or_else(|| Error::ParseError("Table-cell region column overflow".to_owned()))?;
        Ok(Self {
            row,
            column,
            row_count,
            column_count,
        })
    }

    #[must_use]
    pub const fn row(self) -> usize {
        self.row
    }

    #[must_use]
    pub const fn column(self) -> usize {
        self.column
    }

    #[must_use]
    pub const fn row_count(self) -> usize {
        self.row_count.get()
    }

    #[must_use]
    pub const fn column_count(self) -> usize {
        self.column_count.get()
    }

    #[must_use]
    pub const fn end_row(self) -> usize {
        self.row + self.row_count.get() - 1
    }

    #[must_use]
    pub const fn end_column(self) -> usize {
        self.column + self.column_count.get() - 1
    }

    #[must_use]
    pub const fn contains(self, row: usize, column: usize) -> bool {
        row >= self.row
            && row <= self.end_row()
            && column >= self.column
            && column <= self.end_column()
    }

    #[must_use]
    pub const fn overlaps(self, other: Self) -> bool {
        self.row <= other.end_row()
            && other.row <= self.end_row()
            && self.column <= other.end_column()
            && other.column <= self.end_column()
    }
}

pub(crate) fn regions_in_package(
    package: &IWorkPackage,
    table_id: u64,
) -> Result<Vec<IWorkTableCellRegion>> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    parse_regions(&descriptor.model)
}

/// Move or expand native merged-cell ranges after a physical table-axis
/// insertion.
///
/// Call this after the table's dimension has grown so a merge expanded at the
/// former trailing boundary remains within the model bounds during validation.
pub(super) fn shift_merges_for_axis_insertion(
    package: &mut IWorkPackage,
    table_id: u64,
    axis: MergeAxis,
    insertion: usize,
) -> Result<()> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    let existing = parse_regions(&descriptor.model)?;
    let Some(store) = descriptor
        .model
        .merge_owner
        .as_ref()
        .and_then(|owner| owner.formula_store.as_ref())
    else {
        return Ok(());
    };
    if store.formulas.is_empty() {
        return Ok(());
    }
    if store.formulas.len() != existing.len() {
        return Err(Error::InvalidFormat(
            "iWork merge formula storage and region count disagree".to_owned(),
        ));
    }

    let mut expected = Vec::with_capacity(existing.len());
    let mut rewrites = Vec::new();
    for (pair, region) in store.formulas.iter().zip(existing) {
        let updated = region_after_insertion(region, axis, insertion)?;
        if updated != region {
            rewrites.push(MergeFormulaMutation {
                formula_index: pair.formula_index,
                previous: pair.formula.clone(),
                current: Some(rewrite_formula_region(&pair.formula, updated)?),
            });
        }
        expected.push(updated);
    }
    if rewrites.is_empty() {
        return Ok(());
    }

    patch_table_model(package, table_id, |original| {
        transform_merge_owner(original, |owner_data| {
            transform_formula_store(owner_data, |store_data| {
                mutate_formulas(store_data, &rewrites)
            })
        })
    })?;

    if regions_in_package(package, table_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork merged-cell range insertion failed validation".to_owned(),
        ));
    }
    Ok(())
}

/// Identify merge anchors that native iWork carries into the next surviving
/// cell when a leading merge boundary is deleted.
///
/// Call this before physically compacting tiles. The returned source cells
/// must be relocated without releasing their value or comment references.
pub(super) fn merge_anchor_relocations_for_axis_deletion(
    package: &IWorkPackage,
    table_id: u64,
    axis: MergeAxis,
    deletion: usize,
) -> Result<Vec<MergeAnchorRelocation>> {
    let regions = parse_regions(&model::attached_table_descriptor(package, table_id)?.model)?;
    let mut relocations = Vec::new();
    for region in regions {
        if let Some(relocation) = anchor_relocation_after_deletion(region, axis, deletion)? {
            relocations.push(relocation);
        }
    }
    Ok(relocations)
}

/// Shift, contract, or remove native merged-cell ranges after a physical
/// table-axis deletion.
///
/// Call this before decreasing the table dimensions: a merge touching the
/// deleted trailing edge is temporarily out of bounds only after that change.
pub(super) fn shift_merges_for_axis_deletion(
    package: &mut IWorkPackage,
    table_id: u64,
    axis: MergeAxis,
    deletion: usize,
) -> Result<()> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    let existing = parse_regions(&descriptor.model)?;
    let Some(store) = descriptor
        .model
        .merge_owner
        .as_ref()
        .and_then(|owner| owner.formula_store.as_ref())
    else {
        return Ok(());
    };
    if store.formulas.is_empty() {
        return Ok(());
    }
    if store.formulas.len() != existing.len() {
        return Err(Error::InvalidFormat(
            "iWork merge formula storage and region count disagree".to_owned(),
        ));
    }

    let mut expected = Vec::with_capacity(existing.len());
    let mut mutations = Vec::new();
    for (pair, region) in store.formulas.iter().zip(existing) {
        match region_after_deletion(region, axis, deletion)? {
            MergeDeletion::Retain(updated) => {
                if updated != region {
                    mutations.push(MergeFormulaMutation {
                        formula_index: pair.formula_index,
                        previous: pair.formula.clone(),
                        current: Some(rewrite_formula_region(&pair.formula, updated)?),
                    });
                }
                expected.push(updated);
            },
            MergeDeletion::Remove => mutations.push(MergeFormulaMutation {
                formula_index: pair.formula_index,
                previous: pair.formula.clone(),
                current: None,
            }),
        }
    }
    if mutations.is_empty() {
        return Ok(());
    }

    patch_table_model(package, table_id, |original| {
        if expected.is_empty() {
            return remove_merge_owner(original);
        }
        transform_merge_owner(original, |owner_data| {
            transform_formula_store(owner_data, |store_data| {
                mutate_formulas(store_data, &mutations)
            })
        })
    })?;

    if regions_in_package(package, table_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork merged-cell range deletion failed validation".to_owned(),
        ));
    }
    Ok(())
}

pub(crate) fn merge_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: IWorkTableCellRegion,
) -> Result<()> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_region_bounds(&descriptor.model, region)?;
    let existing = parse_regions(&descriptor.model)?;
    if let Some(overlap) = existing.iter().find(|candidate| candidate.overlaps(region)) {
        return Err(Error::ParseError(format!(
            "Table-cell region {region:?} overlaps existing merge {overlap:?}"
        )));
    }

    let table_uuid = parse_table_uuid(&descriptor.model.table_id)?;
    let table_reference = uuid_as_cfuuid(&formula_owner_uuid_for_table(&table_uuid));
    let formula = merge_formula(region, table_reference)?;
    let store = descriptor
        .model
        .merge_owner
        .as_ref()
        .and_then(|owner| owner.formula_store.as_ref());
    let formula_index = store.map_or(0, |store| store.next_formula_index);
    if store.is_some_and(|store| {
        store
            .formulas
            .iter()
            .any(|pair| pair.formula_index == formula_index)
    }) {
        return Err(Error::InvalidFormat(format!(
            "iWork merge formula index {formula_index} is already occupied"
        )));
    }
    let next_formula_index = formula_index.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("iWork merge formula index space is exhausted".to_owned())
    })?;
    let pair = tst::formula_store_archive::FormulaStorePair {
        formula_index,
        formula,
    };
    let pair_data = pair.encode_to_vec();

    patch_table_model(package, table_id, |original| {
        if descriptor.model.merge_owner.is_some() {
            transform_merge_owner(original, |owner_data| {
                if store.is_some() {
                    transform_formula_store(owner_data, |store_data| {
                        append_formula(store_data, next_formula_index, pair_data)
                    })
                } else {
                    let formula_store = tst::FormulaStoreArchive {
                        next_formula_index,
                        formulas: vec![pair.clone()],
                    };
                    add_formula_store(owner_data, &formula_store)
                }
            })
        } else {
            let owner = tst::MergeOwnerArchive {
                owner_id: fresh_merge_owner_id(),
                formula_store: Some(tst::FormulaStoreArchive {
                    next_formula_index,
                    formulas: vec![pair.clone()],
                }),
            };
            add_merge_owner(original, &owner)
        }
    })?;

    let verified = regions_in_package(package, table_id)?;
    if verified.len() != existing.len() + 1 || !verified.contains(&region) {
        return Err(Error::InvalidFormat(
            "iWork table-cell merge failed validation".to_owned(),
        ));
    }
    Ok(())
}

pub(crate) fn unmerge_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    region: IWorkTableCellRegion,
) -> Result<bool> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    let existing = parse_regions(&descriptor.model)?;
    let Some(remove_position) = existing.iter().position(|candidate| *candidate == region) else {
        return Ok(false);
    };
    let owner =
        descriptor.model.merge_owner.as_ref().ok_or_else(|| {
            Error::InvalidFormat("iWork merge region has no merge owner".to_owned())
        })?;
    let store = owner.formula_store.as_ref().ok_or_else(|| {
        Error::InvalidFormat("iWork merge region has no formula store".to_owned())
    })?;
    let remove_index = store.formulas[remove_position].formula_index;

    patch_table_model(package, table_id, |original| {
        if store.formulas.len() == 1 {
            return remove_merge_owner(original);
        }
        transform_merge_owner(original, |owner_data| {
            transform_formula_store(owner_data, |store_data| {
                remove_formula(store_data, remove_index)
            })
        })
    })?;

    let verified = regions_in_package(package, table_id)?;
    if verified.len() + 1 != existing.len() || verified.contains(&region) {
        return Err(Error::InvalidFormat(
            "iWork table-cell unmerge failed validation".to_owned(),
        ));
    }
    Ok(true)
}

fn fresh_merge_owner_id() -> tsp::CfuuidArchive {
    let bytes = litchi_core::id::generate_guid_bytes();
    let mut lower = [0; 8];
    lower.copy_from_slice(&bytes[..8]);
    let mut upper = [0; 8];
    upper.copy_from_slice(&bytes[8..]);
    uuid_as_cfuuid(&tsp::Uuid {
        lower: u64::from_le_bytes(lower),
        upper: u64::from_le_bytes(upper),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::RawMessage;
    use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
    use crate::numbers::{
        FormulaCachedValue, FormulaCellReference, FormulaExpression, NumbersDocument,
        NumbersDocumentBuilder, TableColumnDeletion, TableColumnInsertion, TableRowDeletion,
        TableRowInsertion,
    };
    use crate::pages::{PagesDocumentBuilder, PagesEditor};
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::wire::{
        repeated_length_delimited_payloads, transform_length_delimited_fields_at_path,
    };

    #[test]
    fn region_validates_shape_overflow_and_overlap() {
        assert!(IWorkTableCellRegion::new(0, 0, 0, 2).is_err());
        assert!(IWorkTableCellRegion::new(0, 0, 1, 1).is_err());
        assert!(IWorkTableCellRegion::new(usize::MAX, 0, 2, 1).is_err());
        let region = IWorkTableCellRegion::new(2, 3, 2, 3).unwrap();
        assert_eq!((region.end_row(), region.end_column()), (3, 5));
        assert!(region.contains(3, 5));
        assert!(region.overlaps(IWorkTableCellRegion::new(3, 5, 2, 1).unwrap()));
        assert!(!region.overlaps(IWorkTableCellRegion::new(4, 3, 1, 2).unwrap()));
    }

    #[test]
    fn merge_formula_round_trips_exact_region() {
        let region = IWorkTableCellRegion::new(2, 3, 2, 4).unwrap();
        let table = tsp::CfuuidArchive {
            uuid_w0: Some(1),
            uuid_w1: Some(2),
            uuid_w2: Some(3),
            uuid_w3: Some(4),
            ..Default::default()
        };
        let formula = merge_formula(region, table.clone()).unwrap();
        assert_eq!(parse_merge_formula(&formula, &table).unwrap(), region);
    }

    #[test]
    fn scratch_numbers_table_merge_crud_is_transactional_and_byte_exact() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 5)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let baseline = editor.to_bytes().unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 3).unwrap();
        let second = IWorkTableCellRegion::new(0, 0, 1, 2).unwrap();

        editor.merge_cells(table_id, region).unwrap();
        editor.merge_cells(table_id, second).unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![region, second]
        );
        let merged = editor.to_bytes().unwrap();
        let mut reopened = NumbersEditor::from_bytes(&merged).unwrap();
        assert_eq!(
            reopened.table_cell_merges(table_id).unwrap(),
            vec![region, second]
        );

        let before_invalid = reopened.to_bytes().unwrap();
        assert!(
            reopened
                .merge_cells(table_id, IWorkTableCellRegion::new(2, 3, 2, 1).unwrap())
                .is_err()
        );
        assert_eq!(reopened.to_bytes().unwrap(), before_invalid);
        assert!(reopened.unmerge_cells(table_id, region).unwrap());
        assert!(!reopened.unmerge_cells(table_id, region).unwrap());
        assert_eq!(reopened.table_cell_merges(table_id).unwrap(), vec![second]);
        let mut reopened = NumbersEditor::from_bytes(&reopened.to_bytes().unwrap()).unwrap();
        assert!(reopened.unmerge_cells(table_id, second).unwrap());
        assert_eq!(reopened.to_bytes().unwrap(), baseline);
    }

    #[test]
    fn scratch_numbers_merged_table_axis_insertions_follow_native_semantics() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let original = IWorkTableCellRegion::new(1, 1, 2, 3).unwrap();

        editor.merge_cells(table_id, original).unwrap();
        editor
            .insert_table_row(table_id, TableRowInsertion::body(0))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(2, 1, 2, 3).unwrap()]
        );
        editor
            .insert_table_row(table_id, TableRowInsertion::body(2))
            .unwrap();
        editor
            .insert_table_column(table_id, TableColumnInsertion::body(0))
            .unwrap();
        editor
            .insert_table_column(table_id, TableColumnInsertion::body(2))
            .unwrap();

        let expected = IWorkTableCellRegion::new(2, 2, 3, 4).unwrap();
        assert_eq!(editor.table_cell_merges(table_id).unwrap(), vec![expected]);
        assert_eq!(editor.tables().unwrap()[0].rows, 7);
        assert_eq!(editor.tables().unwrap()[0].columns, 8);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_merges(table_id).unwrap(),
            vec![expected]
        );
    }

    #[test]
    fn scratch_numbers_merged_table_axis_deletions_relocate_anchor_content() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let region = IWorkTableCellRegion::new(1, 1, 3, 3).unwrap();
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Merged".to_owned()))
            .unwrap();
        editor
            .set_cell_comment(table_id, 1, 1, "Merged anchor")
            .unwrap();
        editor.merge_cells(table_id, region).unwrap();

        editor
            .remove_table_row(table_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 2, 3).unwrap()]
        );
        assert_eq!(
            editor
                .cell_comment(table_id, 1, 1)
                .unwrap()
                .unwrap()
                .comment
                .text,
            "Merged anchor"
        );

        editor
            .remove_table_column(table_id, TableColumnDeletion::body(0))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 2, 2).unwrap()]
        );
        editor
            .remove_table_row(table_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        editor
            .remove_table_column(table_id, TableColumnDeletion::body(0))
            .unwrap();
        assert!(editor.table_cell_merges(table_id).unwrap().is_empty());
        assert_eq!(editor.tables().unwrap()[0].rows, 3);
        assert_eq!(editor.tables().unwrap()[0].columns, 4);

        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Text("Merged".to_owned()))
        );
        assert_eq!(
            editor
                .cell_comment(table_id, 1, 1)
                .unwrap()
                .unwrap()
                .comment
                .text,
            "Merged anchor"
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(reopened.table_cell_merges(table_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_numbers_merged_formula_anchor_survives_axis_deletions() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(table_id, 3, 4, CellValue::Number(7.0))
            .unwrap();
        editor
            .set_formula_with_cached_value(
                table_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::Number(7.0),
            )
            .unwrap();
        editor.merge_cells(table_id, region).unwrap();

        editor
            .remove_table_row(table_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula("=E3".to_owned()))
        );

        editor
            .remove_table_column(table_id, TableColumnDeletion::body(0))
            .unwrap();
        assert!(editor.table_cell_merges(table_id).unwrap().is_empty());
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula("=D3".to_owned()))
        );
        editor
            .set_cell(table_id, 2, 3, CellValue::Number(11.0))
            .unwrap();
        assert_eq!(cached_formula_number(&editor, table_id, 1, 1), 11.0);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(reopened.table_cell_merges(table_id).unwrap().is_empty());
    }

    #[test]
    fn merged_table_axis_deletion_removes_and_rewrites_formula_pairs_together() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let collapsed = IWorkTableCellRegion::new(1, 1, 1, 2).unwrap();
        let shifted = IWorkTableCellRegion::new(2, 2, 2, 2).unwrap();
        editor.merge_cells(table_id, collapsed).unwrap();
        editor.merge_cells(table_id, shifted).unwrap();

        editor
            .remove_table_column(table_id, TableColumnDeletion::body(0))
            .unwrap();

        let expected = IWorkTableCellRegion::new(2, 1, 2, 2).unwrap();
        assert_eq!(editor.table_cell_merges(table_id).unwrap(), vec![expected]);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened.table_cell_merges(table_id).unwrap(),
            vec![expected]
        );
    }

    #[test]
    fn merged_table_axis_insertion_preserves_unknown_formula_wire() {
        const MERGE_OWNER_FIELD: u32 = 47;
        const FORMULA_STORE_FIELD: u32 = 2;
        const FORMULAS_FIELD: u32 = 3;
        const FORMULA_FIELD: u32 = 2;
        const AST_ARRAY_FIELD: u32 = 1;
        const AST_NODES_FIELD: u32 = 1;
        const COLON_TRACT_FIELD: u32 = 40;
        const ABSOLUTE_ROWS_FIELD: u32 = 4;
        const PAIR_UNKNOWN_FIELD: u32 = 90;
        const FORMULA_UNKNOWN_FIELD: u32 = 89;
        const NODE_UNKNOWN_FIELD: u32 = 88;
        const TRACT_UNKNOWN_FIELD: u32 = 87;
        const RANGE_UNKNOWN_FIELD: u32 = 86;
        const PAIR_UNKNOWN_VALUE: u64 = 900;
        const FORMULA_UNKNOWN_VALUE: u64 = 890;
        const NODE_UNKNOWN_VALUE: u64 = 880;
        const TRACT_UNKNOWN_VALUE: u64 = 870;
        const RANGE_UNKNOWN_VALUE: u64 = 860;

        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(4, 4)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .merge_cells(table_id, IWorkTableCellRegion::new(1, 1, 2, 2).unwrap())
            .unwrap();
        let mut package = editor.into_package();
        let archive_name = super::super::object_locations(&package)
            .unwrap()
            .remove(&table_id)
            .unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(table_id).unwrap();
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| matches!(message.type_, 6_000 | 6_001))
                    .unwrap();
                let message = object.messages[message_index].clone();
                let mut data = message.data;
                for (path, field, value) in [
                    (
                        &[MERGE_OWNER_FIELD, FORMULA_STORE_FIELD, FORMULAS_FIELD][..],
                        PAIR_UNKNOWN_FIELD,
                        PAIR_UNKNOWN_VALUE,
                    ),
                    (
                        &[
                            MERGE_OWNER_FIELD,
                            FORMULA_STORE_FIELD,
                            FORMULAS_FIELD,
                            FORMULA_FIELD,
                        ][..],
                        FORMULA_UNKNOWN_FIELD,
                        FORMULA_UNKNOWN_VALUE,
                    ),
                    (
                        &[
                            MERGE_OWNER_FIELD,
                            FORMULA_STORE_FIELD,
                            FORMULAS_FIELD,
                            FORMULA_FIELD,
                            AST_ARRAY_FIELD,
                            AST_NODES_FIELD,
                        ][..],
                        NODE_UNKNOWN_FIELD,
                        NODE_UNKNOWN_VALUE,
                    ),
                    (
                        &[
                            MERGE_OWNER_FIELD,
                            FORMULA_STORE_FIELD,
                            FORMULAS_FIELD,
                            FORMULA_FIELD,
                            AST_ARRAY_FIELD,
                            AST_NODES_FIELD,
                            COLON_TRACT_FIELD,
                        ][..],
                        TRACT_UNKNOWN_FIELD,
                        TRACT_UNKNOWN_VALUE,
                    ),
                    (
                        &[
                            MERGE_OWNER_FIELD,
                            FORMULA_STORE_FIELD,
                            FORMULAS_FIELD,
                            FORMULA_FIELD,
                            AST_ARRAY_FIELD,
                            AST_NODES_FIELD,
                            COLON_TRACT_FIELD,
                            ABSOLUTE_ROWS_FIELD,
                        ][..],
                        RANGE_UNKNOWN_FIELD,
                        RANGE_UNKNOWN_VALUE,
                    ),
                ] {
                    data = transform_length_delimited_fields_at_path(&data, path, |payload| {
                        let mut payload = payload.to_vec();
                        append_unknown_varint(&mut payload, field, value);
                        Ok(payload)
                    })?;
                }
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message.type_,
                        data,
                    },
                )?;
                Ok(())
            })
            .unwrap();
        let mut editor = NumbersEditor::from_package(package).unwrap();

        editor
            .insert_table_row(table_id, TableRowInsertion::body(1))
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 3, 2).unwrap()]
        );

        let archive = editor.package().archive(&archive_name).unwrap();
        let object = archive.object(table_id).unwrap();
        let message = object
            .messages
            .iter()
            .find(|message| matches!(message.type_, 6_000 | 6_001))
            .unwrap();
        let merge_owner = repeated_length_delimited_payloads(&message.data, MERGE_OWNER_FIELD)
            .unwrap()
            .remove(0);
        let formula_store = repeated_length_delimited_payloads(merge_owner, FORMULA_STORE_FIELD)
            .unwrap()
            .remove(0);
        let pair = repeated_length_delimited_payloads(formula_store, FORMULAS_FIELD)
            .unwrap()
            .remove(0);
        let formula = repeated_length_delimited_payloads(pair, FORMULA_FIELD)
            .unwrap()
            .remove(0);
        let ast_array = repeated_length_delimited_payloads(formula, AST_ARRAY_FIELD)
            .unwrap()
            .remove(0);
        let range_node = repeated_length_delimited_payloads(ast_array, AST_NODES_FIELD)
            .unwrap()
            .remove(0);
        let tract = repeated_length_delimited_payloads(range_node, COLON_TRACT_FIELD)
            .unwrap()
            .remove(0);
        let row_range = repeated_length_delimited_payloads(tract, ABSOLUTE_ROWS_FIELD)
            .unwrap()
            .remove(0);

        assert!(pair.ends_with(&unknown_suffix(PAIR_UNKNOWN_FIELD, PAIR_UNKNOWN_VALUE)));
        assert!(formula.ends_with(&unknown_suffix(
            FORMULA_UNKNOWN_FIELD,
            FORMULA_UNKNOWN_VALUE
        )));
        assert!(range_node.ends_with(&unknown_suffix(NODE_UNKNOWN_FIELD, NODE_UNKNOWN_VALUE)));
        assert!(tract.ends_with(&unknown_suffix(TRACT_UNKNOWN_FIELD, TRACT_UNKNOWN_VALUE)));
        assert!(row_range.ends_with(&unknown_suffix(RANGE_UNKNOWN_FIELD, RANGE_UNKNOWN_VALUE)));
    }

    #[test]
    fn scratch_pages_table_merge_crud_round_trips() {
        let mut editor = PagesDocumentBuilder::new()
            .body_text("Merged table\n")
            .body_table("Merge", 4, 5)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].model_object_id;
        let region = IWorkTableCellRegion::new(1, 2, 3, 2).unwrap();

        editor.merge_table_cells(table_id, region).unwrap();
        assert_eq!(editor.table(table_id).unwrap().merges, vec![region]);
        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.table_cell_merges(table_id).unwrap(), vec![region]);
        assert!(reopened.unmerge_table_cells(table_id, region).unwrap());
        assert!(reopened.table_cell_merges(table_id).unwrap().is_empty());
    }

    #[test]
    fn scratch_pages_and_keynote_merged_table_axis_insertions_round_trip() {
        let mut pages = PagesDocumentBuilder::new()
            .body_text("Merged table\n")
            .body_table("Merge", 4, 5)
            .build()
            .unwrap();
        let pages_table_id = pages.tables().unwrap()[0].model_object_id;
        pages
            .merge_table_cells(
                pages_table_id,
                IWorkTableCellRegion::new(1, 1, 2, 2).unwrap(),
            )
            .unwrap();
        pages
            .insert_table_row(pages_table_id, TableRowInsertion::body(1))
            .unwrap();
        pages
            .insert_table_column(pages_table_id, TableColumnInsertion::body(0))
            .unwrap();
        let pages_expected = IWorkTableCellRegion::new(1, 2, 3, 2).unwrap();
        assert_eq!(
            pages.table_cell_merges(pages_table_id).unwrap(),
            vec![pages_expected]
        );
        let pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages.table_cell_merges(pages_table_id).unwrap(),
            vec![pages_expected]
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let keynote_table = keynote
            .add_slide_table(
                0,
                "Merge",
                4,
                5,
                DrawablePoint { x: 100.0, y: 150.0 },
                DrawableSize {
                    width: 800.0,
                    height: 400.0,
                },
            )
            .unwrap();
        keynote
            .merge_slide_table_cells(
                0,
                keynote_table.model_object_id,
                IWorkTableCellRegion::new(1, 1, 2, 2).unwrap(),
            )
            .unwrap();
        keynote
            .insert_slide_table_row(0, keynote_table.model_object_id, TableRowInsertion::body(1))
            .unwrap();
        keynote
            .insert_slide_table_column(
                0,
                keynote_table.model_object_id,
                TableColumnInsertion::body(0),
            )
            .unwrap();
        let keynote_expected = IWorkTableCellRegion::new(1, 2, 3, 2).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap(),
            vec![keynote_expected]
        );
        let keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap(),
            vec![keynote_expected]
        );
    }

    #[test]
    fn scratch_pages_and_keynote_merged_table_axis_deletions_relocate_anchor_content() {
        let mut pages = PagesDocumentBuilder::new()
            .body_text("Merged table\n")
            .body_table("Merge", 4, 5)
            .build()
            .unwrap();
        let pages_table_id = pages.tables().unwrap()[0].model_object_id;
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        pages
            .set_table_cell(
                pages_table_id,
                region.row(),
                region.column(),
                CellValue::Text("Merged".to_owned()),
            )
            .unwrap();
        pages.merge_table_cells(pages_table_id, region).unwrap();
        pages
            .remove_table_row(pages_table_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            pages.table_cell_merges(pages_table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        pages
            .remove_table_column(pages_table_id, TableColumnDeletion::body(0))
            .unwrap();
        assert!(pages.table_cell_merges(pages_table_id).unwrap().is_empty());
        assert_eq!(
            pages.table(pages_table_id).unwrap().get_cell(1, 1),
            Some(&CellValue::Text("Merged".to_owned()))
        );
        let pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert!(pages.table_cell_merges(pages_table_id).unwrap().is_empty());

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let keynote_table = keynote
            .add_slide_table(
                0,
                "Merge",
                4,
                5,
                DrawablePoint { x: 100.0, y: 150.0 },
                DrawableSize {
                    width: 800.0,
                    height: 400.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                keynote_table.model_object_id,
                region.row(),
                region.column(),
                CellValue::Text("Merged".to_owned()),
            )
            .unwrap();
        keynote
            .merge_slide_table_cells(0, keynote_table.model_object_id, region)
            .unwrap();
        keynote
            .remove_slide_table_row(0, keynote_table.model_object_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        keynote
            .remove_slide_table_column(
                0,
                keynote_table.model_object_id,
                TableColumnDeletion::body(0),
            )
            .unwrap();
        assert!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap()
                .is_empty()
        );
        assert_eq!(
            keynote
                .slide_table(0, keynote_table.model_object_id)
                .unwrap()
                .get_cell(1, 1),
            Some(&CellValue::Text("Merged".to_owned()))
        );
        let keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn scratch_pages_and_keynote_merged_formula_anchors_survive_axis_deletions() {
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Merge Formula", 5, 6)
            .build()
            .unwrap();
        let pages_table_id = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(pages_table_id, 3, 4, CellValue::Number(7.0))
            .unwrap();
        pages
            .set_table_formula(
                pages_table_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::Number(7.0),
            )
            .unwrap();
        pages.merge_table_cells(pages_table_id, region).unwrap();
        pages
            .remove_table_row(pages_table_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            pages.table_formula(pages_table_id, 1, 1).unwrap(),
            Some("=E3".to_owned())
        );
        pages
            .remove_table_column(pages_table_id, TableColumnDeletion::body(0))
            .unwrap();
        assert!(pages.table_cell_merges(pages_table_id).unwrap().is_empty());
        assert_eq!(
            pages.table_formula(pages_table_id, 1, 1).unwrap(),
            Some("=D3".to_owned())
        );
        let pages = PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages.table_formula(pages_table_id, 1, 1).unwrap(),
            Some("=D3".to_owned())
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let keynote_table = keynote
            .add_slide_table(
                0,
                "Merge Formula",
                5,
                6,
                DrawablePoint { x: 100.0, y: 150.0 },
                DrawableSize {
                    width: 800.0,
                    height: 400.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                keynote_table.model_object_id,
                3,
                4,
                CellValue::Number(7.0),
            )
            .unwrap();
        keynote
            .set_slide_table_formula(
                0,
                keynote_table.model_object_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::Number(7.0),
            )
            .unwrap();
        keynote
            .merge_slide_table_cells(0, keynote_table.model_object_id, region)
            .unwrap();
        keynote
            .remove_slide_table_row(0, keynote_table.model_object_id, TableRowDeletion::body(0))
            .unwrap();
        assert_eq!(
            keynote
                .slide_table_formula(0, keynote_table.model_object_id, 1, 1)
                .unwrap(),
            Some("=E3".to_owned())
        );
        keynote
            .remove_slide_table_column(
                0,
                keynote_table.model_object_id,
                TableColumnDeletion::body(0),
            )
            .unwrap();
        assert!(
            keynote
                .slide_table_cell_merges(0, keynote_table.model_object_id)
                .unwrap()
                .is_empty()
        );
        assert_eq!(
            keynote
                .slide_table_formula(0, keynote_table.model_object_id, 1, 1)
                .unwrap(),
            Some("=D3".to_owned())
        );
        let keynote = KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_formula(0, keynote_table.model_object_id, 1, 1)
                .unwrap(),
            Some("=D3".to_owned())
        );
    }

    #[test]
    fn scratch_keynote_table_merge_crud_round_trips() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let table = editor
            .add_slide_table(
                0,
                "Merge",
                4,
                5,
                DrawablePoint { x: 100.0, y: 150.0 },
                DrawableSize {
                    width: 800.0,
                    height: 400.0,
                },
            )
            .unwrap();
        let region = IWorkTableCellRegion::new(2, 1, 1, 3).unwrap();

        editor
            .merge_slide_table_cells(0, table.model_object_id, region)
            .unwrap();
        assert_eq!(
            editor.slide_table(0, table.model_object_id).unwrap().merges,
            vec![region]
        );
        let mut reopened = KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(
            reopened
                .unmerge_slide_table_cells(0, table.model_object_id, region)
                .unwrap()
        );
        assert!(
            reopened
                .slide_table_cell_merges(0, table.model_object_id)
                .unwrap()
                .is_empty()
        );
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field) << 3));
        data.extend(crate::varint::encode_varint(value));
    }

    fn cached_formula_number(
        editor: &NumbersEditor,
        table_id: u64,
        row: usize,
        column: usize,
    ) -> f64 {
        let location = locate_attached_cell(editor.package(), table_id, row, column).unwrap();
        let data = read_tile_cell(
            editor.package(),
            &location.tile_archive,
            location.tile_id,
            location.tile_row,
            column,
        )
        .unwrap()
        .unwrap();
        match BncCell::parse(&data).unwrap().cached_scalar().unwrap() {
            Some(crate::numbers::bnc::CachedScalar::Number(value)) => value,
            value => panic!("Expected numeric formula cache, found {value:?}"),
        }
    }

    fn unknown_suffix(field: u32, value: u64) -> Vec<u8> {
        let mut suffix = Vec::new();
        append_unknown_varint(&mut suffix, field, value);
        suffix
    }
}
