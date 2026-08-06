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
    use crate::archive::{ArchiveObject, RawMessage};
    use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
    use crate::numbers::{
        FormulaAxisReference, FormulaCachedValue, FormulaCellReference, FormulaExpression,
        NumbersDocument, NumbersDocumentBuilder, TableColumnDeletion, TableColumnInsertion,
        TableRowDeletion, TableRowInsertion,
    };
    use crate::pages::{PagesDocumentBuilder, PagesEditor};
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::wire::{
        repeated_length_delimited_payloads, transform_length_delimited_fields_at_path,
    };

    const RANGE_PROXY_OWNER_KIND: u32 = 5;
    const RANGE_PROXY_UUID_LOWER: u64 = 0x4d45_5247_4550_524f;
    const RANGE_PROXY_UUID_UPPER: u64 = 0x5859_5445_5354_0001;

    #[derive(Clone, Copy)]
    struct TestRangeBounds {
        top: u32,
        left: u32,
        bottom: u32,
        right: u32,
    }

    impl TestRangeBounds {
        const fn new(top: u32, left: u32, bottom: u32, right: u32) -> Self {
            Self {
                top,
                left,
                bottom,
                right,
            }
        }
    }

    fn range_edge_coordinates(bounds: TestRangeBounds) -> (Vec<u32>, Vec<u32>) {
        let mut rows = Vec::new();
        let mut columns = Vec::new();
        for row in bounds.top..=bounds.bottom {
            for column in bounds.left..=bounds.right {
                rows.push(row);
                columns.push(column);
            }
        }
        (rows, columns)
    }

    #[derive(Clone, Copy)]
    struct NativeRangeDependencyIds {
        source_owner_id: u32,
        external_owner_id: u32,
    }

    #[derive(Clone, Copy)]
    struct NativeMergeRangeProxyIds {
        object_id: u64,
        range_tile_id: u64,
        internal_owner_id: u32,
    }

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
            .insert_table_row(
                test_table_selector(&editor, table_id),
                TableRowInsertion::body(0),
            )
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(2, 1, 2, 3).unwrap()]
        );
        editor
            .insert_table_row(
                test_table_selector(&editor, table_id),
                TableRowInsertion::body(2),
            )
            .unwrap();
        editor
            .insert_table_column(
                test_table_selector(&editor, table_id),
                TableColumnInsertion::body(0),
            )
            .unwrap();
        editor
            .insert_table_column(
                test_table_selector(&editor, table_id),
                TableColumnInsertion::body(2),
            )
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
            .remove_table_row(
                test_table_selector(&editor, table_id),
                TableRowDeletion::body(0),
            )
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
            .remove_table_column(
                test_table_selector(&editor, table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 2, 2).unwrap()]
        );
        editor
            .remove_table_row(
                test_table_selector(&editor, table_id),
                TableRowDeletion::body(0),
            )
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        editor
            .remove_table_column(
                test_table_selector(&editor, table_id),
                TableColumnDeletion::body(0),
            )
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
            .set_cell(table_id, 3, 4, CellValue::number(7.0).unwrap())
            .unwrap();
        editor
            .set_formula_with_cached_value(
                table_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(table_id, region).unwrap();

        editor
            .remove_table_row(
                test_table_selector(&editor, table_id),
                TableRowDeletion::body(0),
            )
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
            .remove_table_column(
                test_table_selector(&editor, table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        assert!(editor.table_cell_merges(table_id).unwrap().is_empty());
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula("=D3".to_owned()))
        );
        editor
            .set_cell(table_id, 2, 3, CellValue::number(11.0).unwrap())
            .unwrap();
        assert_eq!(cached_formula_number(&editor, table_id, 1, 1), 11.0);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert!(reopened.table_cell_merges(table_id).unwrap().is_empty());
    }

    #[test]
    fn merged_cross_table_formula_anchor_rebases_uuid_hosts_and_external_edges() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let target_table = editor
            .add_empty_table(test_sheet_selector(&editor, sheet_id), "Referenced", 5, 6)
            .unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(
                target_table.object_id,
                1,
                0,
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor
            .set_formula_with_cached_value(
                source_table_id,
                region.row(),
                region.column(),
                FormulaExpression::table_cell(
                    target_table.object_id,
                    FormulaCellReference::relative(1, 0),
                ),
                FormulaCachedValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(source_table_id, region).unwrap();

        let mut package = editor.into_package();
        install_uuid_host_references(&mut package, region.row() as u32, region.column() as u32);
        let mut editor = NumbersEditor::from_package(package).unwrap();
        let external_owner_id = formula_owner_at_host(&editor, 1, 1)
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record[0]
            .expanded_edges
            .as_ref()
            .unwrap()
            .internal_owner_id_for_edge[0];

        editor
            .remove_table_row(
                test_table_selector(&editor, source_table_id),
                TableRowDeletion::body(0),
            )
            .unwrap();
        assert_eq!(
            editor.table_cell_merges(source_table_id).unwrap(),
            vec![IWorkTableCellRegion::new(1, 1, 1, 2).unwrap()]
        );
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula("=Sheet 1::Referenced::A3".to_owned()))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 0.0);
        assert_formula_host_dependencies(&editor, external_owner_id, 2, 0);

        editor
            .remove_table_column(
                test_table_selector(&editor, source_table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        assert!(
            editor
                .table_cell_merges(source_table_id)
                .unwrap()
                .is_empty()
        );
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula("=Sheet 1::Referenced::B3".to_owned()))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 0.0);
        assert_formula_host_dependencies(&editor, external_owner_id, 2, 1);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_formula_host_dependencies(&reopened, external_owner_id, 2, 1);
    }

    #[test]
    fn merged_cross_table_range_anchor_rebases_expanded_edges_and_cache() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let target_table = editor
            .add_empty_table(test_sheet_selector(&editor, sheet_id), "Referenced", 5, 6)
            .unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(
                target_table.object_id,
                1,
                0,
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor
            .set_cell(
                target_table.object_id,
                2,
                0,
                CellValue::number(3.0).unwrap(),
            )
            .unwrap();
        editor
            .set_formula_with_cached_value(
                source_table_id,
                region.row(),
                region.column(),
                FormulaExpression::function(
                    "SUM",
                    [FormulaExpression::table_range(
                        target_table.object_id,
                        FormulaCellReference::relative(1, 0),
                        FormulaCellReference::relative(2, 0),
                    )],
                ),
                FormulaCachedValue::number(10.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(source_table_id, region).unwrap();
        let mut package = editor.into_package();
        install_uuid_host_references(&mut package, region.row() as u32, region.column() as u32);
        let mut editor = NumbersEditor::from_package(package).unwrap();
        let external_owner_id = formula_owner_at_host(&editor, 1, 1)
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record[0]
            .expanded_edges
            .as_ref()
            .unwrap()
            .internal_owner_id_for_edge[0];

        editor
            .remove_table_row(
                test_table_selector(&editor, source_table_id),
                TableRowDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::A3:A4)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 3.0);
        assert_formula_host_range_edges(&editor, external_owner_id, &[2, 3], &[0, 0]);

        editor
            .remove_table_column(
                test_table_selector(&editor, source_table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::B3:B4)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 0.0);
        assert_formula_host_range_edges(&editor, external_owner_id, &[2, 3], &[1, 1]);
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_formula_host_range_edges(&reopened, external_owner_id, &[2, 3], &[1, 1]);
    }

    #[test]
    fn merged_cross_table_range_anchor_rebases_native_range_records_and_cache() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let target_table = editor
            .add_empty_table(test_sheet_selector(&editor, sheet_id), "Referenced", 5, 6)
            .unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(
                target_table.object_id,
                1,
                0,
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor
            .set_cell(
                target_table.object_id,
                2,
                0,
                CellValue::number(3.0).unwrap(),
            )
            .unwrap();
        editor
            .set_formula_with_cached_value(
                source_table_id,
                region.row(),
                region.column(),
                FormulaExpression::function(
                    "SUM",
                    [FormulaExpression::table_range(
                        target_table.object_id,
                        FormulaCellReference::relative(1, 0),
                        FormulaCellReference::relative(2, 0),
                    )],
                ),
                FormulaCachedValue::number(10.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(source_table_id, region).unwrap();

        let mut package = editor.into_package();
        install_uuid_host_references(&mut package, region.row() as u32, region.column() as u32);
        let dependency_ids = install_native_cross_table_range_dependencies(
            &mut package,
            region.row() as u32,
            region.column() as u32,
            TestRangeBounds::new(1, 0, 2, 0),
        );
        let proxy = install_native_merge_range_proxy(
            &mut package,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 2, 2),
        );
        let mut editor = NumbersEditor::from_package(package).unwrap();

        editor
            .remove_table_row(
                test_table_selector(&editor, source_table_id),
                TableRowDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::A3:A4)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 3.0);
        assert_formula_host_native_range_dependencies(
            &editor,
            dependency_ids.external_owner_id,
            TestRangeBounds::new(2, 0, 3, 0),
        );
        assert_native_merge_range_proxy(
            &editor,
            proxy,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 1, 2),
        );

        editor
            .remove_table_column(
                test_table_selector(&editor, source_table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::B3:B4)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 0.0);
        assert_formula_host_native_range_dependencies(
            &editor,
            dependency_ids.external_owner_id,
            TestRangeBounds::new(2, 1, 3, 1),
        );
        assert_native_merge_range_proxy(
            &editor,
            proxy,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 1, 1),
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_formula_host_native_range_dependencies(
            &reopened,
            dependency_ids.external_owner_id,
            TestRangeBounds::new(2, 1, 3, 1),
        );
        assert_native_merge_range_proxy(
            &reopened,
            proxy,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 1, 1),
        );
    }

    #[test]
    fn merged_cross_table_whole_row_anchor_rebases_expanded_edges_and_cache() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let target_table = editor
            .add_empty_table(test_sheet_selector(&editor, sheet_id), "Referenced", 5, 6)
            .unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(
                target_table.object_id,
                1,
                0,
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor
            .set_cell(
                target_table.object_id,
                2,
                0,
                CellValue::number(3.0).unwrap(),
            )
            .unwrap();
        editor
            .set_formula_with_cached_value(
                source_table_id,
                region.row(),
                region.column(),
                FormulaExpression::function(
                    "SUM",
                    [FormulaExpression::table_rows(
                        target_table.object_id,
                        FormulaAxisReference::relative(1),
                        FormulaAxisReference::relative(2),
                    )],
                ),
                FormulaCachedValue::number(10.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(source_table_id, region).unwrap();

        let mut package = editor.into_package();
        install_uuid_host_references(&mut package, region.row() as u32, region.column() as u32);
        let mut editor = NumbersEditor::from_package(package).unwrap();
        let external_owner_id = formula_owner_at_host(&editor, 1, 1)
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record[0]
            .expanded_edges
            .as_ref()
            .unwrap()
            .internal_owner_id_for_edge[0];
        let (initial_rows, initial_columns) =
            range_edge_coordinates(TestRangeBounds::new(1, 0, 2, 5));
        assert_formula_host_range_edges(
            &editor,
            external_owner_id,
            &initial_rows,
            &initial_columns,
        );

        editor
            .remove_table_row(
                test_table_selector(&editor, source_table_id),
                TableRowDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::3:4)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 3.0);
        let (rebased_rows, rebased_columns) =
            range_edge_coordinates(TestRangeBounds::new(2, 0, 3, 5));
        assert_formula_host_range_edges(
            &editor,
            external_owner_id,
            &rebased_rows,
            &rebased_columns,
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_formula_host_range_edges(
            &reopened,
            external_owner_id,
            &rebased_rows,
            &rebased_columns,
        );
    }

    #[test]
    fn merged_cross_table_whole_column_anchor_rebases_native_range_records_and_cache() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(5, 6)
            .build()
            .unwrap();
        let source_table_id = editor.tables().unwrap()[0].object_id;
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let target_table = editor
            .add_empty_table(test_sheet_selector(&editor, sheet_id), "Referenced", 5, 6)
            .unwrap();
        let region = IWorkTableCellRegion::new(1, 1, 2, 2).unwrap();
        editor
            .set_cell(
                target_table.object_id,
                0,
                1,
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        editor
            .set_cell(
                target_table.object_id,
                0,
                2,
                CellValue::number(3.0).unwrap(),
            )
            .unwrap();
        editor
            .set_formula_with_cached_value(
                source_table_id,
                region.row(),
                region.column(),
                FormulaExpression::function(
                    "SUM",
                    [FormulaExpression::table_columns(
                        target_table.object_id,
                        FormulaAxisReference::relative(1),
                        FormulaAxisReference::relative(2),
                    )],
                ),
                FormulaCachedValue::number(10.0).unwrap(),
            )
            .unwrap();
        editor.merge_cells(source_table_id, region).unwrap();

        let mut package = editor.into_package();
        install_uuid_host_references(&mut package, region.row() as u32, region.column() as u32);
        let dependency_ids = install_native_cross_table_range_dependencies(
            &mut package,
            region.row() as u32,
            region.column() as u32,
            TestRangeBounds::new(0, 1, 4, 2),
        );
        let proxy = install_native_merge_range_proxy(
            &mut package,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 2, 2),
        );
        let mut editor = NumbersEditor::from_package(package).unwrap();

        editor
            .remove_table_column(
                test_table_selector(&editor, source_table_id),
                TableColumnDeletion::body(0),
            )
            .unwrap();
        let document = NumbersDocument::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            document.sheets().unwrap()[0].tables[0].get_cell(1, 1),
            Some(&CellValue::Formula(
                "=SUM(Sheet 1::Referenced::C:D)".to_owned()
            ))
        );
        assert_eq!(cached_formula_number(&editor, source_table_id, 1, 1), 3.0);
        assert_formula_host_native_range_dependencies(
            &editor,
            dependency_ids.external_owner_id,
            TestRangeBounds::new(0, 2, 4, 3),
        );
        assert_native_merge_range_proxy(
            &editor,
            proxy,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 2, 1),
        );
        let reopened = NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_formula_host_native_range_dependencies(
            &reopened,
            dependency_ids.external_owner_id,
            TestRangeBounds::new(0, 2, 4, 3),
        );
        assert_native_merge_range_proxy(
            &reopened,
            proxy,
            dependency_ids.source_owner_id,
            TestRangeBounds::new(1, 1, 2, 1),
        );
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
            .remove_table_column(
                test_table_selector(&editor, table_id),
                TableColumnDeletion::body(0),
            )
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
            .insert_table_row(
                test_table_selector(&editor, table_id),
                TableRowInsertion::body(1),
            )
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
        assert_eq!(editor.table(table_id).unwrap().merges(), &[region]);
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
            .set_table_cell(pages_table_id, 3, 4, CellValue::number(7.0).unwrap())
            .unwrap();
        pages
            .set_table_formula(
                pages_table_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::number(7.0).unwrap(),
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
                CellValue::number(7.0).unwrap(),
            )
            .unwrap();
        keynote
            .set_slide_table_formula(
                0,
                keynote_table.model_object_id,
                region.row(),
                region.column(),
                FormulaExpression::cell(FormulaCellReference::relative(3, 4)),
                FormulaCachedValue::number(7.0).unwrap(),
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
            editor
                .slide_table(0, table.model_object_id)
                .unwrap()
                .merges(),
            &[region]
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
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }

    fn install_uuid_host_references(package: &mut IWorkPackage, row: u32, column: u32) {
        let component = package
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap()
            .to_owned();
        package
            .update_archive(&component, |archive| {
                let mut installed = false;
                for object in &mut archive.objects {
                    for message in &mut object.messages {
                        if message.type_ != 4_008 {
                            continue;
                        }
                        let mut owner =
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                        let contains_host =
                            owner
                                .cell_dependencies
                                .as_ref()
                                .is_some_and(|dependencies| {
                                    dependencies
                                        .cell_record
                                        .iter()
                                        .any(|record| record.row == row && record.column == column)
                                });
                        if !contains_host {
                            continue;
                        }
                        let owner_uuid = owner.formula_owner_uid;
                        owner.uuid_references = Some(tsce::UuidReferencesArchive {
                            table_refs: vec![tsce::uuid_references_archive::TableRef {
                                owner_uuid,
                                coord_set: Some(uuid_host_coordinate_set(row, column)),
                            }],
                            table_uuid_refs: vec![
                                tsce::uuid_references_archive::TableWithUuidRef {
                                    owner_uuid,
                                    uuid_refs: vec![tsce::uuid_references_archive::UuidRef {
                                        uuid: tsp::Uuid {
                                            lower: 0x0123_4567_89ab_cdef,
                                            upper: 0xfedc_ba98_7654_3210,
                                        },
                                        coord_set: Some(uuid_host_coordinate_set(row, column)),
                                    }],
                                },
                            ],
                        });
                        message.data = owner.encode_to_vec();
                        installed = true;
                    }
                }
                installed.then_some(()).ok_or_else(|| {
                    Error::InvalidFormat(
                        "Test formula owner has no matching dependency host".to_owned(),
                    )
                })
            })
            .unwrap();
    }

    fn install_native_cross_table_range_dependencies(
        package: &mut IWorkPackage,
        row: u32,
        column: u32,
        bounds: TestRangeBounds,
    ) -> NativeRangeDependencyIds {
        let component = package
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap()
            .to_owned();
        let mut dependency_ids = None;
        package
            .update_archive(&component, |archive| {
                let range_tile_id = archive
                    .objects
                    .iter()
                    .filter_map(|object| object.archive_info.identifier)
                    .max()
                    .and_then(|identifier| identifier.checked_add(1))
                    .ok_or_else(|| {
                        Error::ParseError(
                            "Test CalculationEngine cannot allocate a range tile identifier"
                                .to_owned(),
                        )
                    })?;
                let mut source = None;
                for object in &archive.objects {
                    let Some(object_id) = object.archive_info.identifier else {
                        continue;
                    };
                    for (message_index, message) in object.messages.iter().enumerate() {
                        if message.type_ != 4_008 {
                            continue;
                        }
                        let owner =
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
                        let contains_host =
                            owner
                                .cell_dependencies
                                .as_ref()
                                .is_some_and(|dependencies| {
                                    dependencies
                                        .cell_record
                                        .iter()
                                        .any(|record| record.row == row && record.column == column)
                                });
                        if contains_host {
                            source = Some((object_id, message_index, owner));
                            break;
                        }
                    }
                    if source.is_some() {
                        break;
                    }
                }
                let Some((source_object_id, message_index, mut owner)) = source else {
                    return Err(Error::InvalidFormat(
                        "Test formula owner has no matching range host".to_owned(),
                    ));
                };
                let host = owner
                    .cell_dependencies
                    .as_ref()
                    .and_then(|dependencies| {
                        dependencies
                            .cell_record
                            .iter()
                            .find(|record| record.row == row && record.column == column)
                    })
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Test formula owner has no matching dependency record".to_owned(),
                        )
                    })?;
                let edges = host.expanded_edges.as_ref().ok_or_else(|| {
                    Error::InvalidFormat(
                        "Test formula host has no expanded dependency edges".to_owned(),
                    )
                })?;
                let target_owner_id = edges
                    .internal_owner_id_for_edge
                    .first()
                    .copied()
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Test formula host has no cross-table dependency owner".to_owned(),
                        )
                    })?;
                if edges
                    .internal_owner_id_for_edge
                    .iter()
                    .any(|owner_id| *owner_id != target_owner_id)
                {
                    return Err(Error::InvalidFormat(
                        "Test formula host has mixed cross-table dependency owners".to_owned(),
                    ));
                }
                let cell_tile_ids = owner
                    .tiled_cell_dependencies
                    .as_ref()
                    .map(|dependencies| {
                        dependencies
                            .cell_record_tiles
                            .iter()
                            .map(|reference| reference.identifier)
                            .collect::<Vec<_>>()
                    })
                    .unwrap_or_default();
                let host_record = owner
                    .cell_dependencies
                    .as_mut()
                    .and_then(|dependencies| {
                        dependencies
                            .cell_record
                            .iter_mut()
                            .find(|record| record.row == row && record.column == column)
                    })
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Test formula owner has no mutable dependency record".to_owned(),
                        )
                    })?;
                host_record.expanded_edges = Some(tsce::ExpandedEdgesArchive::default());
                owner.range_dependencies = Some(tsce::RangeDependenciesArchive {
                    back_dependency: vec![tsce::RangeBackDependencyArchive {
                        cell_coord_row: row,
                        cell_coord_column: column,
                        internal_range_reference: Some(tsce::InternalRangeReferenceArchive {
                            owner_id: target_owner_id,
                            range: tsce::RangeCoordinateArchive {
                                top_left_column: bounds.left,
                                top_left_row: bounds.top,
                                bottom_right_column: bounds.right,
                                bottom_right_row: bounds.bottom,
                            },
                        }),
                        ..Default::default()
                    }],
                });
                owner.tiled_range_dependencies = Some(tsce::RangeDependenciesTiledArchive {
                    range_precedents_tile: vec![tsp::Reference {
                        identifier: range_tile_id,
                        ..Default::default()
                    }],
                });
                let source_object = archive.object_mut(source_object_id).ok_or_else(|| {
                    Error::InvalidFormat(
                        "Test formula owner object is missing during range setup".to_owned(),
                    )
                })?;
                let message_type = source_object.messages[message_index].type_;
                source_object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data: owner.encode_to_vec(),
                    },
                )?;
                source_object.archive_info.message_infos[message_index]
                    .object_references
                    .push(range_tile_id);

                let mut found_tiled_host = false;
                for tile_id in cell_tile_ids {
                    let tile_object = archive.object_mut(tile_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Test formula dependency tile {tile_id} is missing"
                        ))
                    })?;
                    let tile_message_index = tile_object
                        .messages
                        .iter()
                        .position(|message| message.type_ == 4_009)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "Test formula dependency tile {tile_id} has no payload"
                            ))
                        })?;
                    let message = tile_object.messages[tile_message_index].clone();
                    let mut tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
                    for record in &mut tile.cell_records {
                        if record.row == row && record.column == column {
                            record.expanded_edges = Some(tsce::ExpandedEdgesArchive::default());
                            found_tiled_host = true;
                        }
                    }
                    tile_object.replace_message(
                        tile_message_index,
                        RawMessage {
                            type_: message.type_,
                            data: tile.encode_to_vec(),
                        },
                    )?;
                }
                if !found_tiled_host {
                    return Err(Error::InvalidFormat(
                        "Test formula host is absent from its dependency tiles".to_owned(),
                    ));
                }
                archive.insert_object(ArchiveObject::new(
                    range_tile_id,
                    vec![RawMessage {
                        type_: 4_010,
                        data: tsce::RangePrecedentsTileArchive {
                            to_owner_id: target_owner_id,
                            from_to_range: vec![
                                tsce::range_precedents_tile_archive::FromToRangeArchive {
                                    from_coord: tsce::CellCoordinateArchive {
                                        column: Some(column),
                                        row: Some(row),
                                        ..Default::default()
                                    },
                                    refers_to_rect: tsce::CellRectArchive {
                                        origin: tsce::CellCoordinateArchive {
                                            column: Some(bounds.left),
                                            row: Some(bounds.top),
                                            ..Default::default()
                                        },
                                        size: tsce::ColumnRowSize {
                                            num_columns: (bounds.right != bounds.left)
                                                .then_some(bounds.right - bounds.left + 1),
                                            num_rows: (bounds.bottom != bounds.top)
                                                .then_some(bounds.bottom - bounds.top + 1),
                                        },
                                    },
                                },
                            ],
                        }
                        .encode_to_vec(),
                    }],
                )?)?;
                dependency_ids = Some(NativeRangeDependencyIds {
                    source_owner_id: owner.internal_formula_owner_id,
                    external_owner_id: target_owner_id,
                });
                Ok(())
            })
            .unwrap();
        dependency_ids.unwrap()
    }

    fn install_native_merge_range_proxy(
        package: &mut IWorkPackage,
        source_owner_id: u32,
        bounds: TestRangeBounds,
    ) -> NativeMergeRangeProxyIds {
        let component = package
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap()
            .to_owned();
        let mut proxy_ids = None;
        package
            .update_archive(&component, |archive| {
                let first_object_id = archive
                    .objects
                    .iter()
                    .filter_map(|object| object.archive_info.identifier)
                    .max()
                    .and_then(|identifier| identifier.checked_add(1))
                    .ok_or_else(|| {
                        Error::ParseError(
                            "Test CalculationEngine cannot allocate a range-proxy object ID"
                                .to_owned(),
                        )
                    })?;
                let range_tile_id = first_object_id.checked_add(1).ok_or_else(|| {
                    Error::ParseError(
                        "Test CalculationEngine cannot allocate a range-proxy tile ID".to_owned(),
                    )
                })?;
                let internal_owner_id = archive
                    .objects
                    .iter()
                    .flat_map(|object| &object.messages)
                    .filter(|message| message.type_ == 4_008)
                    .map(|message| {
                        Ok::<_, Error>(
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?
                                .internal_formula_owner_id,
                        )
                    })
                    .collect::<Result<Vec<_>>>()?
                    .into_iter()
                    .max()
                    .and_then(|identifier| identifier.checked_add(1))
                    .ok_or_else(|| {
                        Error::ParseError(
                            "Test CalculationEngine cannot allocate a range-proxy owner ID"
                                .to_owned(),
                        )
                    })?;
                let owner = tsce::FormulaOwnerDependenciesArchive {
                    formula_owner_uid: tsp::Uuid {
                        lower: RANGE_PROXY_UUID_LOWER,
                        upper: RANGE_PROXY_UUID_UPPER,
                    },
                    internal_formula_owner_id: internal_owner_id,
                    owner_kind: Some(RANGE_PROXY_OWNER_KIND),
                    range_dependencies: Some(tsce::RangeDependenciesArchive {
                        back_dependency: vec![tsce::RangeBackDependencyArchive {
                            cell_coord_row: 0,
                            cell_coord_column: 0,
                            internal_range_reference: Some(tsce::InternalRangeReferenceArchive {
                                owner_id: source_owner_id,
                                range: tsce::RangeCoordinateArchive {
                                    top_left_column: bounds.left,
                                    top_left_row: bounds.top,
                                    bottom_right_column: bounds.right,
                                    bottom_right_row: bounds.bottom,
                                },
                            }),
                            ..Default::default()
                        }],
                    }),
                    tiled_range_dependencies: Some(tsce::RangeDependenciesTiledArchive {
                        range_precedents_tile: vec![tsp::Reference {
                            identifier: range_tile_id,
                            ..Default::default()
                        }],
                    }),
                    ..Default::default()
                };
                let range_tile = tsce::RangePrecedentsTileArchive {
                    to_owner_id: source_owner_id,
                    from_to_range: vec![tsce::range_precedents_tile_archive::FromToRangeArchive {
                        from_coord: tsce::CellCoordinateArchive {
                            column: Some(0),
                            row: Some(0),
                            ..Default::default()
                        },
                        refers_to_rect: tsce::CellRectArchive {
                            origin: tsce::CellCoordinateArchive {
                                column: Some(bounds.left),
                                row: Some(bounds.top),
                                ..Default::default()
                            },
                            size: tsce::ColumnRowSize {
                                num_columns: (bounds.right != bounds.left)
                                    .then_some(bounds.right - bounds.left + 1),
                                num_rows: (bounds.bottom != bounds.top)
                                    .then_some(bounds.bottom - bounds.top + 1),
                            },
                        },
                    }],
                };
                archive.insert_object(ArchiveObject::new(
                    first_object_id,
                    vec![RawMessage {
                        type_: 4_008,
                        data: owner.encode_to_vec(),
                    }],
                )?)?;
                archive.insert_object(ArchiveObject::new(
                    range_tile_id,
                    vec![RawMessage {
                        type_: 4_010,
                        data: range_tile.encode_to_vec(),
                    }],
                )?)?;
                proxy_ids = Some(NativeMergeRangeProxyIds {
                    object_id: first_object_id,
                    range_tile_id,
                    internal_owner_id,
                });
                Ok(())
            })
            .unwrap();
        proxy_ids.unwrap()
    }

    fn uuid_host_coordinate_set(row: u32, column: u32) -> tsce::CellCoordSetArchive {
        tsce::CellCoordSetArchive {
            column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
                column,
                row_set: tsce::IndexSetArchive {
                    entries: vec![tsce::index_set_archive::IndexSetEntry {
                        range_begin: i32::try_from(row).unwrap(),
                        range_end: None,
                    }],
                },
            }],
        }
    }

    fn formula_owner_at_host(
        editor: &NumbersEditor,
        row: u32,
        column: u32,
    ) -> tsce::FormulaOwnerDependenciesArchive {
        let component = editor
            .package()
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap();
        let archive = editor.package().archive(component).unwrap();
        archive
            .objects
            .iter()
            .flat_map(|object| &object.messages)
            .filter(|message| message.type_ == 4_008)
            .filter_map(|message| {
                tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()
            })
            .find(|owner| {
                owner
                    .cell_dependencies
                    .as_ref()
                    .is_some_and(|dependencies| {
                        dependencies
                            .cell_record
                            .iter()
                            .any(|record| record.row == row && record.column == column)
                    })
            })
            .unwrap()
    }

    fn assert_formula_host_dependencies(
        editor: &NumbersEditor,
        external_owner_id: u32,
        target_row: u32,
        target_column: u32,
    ) {
        let owner = formula_owner_at_host(editor, 1, 1);
        let record = &owner.cell_dependencies.as_ref().unwrap().cell_record[0];
        assert_eq!((record.row, record.column), (1, 1));
        let edges = record.expanded_edges.as_ref().unwrap();
        assert_eq!(edges.internal_owner_id_for_edge, [external_owner_id]);
        assert_eq!(edges.edge_with_owner_rows, [target_row]);
        assert_eq!(edges.edge_with_owner_columns, [target_column]);
        let references = owner.uuid_references.as_ref().unwrap();
        assert_uuid_host_coordinate(references.table_refs[0].coord_set.as_ref().unwrap(), 1, 1);
        assert_uuid_host_coordinate(
            references.table_uuid_refs[0].uuid_refs[0]
                .coord_set
                .as_ref()
                .unwrap(),
            1,
            1,
        );
    }

    fn assert_formula_host_range_edges(
        editor: &NumbersEditor,
        external_owner_id: u32,
        target_rows: &[u32],
        target_columns: &[u32],
    ) {
        assert_eq!(target_rows.len(), target_columns.len());
        let owner = formula_owner_at_host(editor, 1, 1);
        let record = &owner.cell_dependencies.as_ref().unwrap().cell_record[0];
        assert_eq!((record.row, record.column), (1, 1));
        let edges = record.expanded_edges.as_ref().unwrap();
        assert_eq!(
            edges.internal_owner_id_for_edge,
            vec![external_owner_id; target_rows.len()]
        );
        assert_eq!(edges.edge_with_owner_rows, target_rows);
        assert_eq!(edges.edge_with_owner_columns, target_columns);
        let references = owner.uuid_references.as_ref().unwrap();
        assert_uuid_host_coordinate(references.table_refs[0].coord_set.as_ref().unwrap(), 1, 1);
        assert_uuid_host_coordinate(
            references.table_uuid_refs[0].uuid_refs[0]
                .coord_set
                .as_ref()
                .unwrap(),
            1,
            1,
        );
    }

    fn assert_formula_host_native_range_dependencies(
        editor: &NumbersEditor,
        external_owner_id: u32,
        bounds: TestRangeBounds,
    ) {
        let owner = formula_owner_at_host(editor, 1, 1);
        let record = &owner.cell_dependencies.as_ref().unwrap().cell_record[0];
        assert_eq!((record.row, record.column), (1, 1));
        assert_eq!(
            record
                .expanded_edges
                .as_ref()
                .unwrap()
                .internal_owner_id_for_edge,
            []
        );
        let dependency = &owner.range_dependencies.as_ref().unwrap().back_dependency[0];
        assert_eq!(
            (dependency.cell_coord_row, dependency.cell_coord_column),
            (1, 1)
        );
        let reference = dependency.internal_range_reference.as_ref().unwrap();
        assert_eq!(reference.owner_id, external_owner_id);
        assert_eq!(
            (
                reference.range.top_left_row,
                reference.range.top_left_column,
                reference.range.bottom_right_row,
                reference.range.bottom_right_column,
            ),
            (bounds.top, bounds.left, bounds.bottom, bounds.right)
        );
        let component = editor
            .package()
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap();
        let archive = editor.package().archive(component).unwrap();
        let tile_id = owner
            .tiled_range_dependencies
            .as_ref()
            .unwrap()
            .range_precedents_tile[0]
            .identifier;
        let tile = archive.object(tile_id).unwrap();
        let tile = tsce::RangePrecedentsTileArchive::decode(
            tile.messages
                .iter()
                .find(|message| message.type_ == 4_010)
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert_eq!(tile.to_owner_id, external_owner_id);
        let range = &tile.from_to_range[0];
        assert_eq!(
            (range.from_coord.row, range.from_coord.column),
            (Some(1), Some(1))
        );
        assert_eq!(
            (
                range.refers_to_rect.origin.row,
                range.refers_to_rect.origin.column
            ),
            (Some(bounds.top), Some(bounds.left))
        );
        assert_eq!(
            range.refers_to_rect.size.num_rows,
            (bounds.bottom != bounds.top).then_some(bounds.bottom - bounds.top + 1)
        );
        assert_eq!(
            range.refers_to_rect.size.num_columns,
            (bounds.right != bounds.left).then_some(bounds.right - bounds.left + 1)
        );
        let references = owner.uuid_references.as_ref().unwrap();
        assert_uuid_host_coordinate(references.table_refs[0].coord_set.as_ref().unwrap(), 1, 1);
        assert_uuid_host_coordinate(
            references.table_uuid_refs[0].uuid_refs[0]
                .coord_set
                .as_ref()
                .unwrap(),
            1,
            1,
        );
    }

    fn assert_native_merge_range_proxy(
        editor: &NumbersEditor,
        proxy: NativeMergeRangeProxyIds,
        source_owner_id: u32,
        bounds: TestRangeBounds,
    ) {
        let component = editor
            .package()
            .calculation_engine_entry_name()
            .unwrap()
            .unwrap();
        let archive = editor.package().archive(component).unwrap();
        let object = archive.object(proxy.object_id).unwrap();
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(
            object
                .messages
                .iter()
                .find(|message| message.type_ == 4_008)
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert_eq!(owner.internal_formula_owner_id, proxy.internal_owner_id);
        assert_eq!(owner.formula_owner, None);
        let dependency = &owner.range_dependencies.as_ref().unwrap().back_dependency[0];
        let reference = dependency.internal_range_reference.as_ref().unwrap();
        assert_eq!(reference.owner_id, source_owner_id);
        assert_eq!(
            (
                reference.range.top_left_row,
                reference.range.top_left_column,
                reference.range.bottom_right_row,
                reference.range.bottom_right_column,
            ),
            (bounds.top, bounds.left, bounds.bottom, bounds.right)
        );
        let tile = archive.object(proxy.range_tile_id).unwrap();
        let tile = tsce::RangePrecedentsTileArchive::decode(
            tile.messages
                .iter()
                .find(|message| message.type_ == 4_010)
                .unwrap()
                .data
                .as_slice(),
        )
        .unwrap();
        assert_eq!(tile.to_owner_id, source_owner_id);
        let range = &tile.from_to_range[0].refers_to_rect;
        assert_eq!(
            (range.origin.row, range.origin.column),
            (Some(bounds.top), Some(bounds.left))
        );
        assert_eq!(
            range.size.num_rows,
            (bounds.bottom != bounds.top).then_some(bounds.bottom - bounds.top + 1)
        );
        assert_eq!(
            range.size.num_columns,
            (bounds.right != bounds.left).then_some(bounds.right - bounds.left + 1)
        );
    }

    fn assert_uuid_host_coordinate(coordinates: &tsce::CellCoordSetArchive, row: i32, column: u32) {
        assert_eq!(coordinates.column_entries.len(), 1);
        let entry = &coordinates.column_entries[0];
        assert_eq!(entry.column, column);
        assert_eq!(entry.row_set.entries.len(), 1);
        assert_eq!(entry.row_set.entries[0].range_begin, row);
        assert_eq!(entry.row_set.entries[0].range_end, None);
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
            Some(crate::numbers::bnc::CachedScalar::Number(value)) => value.get(),
            value => panic!("Expected numeric formula cache, found {value:?}"),
        }
    }

    fn unknown_suffix(field: u32, value: u64) -> Vec<u8> {
        let mut suffix = Vec::new();
        append_unknown_varint(&mut suffix, field, value);
        suffix
    }
}
