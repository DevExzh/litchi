//! Typed native table-cell merge storage shared by the iWork suite.

use std::num::NonZeroUsize;

use prost::Message;

use super::*;
use crate::numbers::formula_owner::{formula_owner_uuid_for_table, uuid_as_cfuuid};

mod formula;
mod wire;

#[cfg(test)]
use formula::parse_merge_formula;
use formula::{merge_formula, parse_regions, parse_table_uuid, validate_region_bounds};
use wire::{
    add_formula_store, add_merge_owner, append_formula, patch_table_model, remove_formula,
    remove_merge_owner, transform_formula_store, transform_merge_owner,
};

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
    use crate::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
    use crate::numbers::NumbersDocumentBuilder;
    use crate::pages::{PagesDocumentBuilder, PagesEditor};
    use crate::shapes::{DrawablePoint, DrawableSize};

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
}
