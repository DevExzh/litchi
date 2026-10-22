//! Immutable merged-cell edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_merged_cells`] starts an [`Edit`]; its verbs replace,
//! add, remove or clear a worksheet's `BrtBeginMergeCells` collection, and
//! every other worksheet record keeps its bytes (the binary index stays
//! current). [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_merged_cells`] replays a patch onto an equal package. An
//! exact no-op keeps every byte, and a changed edit of a signed workbook is
//! refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::merged_cells::MergedCell;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_merged_cells()?;
//! let range = MergedCell::new(0, 1, 0, 1);
//! assert!(edit.add_merged_cell_range(0, range.clone())?);
//! let (after, patch) = edit.commit()?.into_parts();
//! assert_eq!(after.workbook()?.merged_cell_ranges(0)?, vec![range]);
//! assert_eq!(package.apply_merged_cells(&patch)?.package().to_bytes()?, after.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use super::MergedCell;
use crate::Workbook;
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("merged-cell");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Merged ranges of one worksheet ([`super::MAX_MERGED_CELL_RANGES`]).
    Ranges,
}

/// A staged merged-cell edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a merged-cell edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_merged_cells(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
        })
    }

    /// Apply a merged-cell patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_merged_cells(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged merged ranges of the worksheet at complete-catalog position
    /// `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab or a malformed
    /// stream.
    pub fn merged_cell_ranges(&self, sheet: usize) -> Result<Vec<MergedCell>> {
        let worksheet = self.staged.worksheet(sheet)?;
        Ok(
            Workbook::inspect_merge_block(self.staged.blob(&worksheet)?, self.staged.raw_limits())?
                .ranges,
        )
    }

    /// Replace every merged range of a worksheet; an empty list removes the
    /// collection. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab, an invalid,
    /// duplicate or overlapping range, more than
    /// [`super::MAX_MERGED_CELL_RANGES`] ranges, and a changed edit of a
    /// signed workbook.
    pub fn set_merged_cell_ranges(&mut self, sheet: usize, ranges: &[MergedCell]) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let original = self.staged.blob(&worksheet)?;
        let raw_limits = self.staged.raw_limits();
        let layout = Workbook::inspect_merge_block(original, raw_limits)?;
        let normalized = Workbook::normalize_merge_ranges(ranges)?;
        let replacement = Workbook::serialize_merge_block(&normalized, raw_limits)?;
        let (start, end) = layout
            .block_span
            .unwrap_or((layout.insertion_offset, layout.insertion_offset));
        let mut updated = Vec::new();
        updated
            .try_reserve_exact(original.len() - (end - start) + replacement.len())
            .map_err(|source| crate::package::PackageError::Allocation {
                resource: "XLSB merged-cell worksheet",
                source,
            })?;
        updated.extend_from_slice(&original[..start]);
        updated.extend_from_slice(&replacement);
        updated.extend_from_slice(&original[end..]);
        Workbook::inspect_merge_block(&updated, raw_limits)?;
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Add one merged range to a worksheet. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_merged_cell_ranges`].
    pub fn add_merged_cell_range(&mut self, sheet: usize, range: MergedCell) -> Result<bool> {
        let mut ranges = self.merged_cell_ranges(sheet)?;
        ranges.push(range);
        self.set_merged_cell_ranges(sheet, &ranges)
    }

    /// Remove an exact merged range from a worksheet. Returns whether it was
    /// present.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_merged_cell_ranges`].
    pub fn remove_merged_cell_range(&mut self, sheet: usize, range: &MergedCell) -> Result<bool> {
        let mut ranges = self.merged_cell_ranges(sheet)?;
        let Some(index) = ranges.iter().position(|candidate| candidate == range) else {
            return Ok(false);
        };
        ranges.remove(index);
        self.set_merged_cell_ranges(sheet, &ranges)?;
        Ok(true)
    }

    /// Remove every merged range of a worksheet. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::set_merged_cell_ranges`].
    pub fn clear_merged_cell_ranges(&mut self, sheet: usize) -> Result<bool> {
        self.set_merged_cell_ranges(sheet, &[])
    }

    /// Validate the staged worksheets and publish them to a new package. An
    /// unchanged staging returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed staging of a signed workbook
    /// and the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
