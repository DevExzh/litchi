//! Immutable auto-filter edits of an XLSB [`Package`] (unified-ops record
//! 0322).
//!
//! [`Package::edit_auto_filter`] starts
//! an [`Edit`]. [`Edit::set_auto_filter`] replaces the filter range in place
//! (the `BrtBeginAFilter` payload; filter columns that still fall inside the
//! range keep their bytes), inserts a new collection at its stream position,
//! or removes the collection. Every other worksheet record keeps its bytes.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_auto_filter`] replays a patch onto an equal package.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::auto_filter::AutoFilter;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_auto_filter()?;
//! let range = AutoFilter { first_row: 0, last_row: 9, first_column: 0, last_column: 2 };
//! assert!(edit.set_auto_filter(0, Some(range))?);
//! let after = edit.commit()?.into_package();
//! assert_eq!(after.edit_auto_filter()?.auto_filter(0)?, Some(range));
//! # Ok(())
//! # }
//! ```

use super::{AutoFilter, encode_range, read_layout};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{
    Rank, SheetLayout, Staged, family, record_with_limits, snapshot_options, splice,
};
use crate::raw::{Records, kind};

family!("auto-filter");
snapshot_options!("auto-filter");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Auto-filters of one worksheet (one).
    Filters,
}

/// A staged auto-filter edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start an auto-filter edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_auto_filter(&self) -> Result<Edit> {
        self.edit_auto_filter_with_options(EditOptions::default())
    }

    /// Start an auto-filter edit with an explicit retained-payload ceiling.
    pub fn edit_auto_filter_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply an auto-filter patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_auto_filter(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged auto-filter range of the worksheet at complete-catalog
    /// position `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// stream.
    pub fn auto_filter(&self, sheet: usize) -> Result<Option<AutoFilter>> {
        let worksheet = self.staged.worksheet(sheet)?;
        let bytes = self.staged.blob(&worksheet)?;
        let limits = self.staged.raw_limits();
        read_layout(
            bytes,
            &SheetLayout::scan_with_limits(bytes, limits)?,
            limits,
        )
    }

    /// Replace (`Some`, keeping the filter columns still inside the range),
    /// insert or remove (`None`) the auto-filter. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a range outside the grid, a malformed stream or
    /// a changed edit of a signed workbook.
    pub fn set_auto_filter(&mut self, sheet: usize, value: Option<AutoFilter>) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let bytes = self.staged.blob(&worksheet)?;
        let limits = self.staged.raw_limits();
        let layout = SheetLayout::scan_with_limits(bytes, limits)?;
        let unit = layout.units_of(kind::BEGIN_A_FILTER).next().cloned();
        let current = read_layout(bytes, &layout, limits)?;
        if current == value {
            return Ok(false);
        }
        let replacement = match value {
            None => Vec::new(),
            Some(range) => {
                let mut out = record_with_limits(
                    kind::BEGIN_A_FILTER,
                    &encode_range(range, limits)?,
                    limits,
                )?;
                if let Some(unit) = &unit {
                    out.extend(columns_inside(
                        &bytes[unit.span.clone()],
                        range.last_column - range.first_column,
                        limits,
                    )?);
                }
                out.extend(record_with_limits(kind::END_A_FILTER, &[], limits)?);
                out
            },
        };
        let span = match &unit {
            Some(unit) => unit.span.clone(),
            None => {
                let at = layout.insertion(Rank::AutoFilter);
                at..at
            },
        };
        let updated = splice(bytes, span, &replacement);
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

/// The filter-column collections of an auto-filter unit whose column index
/// (`dwCol`) is at most `last`, with their exact bytes.
fn columns_inside(unit: &[u8], last: u32, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    let mut out = Vec::new();
    let mut records = Records::new(unit, limits);
    let _header = records.next().transpose()?;
    let mut keep = false;
    let mut depth = 0usize;
    while let Some(item) = records.next() {
        let item = item?;
        let start = item.offset();
        let end = records.offset();
        if depth == 0 {
            if item.kind() == kind::END_A_FILTER {
                break;
            }
            keep = item.kind() == kind::BEGIN_FILTER_COLUMN
                && item.payload().len() >= 4
                && u32::from_le_bytes([
                    item.payload()[0],
                    item.payload()[1],
                    item.payload()[2],
                    item.payload()[3],
                ]) <= last;
        }
        if item.kind() == kind::BEGIN_FILTER_COLUMN {
            depth += 1;
        } else if item.kind() == kind::END_FILTER_COLUMN {
            depth = depth.saturating_sub(1);
        }
        if keep {
            out.extend_from_slice(&unit[start..end]);
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::auto_filter::AutoFilter;
    use crate::raw::Limits;

    #[test]
    fn fixed_range_payload_obeys_the_required_raw_limit() {
        let range = AutoFilter {
            first_row: 0,
            last_row: 4,
            first_column: 0,
            last_column: 1,
        };
        let error = super::encode_range(range, Limits::new(15, 64)).unwrap_err();
        assert!(matches!(
            error,
            crate::package::error::Error::Wire(crate::raw::Error::PayloadLimit {
                length: 16,
                limit: 15,
                ..
            })
        ));
    }

    #[test]
    fn set_move_remove_replay() {
        let package = Package::new().unwrap();
        let range = AutoFilter {
            first_row: 0,
            last_row: 4,
            first_column: 0,
            last_column: 1,
        };
        let mut edit = package.edit_auto_filter().unwrap();
        assert!(edit.set_auto_filter(0, Some(range)).unwrap());
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(
            after.edit_auto_filter().unwrap().auto_filter(0).unwrap(),
            Some(range)
        );
        assert_eq!(
            package
                .apply_auto_filter(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        let wider = AutoFilter {
            last_column: 3,
            ..range
        };
        let mut edit = after.edit_auto_filter().unwrap();
        assert!(edit.set_auto_filter(0, Some(wider)).unwrap());
        let wide = edit.commit().unwrap().into_package();
        assert_eq!(
            wide.workbook()
                .unwrap()
                .grid(0usize)
                .unwrap()
                .unwrap()
                .worksheet()
                .auto_filter(),
            Some(wider)
        );
        let mut edit = wide.edit_auto_filter().unwrap();
        assert!(edit.set_auto_filter(0, None).unwrap());
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(removed.to_bytes().unwrap(), package.to_bytes().unwrap());
    }
}
