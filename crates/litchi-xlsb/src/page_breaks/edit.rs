//! Immutable page-break edits of an XLSB [`Package`] (unified-ops record
//! 0322).
//!
//! [`Package::edit_page_breaks`] starts
//! an [`Edit`]; [`Edit::set_page_breaks`] replaces, inserts or removes a
//! worksheet's `BrtBeginRwBrk` and `BrtBeginColBrk` collections, and every
//! other worksheet record keeps its bytes. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_page_breaks`] replays a patch onto an
//! equal package. An exact no-op keeps every byte, and a changed edit of a
//! signed workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::page_breaks::{Break, PageBreaks};
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_page_breaks()?;
//! let breaks = PageBreaks { rows: vec![Break::row(9)], columns: vec![] };
//! assert!(edit.set_page_breaks(0, &breaks)?);
//! let (after, patch) = edit.commit()?.into_parts();
//! assert_eq!(after.edit_page_breaks()?.page_breaks(0)?, breaks);
//! assert_eq!(package.apply_page_breaks(&patch)?.package().to_bytes()?, after.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use super::{PageBreaks, encode, read_layout};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Rank, SheetLayout, Staged, family, snapshot_options, splice};
use crate::raw::kind;

family!("page-break");
snapshot_options!("page-break");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Breaks of one direction on one worksheet ([`super::MAX_BREAKS`]).
    Breaks,
}

/// A staged page-break edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start a page-break edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_page_breaks(&self) -> Result<Edit> {
        self.edit_page_breaks_with_options(EditOptions::default())
    }

    /// Start a page-break edit with an explicit retained-payload ceiling.
    pub fn edit_page_breaks_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a page-break patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_page_breaks(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged page breaks of the worksheet at complete-catalog position
    /// `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing tab, a non-worksheet tab or a
    /// malformed stream.
    pub fn page_breaks(&self, sheet: usize) -> Result<PageBreaks> {
        let worksheet = self.staged.worksheet(sheet)?;
        let bytes = self.staged.blob(&worksheet)?;
        let limits = self.staged.raw_limits();
        read_layout(
            bytes,
            &SheetLayout::scan_with_limits(bytes, limits)?,
            limits,
        )
    }

    /// Replace the page breaks of a worksheet. A direction whose breaks are
    /// unchanged keeps its bytes; an empty direction removes its
    /// collection. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a break outside the grid, a duplicate index,
    /// more than [`super::MAX_BREAKS`] breaks of one direction, an automatic
    /// break, or a changed edit of a signed workbook.
    pub fn set_page_breaks(&mut self, sheet: usize, breaks: &PageBreaks) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let mut bytes = self.staged.blob(&worksheet)?.to_vec();
        let limits = self.staged.raw_limits();
        let current = read_layout(
            &bytes,
            &SheetLayout::scan_with_limits(&bytes, limits)?,
            limits,
        )?;
        for (row, wanted, held) in [
            (true, &breaks.rows, &current.rows),
            (false, &breaks.columns, &current.columns),
        ] {
            if wanted == held {
                continue;
            }
            let encoded = encode(wanted, row, limits)?;
            let layout = SheetLayout::scan_with_limits(&bytes, limits)?;
            let (begin, rank) = if row {
                (kind::BEGIN_RW_BRK, Rank::RowBreaks)
            } else {
                (kind::BEGIN_COL_BRK, Rank::ColumnBreaks)
            };
            let span = match layout.units_of(begin).next() {
                Some(unit) => unit.span.clone(),
                None => {
                    let at = layout.insertion(rank);
                    at..at
                },
            };
            bytes = splice(&bytes, span, &encoded);
        }
        self.staged.replace_worksheet(&worksheet, bytes)
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

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::page_breaks::{Break, PageBreaks};
    use crate::raw::Limits;

    #[test]
    fn fixed_break_payload_obeys_the_required_raw_limit() {
        let error = super::encode(&[Break::row(4)], true, Limits::new(19, 64)).unwrap_err();
        assert!(matches!(
            error,
            crate::package::error::Error::Wire(crate::raw::Error::PayloadLimit {
                length: 20,
                limit: 19,
                ..
            })
        ));
    }

    #[test]
    fn set_replace_remove_and_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_page_breaks().unwrap();
        let breaks = PageBreaks {
            rows: vec![Break::row(4), Break::row(9)],
            columns: vec![Break::column(3)],
        };
        assert!(edit.set_page_breaks(0, &breaks).unwrap());
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(
            after.edit_page_breaks().unwrap().page_breaks(0).unwrap(),
            breaks
        );
        assert_eq!(
            package
                .apply_page_breaks(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        // An exact no-op keeps the snapshot.
        let mut edit = after.edit_page_breaks().unwrap();
        assert!(!edit.set_page_breaks(0, &breaks).unwrap());
        assert!(edit.commit().unwrap().patch().is_noop());
        // Removing every break restores the original stream.
        let mut edit = after.edit_page_breaks().unwrap();
        assert!(edit.set_page_breaks(0, &PageBreaks::default()).unwrap());
        let restored = edit.commit().unwrap().into_package();
        assert_eq!(restored.to_bytes().unwrap(), package.to_bytes().unwrap());
        // The inverse patch undoes the edit.
        assert_eq!(
            after
                .apply_page_breaks(&patch.inverse())
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            package.to_bytes().unwrap()
        );
    }

    #[test]
    fn refuses_invalid_breaks() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_page_breaks().unwrap();
        let mut automatic = Break::row(2);
        automatic.manual = false;
        for rows in [
            vec![Break::row(2), Break::row(2)],
            vec![automatic],
            vec![Break::row(2_000_000)],
        ] {
            let breaks = PageBreaks {
                rows,
                columns: vec![],
            };
            assert!(edit.set_page_breaks(0, &breaks).is_err());
        }
    }
}
