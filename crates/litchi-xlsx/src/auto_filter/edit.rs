//! Immutable auto-filter edits of a [`Package`] (unified-ops record 0320).
//!
//! [`Package::edit_auto_filter`] starts an [`Edit`] over an immutable
//! snapshot. [`Edit::auto_filter`] reads a worksheet's direct auto-filter
//! and [`Edit::set_auto_filter`] replaces, inserts or removes it; every byte
//! outside the `autoFilter` span is copied exactly. [`Edit::commit`] returns
//! a [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_auto_filter`] replays a patch onto an
//! equal package. An exact no-op keeps every byte, and a changed edit of a
//! signed workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::auto_filter::{Definition, Range};
//!
//! # fn main() -> Result<(), litchi_xlsx::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_auto_filter()?;
//! edit.set_auto_filter(0, Some(Definition::new(Some(Range::new("A1:C10")?))))?;
//! let after = edit.commit()?.into_package();
//! assert!(after.workbook().sheet(0)?.unwrap().auto_filter()?.is_some());
//! # Ok(())
//! # }
//! ```

use super::Definition;
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("auto-filter");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Filter columns of one auto-filter (the leaf's definition limit).
    Columns,
}

/// A staged auto-filter edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start an auto-filter edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_auto_filter(&self) -> Result<Edit> {
        self.edit_auto_filter_with_options(crate::EditOptions::default())
    }

    /// Start an auto-filter edit with an explicit retained-patch budget.
    pub fn edit_auto_filter_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_auto_filter", options)?,
        })
    }

    /// Apply an auto-filter patch to a new package; this one is unchanged.
    /// An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_auto_filter(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_auto_filter")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged direct auto-filter of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, or malformed
    /// markup.
    pub fn auto_filter(&self, sheet: usize) -> Result<Option<Definition>> {
        let worksheet = self.staged.worksheet(sheet)?;
        super::parse_auto_filter(self.staged.blob(&worksheet)?)
    }

    /// Replace, insert (`Some`) or remove (`None`) the direct auto-filter of
    /// the worksheet at `sheet`. Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, an invalid
    /// definition, a declaration selected through markup compatibility, or a
    /// changed edit of a signed workbook.
    pub fn set_auto_filter(&mut self, sheet: usize, value: Option<Definition>) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let xml = super::replace_auto_filter(self.staged.blob(&worksheet)?, value.as_ref())?;
        let xml = crate::part_edit::WorksheetLayout::scan(&xml)?.in_root_prefix(&xml, "autoFilter");
        if super::parse_auto_filter(&xml)? != value {
            return Err(crate::error::invalid("the auto-filter did not read back"));
        }
        self.staged.replace_blob(&worksheet, xml)
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::auto_filter::{Definition, Range};

    #[test]
    fn set_remove_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_auto_filter().unwrap();
        let value = Definition::new(Some(Range::new("A1:B5").unwrap()));
        assert!(edit.set_auto_filter(0, Some(value.clone())).unwrap());
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(
            after.edit_auto_filter().unwrap().auto_filter(0).unwrap(),
            Some(value)
        );
        assert!(
            package
                .apply_auto_filter(&patch)
                .unwrap()
                .package()
                .same_content(&after)
                .unwrap()
        );
        let mut edit = after.edit_auto_filter().unwrap();
        assert!(edit.set_auto_filter(0, None).unwrap());
        let removed = edit.commit().unwrap().into_package();
        assert!(removed.same_content(&package).unwrap());
    }

    #[test]
    fn retained_patch_budget_rejects_one_below_and_accepts_exact() {
        let package = Package::new().unwrap();
        let value = Definition::new(Some(Range::new("A1:B5").unwrap()));

        let mut probe = package.edit_auto_filter().unwrap();
        probe.set_auto_filter(0, Some(value.clone())).unwrap();
        let (after, patch) = probe.commit().unwrap().into_parts();
        let required = package
            .retained_bytes()
            .saturating_add(after.retained_bytes())
            .saturating_add(patch.retained_bytes());
        assert!(required > 0);

        let mut rejected = package
            .edit_auto_filter_with_options(
                crate::EditOptions::default().with_max_retained_bytes(required - 1),
            )
            .unwrap();
        rejected.set_auto_filter(0, Some(value.clone())).unwrap();
        assert!(matches!(
            rejected.commit(),
            Err(crate::Error::ResourceLimit(limit)) if limit.limit == (required - 1) as u64
        ));

        let mut exact = package
            .edit_auto_filter_with_options(
                crate::EditOptions::default().with_max_retained_bytes(required),
            )
            .unwrap();
        exact.set_auto_filter(0, Some(value)).unwrap();
        let commit = exact.commit().unwrap();
        assert_eq!(
            package
                .retained_bytes()
                .saturating_add(commit.package().retained_bytes())
                .saturating_add(commit.patch().retained_bytes()),
            required
        );
    }
}
