//! Immutable Ribbon customization edits of an XLSB [`Package`] (record
//! 0360b).
//!
//! [`Package::edit_ribbon`] starts an [`Edit`] of the bounded, inert
//! package-level Ribbon customizations; [`Edit::put`] creates or replaces
//! one family and [`Edit::remove`] removes one family and its unreferenced
//! part. [`Edit::commit`] returns a [`Commit`] with the new package, the
//! exact part-level [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_ribbon`] replays a patch onto an equal package. The
//! version and bytes a family already stores publish nothing, so a signed
//! workbook keeps its signature byte for byte; any other change of a signed
//! workbook is refused with [`Error::Signed`].

use litchi_ooxml_common::ribbon::{self as common_ribbon, Family, Version};

use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{Staged, family};

family!("Ribbon");

pub use crate::opc::LimitKind;

/// A staged Ribbon edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
}

/// The Ribbon customization a package stores in `family`.
fn slot(set: common_ribbon::Set<'_>, family: Family) -> Option<common_ribbon::Ui<'_>> {
    match family {
        Family::Legacy => set.legacy(),
        Family::Modern => set.modern(),
    }
}

impl Package {
    /// The bounded, inert package-level Ribbon customizations.
    ///
    /// # Errors
    ///
    /// Returns the Ribbon graph and decoding errors.
    pub fn ribbon(&self) -> Result<common_ribbon::Set<'_>> {
        Ok(common_ribbon::load(
            self.opc_package(),
            &common_ribbon::Limits::standard(),
        )?)
    }

    /// Start a Ribbon edit. Starting one copies nothing.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of a package whose projection
    /// was not parsed yet.
    pub fn edit_ribbon(&self) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::unparsed(self)?,
        })
    }

    /// Apply a Ribbon patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_ribbon(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply_unparsed(self)?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged version and XML of one Ribbon family, if present.
    ///
    /// # Errors
    ///
    /// Returns the Ribbon graph and decoding errors.
    pub fn get(&self, family: Family) -> Result<Option<(Version, Vec<u8>)>> {
        let set = common_ribbon::load(self.staged.opc(), &common_ribbon::Limits::standard())?;
        Ok(slot(set, family).map(|ui| (ui.version(), ui.xml().to_vec())))
    }

    /// Create or replace one Ribbon customization family. Returns whether
    /// anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for XML that violates the Ribbon constraints or
    /// exceeds a bound, and a changed edit of a signed workbook.
    pub fn put(&mut self, version: Version, xml: Vec<u8>) -> Result<bool> {
        let set = common_ribbon::load(self.staged.opc(), &common_ribbon::Limits::standard())?;
        if slot(set, version.family()).is_some_and(|ui| ui.version() == version && ui.xml() == xml)
        {
            return Ok(false);
        }
        self.staged.update(|candidate| {
            let mut edit = common_ribbon::Edit::new(candidate)?;
            let changed = edit.put(version, xml)?;
            *candidate = edit.commit()?.into_package();
            Ok(changed)
        })
    }

    /// Remove one Ribbon relationship family and its unreferenced part.
    /// Returns whether the family was present.
    ///
    /// # Errors
    ///
    /// Returns the Ribbon graph errors and a changed edit of a signed
    /// workbook.
    pub fn remove(&mut self, family: Family) -> Result<bool> {
        let set = common_ribbon::load(self.staged.opc(), &common_ribbon::Limits::standard())?;
        if slot(set, family).is_none() {
            return Ok(false);
        }
        self.staged.update(|candidate| {
            let mut edit = common_ribbon::Edit::new(candidate)?;
            let changed = edit.remove(family)?;
            *candidate = edit.commit()?.into_package();
            Ok(changed)
        })
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged Ribbon graph to a new package. An unchanged edit
    /// returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}
