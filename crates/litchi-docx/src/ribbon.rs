//! The package-level Ribbon (`customUI`) family (record 0355).
//!
//! [`Package::ribbon`](crate::Package::ribbon) reads the legacy and modern
//! Ribbon slots. [`Package::edit_ribbon`](crate::Package::edit_ribbon) stages
//! an [`Edit`]; [`Edit::commit`] publishes a new package and leaves the
//! source untouched, and [`Package::apply_ribbon`](crate::Package::apply_ribbon)
//! replays the [`Patch`] onto another package with the same before state.
//! The XML stays inert: callbacks and commands are never invoked.
//!
//! ```
//! use litchi_docx::{Package, ribbon};
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_ribbon();
//! edit.put(
//!     ribbon::Version::V2010,
//!     br#"<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"/>"#.to_vec(),
//! )?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! assert!(commit.package().ribbon()?.modern().is_some());
//! // The source is unchanged, and the inverse patch restores it.
//! assert!(package.ribbon()?.modern().is_none());
//! let undone = commit.package().apply_ribbon(&commit.patch().inverse())?;
//! assert!(undone.package().ribbon()?.modern().is_none());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

pub use litchi_ooxml_common::ribbon::{Family, Set, Ui, Version};

use crate::package::family::{graph_edit, graph_family};

graph_family!("Ribbon", own_limits);

/// The typed budgets of the Ribbon family: its limit errors are
/// [`litchi_ooxml_common::Error::Limit`] carried by
/// [`crate::Error::Common`] (record 0360).
pub use litchi_ooxml_common::ribbon::LimitKind;
graph_edit!("Ribbon", edit_ribbon, apply_ribbon);

impl Edit {
    /// Store opaque Ribbon XML for `version`, replacing the slot's current
    /// part. The version and bytes the slot already stores stage nothing.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for XML the Ribbon owner rejects, and the
    /// signed-package refusal for a change to a signed package (ADR 0006).
    pub fn put(&mut self, version: Version, xml: Vec<u8>) -> Result<&mut Self> {
        self.stage(|package| package.put_ribbon(version, xml).map(|_| ()))?;
        Ok(self)
    }

    /// Remove one Ribbon relationship family and its orphaned part; returns
    /// whether the family was present.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn remove(&mut self, family: Family) -> Result<bool> {
        self.stage(|package| package.remove_ribbon(family))
    }
}
