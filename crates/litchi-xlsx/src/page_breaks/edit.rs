//! Immutable, source-bound worksheet page-break edits of a
//! [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_page_breaks`] starts an [`Edit`] of one
//! worksheet's row and column breaks; [`Edit::commit`] returns a [`Commit`]
//! holding the new package, the exact [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_page_breaks`] replays a patch onto an equal
//! package. The source package never changes; an exact no-op returns it, a
//! signed one included, and a changed edit of a signed workbook is refused
//! with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::page_breaks::{Break, Collection};
//!
//! # fn main() -> litchi_xlsx::Result<()> {
//! let package = Package::new()?;
//! let mut edit = package.edit_page_breaks("Sheet1")?;
//! edit.set_horizontal(Collection::horizontal([Break::new(7, 0, 0)?])?)?;
//! let (after, patch) = edit.commit()?.into_parts();
//! assert!(after.page_breaks("Sheet1")?.page_breaks().horizontal().is_some());
//! let restored = after.apply_page_breaks(&patch.inverse())?.into_package();
//! assert!(restored.page_breaks("Sheet1")?.page_breaks().horizontal().is_none());
//! # Ok(())
//! # }
//! ```

use litchi_opc::OpcPackage;

use super::{Collection, PageBreaks, Snapshot};
use crate::Selector;
use crate::error::{Result, invalid};

pub use super::LimitKind;

crate::record::record_family!("page-break");

/// The exact, reversible change of a page-break edit: the worksheet part's
/// before and after states.
#[derive(Clone, Debug)]
pub struct Patch(pub(crate) super::Patch);

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.0.is_empty()
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// The worksheet state the patch requires before application.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The worksheet state the patch produces.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        self.0.after()
    }

    /// The semantic payload retained by this reversible patch.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        let before = self
            .0
            .before()
            .source_arc()
            .map_or(0, |bytes| bytes.capacity());
        let after = self
            .0
            .after()
            .source_arc()
            .map_or(0, |bytes| bytes.capacity());
        before.saturating_add(after)
    }
}

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.0.apply(package)
    }
}

/// A staged page-break edit of one worksheet of a [`crate::Package`].
/// Dropping it discards it.
pub struct Edit {
    target: crate::record::Source,
    before: Snapshot,
    staged: PageBreaks,
    max_retained_bytes: Option<usize>,
}

impl crate::Package {
    /// Start a page-break edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and an error when the selector does not select a worksheet
    /// or its page-break XML is invalid.
    pub fn edit_page_breaks<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.edit_page_breaks_with_options(selector, crate::EditOptions::default())
    }

    /// Start a page-break edit with an explicit retained-patch budget.
    pub fn edit_page_breaks_with_options<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
        options: crate::EditOptions,
    ) -> Result<Edit> {
        let target = crate::record::Source::new(self, "edit_page_breaks")?;
        let before = Snapshot::load(&target, selector)?;
        let staged = before.page_breaks().clone();
        Ok(Edit {
            target,
            before,
            staged,
            max_retained_bytes: options.max_retained_bytes_value(),
        })
    }

    /// Apply an exact page-break patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_page_breaks(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_page_breaks")
    }
}

impl Edit {
    /// Exact source state captured at transaction start.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Currently staged page-break state.
    #[must_use]
    pub const fn page_breaks(&self) -> &PageBreaks {
        &self.staged
    }

    /// Replace the complete staged page-break state.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` violates page-break invariants.
    pub fn set(&mut self, value: PageBreaks) -> Result<bool> {
        value.validate()?;
        if self.staged == value {
            return Ok(false);
        }
        self.staged = value;
        Ok(true)
    }

    /// Clone-edit the complete staged value; a failed closure changes nothing.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or a page-break validation error.
    pub fn edit(&mut self, edit: impl FnOnce(&mut PageBreaks) -> Result<()>) -> Result<bool> {
        let mut draft = self.staged.clone();
        edit(&mut draft)?;
        self.set(draft)
    }

    /// Replace horizontal row breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not horizontal or is invalid.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_horizontal(collection))
    }

    /// Replace vertical column breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not vertical or is invalid.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_vertical(collection))
    }

    /// Remove horizontal row breaks.
    pub fn remove_horizontal(&mut self) -> bool {
        self.staged.remove_horizontal()
    }

    /// Remove vertical column breaks.
    pub fn remove_vertical(&mut self) -> bool {
        self.staged.remove_vertical()
    }

    /// Remove both authored collections.
    pub fn clear(&mut self) -> bool {
        self.staged.clear()
    }

    /// Whether the exact authored semantic state differs from its source.
    fn changed(&self) -> bool {
        self.before.page_breaks() != &self.staged
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged = self.before.page_breaks().clone();
    }

    /// Validate, rewrite the worksheet, read it back and publish it as a new
    /// package. An unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes, signed packages, invalid
    /// staged state, rewrite failures, or semantic readback failures.
    pub fn commit(self) -> Result<Commit> {
        if !self.changed() {
            let patch = Patch(super::Patch::new(self.before.clone(), self.before));
            if let Some(maximum) = self.max_retained_bytes {
                let observed = self
                    .target
                    .retained_bytes()
                    .saturating_add(self.target.retained_bytes())
                    .saturating_add(patch.retained_bytes());
                if observed > maximum {
                    return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                        resource: litchi_core::Resource::Memory,
                        observed: u64::try_from(observed).unwrap_or(u64::MAX),
                        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                        scope: "XLSX page-break retained package".into(),
                    }));
                }
            }
            return Ok(Commit::new(
                self.target.into_package(),
                patch,
                Diagnostics::new(false),
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        if !self.before.matches_current_source(&self.target) {
            return Err(Error::PatchConflict {
                part: self.before.worksheet_part_name().to_string(),
            });
        }
        let output = super::replace(self.before.source_xml(), &self.staged)?;
        let estimated_patch_bytes = self
            .before
            .source_arc()
            .map(|before| before.capacity())
            .unwrap_or(0)
            .saturating_add(output.capacity());
        let mut candidate = OpcPackage::clone(&self.target);
        candidate
            .get_part_mut(self.before.worksheet_part_name())?
            .set_blob(output);
        let snapshot = Snapshot::load(&candidate, self.before.sheet_position())?;
        if snapshot.page_breaks() != &self.staged {
            return Err(invalid(
                "page-break publication changed the staged semantic state",
            ));
        }
        let package = self.target.publish(candidate)?;
        if let Some(maximum) = self.max_retained_bytes {
            let observed = self
                .target
                .retained_bytes()
                .saturating_add(package.retained_bytes())
                .saturating_add(estimated_patch_bytes);
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX page-break retained package".into(),
                }));
            }
        }
        let patch = Patch(super::Patch::new(self.before, snapshot));
        if let Some(maximum) = self.max_retained_bytes {
            let observed = self
                .target
                .retained_bytes()
                .saturating_add(package.retained_bytes())
                .saturating_add(patch.retained_bytes());
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX page-break retained package".into(),
                }));
            }
        }
        Ok(Commit::new(package, patch, Diagnostics::new(true)))
    }
}
