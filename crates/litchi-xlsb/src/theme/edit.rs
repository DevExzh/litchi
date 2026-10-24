//! Immutable Theme edits of an XLSB [`Package`] (record 0360b).
//!
//! [`Package::edit_theme`] starts an [`Edit`] of the workbook's optional
//! DrawingML Theme part: it can create the part, replace its typed color and
//! font schemes, its name and its applied-theme family, or remove it. An
//! existing part keeps every byte of markup outside the typed vocabulary; a
//! scheme whose markup the typed model cannot represent is refused before
//! anything is staged. [`Edit::commit`] returns a [`Commit`] with the new
//! package, the exact, source-checked [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_theme`] replays a patch (or its inverse) onto an equal
//! package. An exact no-op keeps every byte, and a changed edit of a signed
//! workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsb::Package;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_theme()?;
//! if edit.theme()?.is_some() {
//!     edit.set_name("Litchi")?;
//! }
//! let (after, patch) = edit.commit()?.into_parts();
//! let restored = after.apply_theme(&patch.inverse())?.into_package();
//! assert_eq!(restored.to_bytes()?, package.to_bytes()?);
//! # Ok(())
//! # }
//! ```

use super::lifecycle::{OwnerDraft, OwnerPatch, read_owner};
use super::{Family, FontSet, Limits, Palette, Snapshot, Theme};
use crate::package::Package;
use crate::package::error::Result;

pub use crate::package::PackageError as Error;
pub use litchi_core::edit::Diagnostics;

/// The result of a Theme edit's `commit`: the new package, the patch and
/// the diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Decoded Theme XML bytes ([`Limits::max_xml_bytes`]).
    XmlBytes,
    /// Theme XML element events ([`Limits::max_nodes`]).
    XmlNodes,
    /// Theme XML nesting depth ([`Limits::max_depth`]).
    XmlDepth,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::XmlBytes => "Theme XML bytes",
            Self::XmlNodes => "Theme XML nodes",
            Self::XmlDepth => "Theme XML depth",
        })
    }
}

/// The configuration of a Theme edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Select the Theme XML policy of the edit; the default is the package's
    /// [`crate::SemanticLimits::theme_value`].
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// The configured Theme XML policy, if any.
    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// The exact, source-checked change of a Theme edit.
#[derive(Debug, Clone)]
pub struct Patch(OwnerPatch);

impl Patch {
    /// Whether the patch changes no Theme byte or owner topology.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.0.is_empty()
    }

    /// The Theme the patch requires before publication; `None` when the
    /// part is absent.
    #[must_use]
    pub fn before(&self) -> Option<&Snapshot> {
        self.0.before().snapshot()
    }

    /// The Theme the patch publishes; `None` when it removes the part.
    #[must_use]
    pub fn after(&self) -> Option<&Snapshot> {
        self.0.after().snapshot()
    }

    /// The reverse patch, restoring the exact source state.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }
}

/// A staged Theme edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Package,
    limits: Limits,
    draft: Option<OwnerDraft>,
}

impl Package {
    /// The optional Theme part, read under the package's
    /// [`crate::SemanticLimits::theme_value`].
    ///
    /// # Errors
    ///
    /// Returns the Theme graph, decoding and limit errors.
    pub fn theme(&self) -> Result<Option<Snapshot>> {
        super::read(self.opc_package(), self.semantic_limits().theme_value())
    }

    /// Start a Theme edit under the package's limits.
    ///
    /// # Errors
    ///
    /// Returns the Theme graph, decoding and limit errors.
    pub fn edit_theme(&self) -> Result<Edit> {
        self.edit_theme_with_options(EditOptions::default())
    }

    /// Start a Theme edit with explicit options.
    ///
    /// # Errors
    ///
    /// Returns the Theme graph, decoding and limit errors.
    pub fn edit_theme_with_options(&self, options: EditOptions) -> Result<Edit> {
        let limits = options
            .limits
            .unwrap_or_else(|| self.semantic_limits().theme_value());
        Ok(Edit {
            source: self.clone(),
            limits,
            draft: None,
        })
    }

    /// Apply a Theme patch to a new package; this one is unchanged. An exact
    /// no-op returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package's Theme closure is
    /// not the patch's before state and [`Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_theme(&self, patch: &Patch) -> Result<Commit> {
        let mut candidate = self.opc_package().clone();
        patch.0.apply(&mut candidate)?;
        if patch.is_noop() {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        Ok(Commit::new(
            self.with_unparsed(candidate),
            patch.clone(),
            Diagnostics::CHANGED,
        ))
    }
}

impl Edit {
    fn ensure_prepared(&mut self) -> Result<()> {
        if self.draft.is_some() {
            return Ok(());
        }
        self.draft = Some(read_owner(self.source.opc_package(), self.limits)?.edit());
        Ok(())
    }

    fn draft(&self) -> Result<&OwnerDraft> {
        self.draft
            .as_ref()
            .ok_or(Error::InvalidFormat("Theme edit is not prepared"))
    }

    fn draft_mut(&mut self) -> Result<&mut OwnerDraft> {
        self.draft
            .as_mut()
            .ok_or(Error::InvalidFormat("Theme edit is not prepared"))
    }

    /// The staged typed Theme, or `None` when the part is absent or its
    /// removal is staged.
    #[must_use]
    pub fn theme(&mut self) -> Result<Option<&Theme>> {
        self.ensure_prepared()?;
        Ok(self.draft()?.theme())
    }

    /// The part name of the present Theme, or the one a new Theme is
    /// created at.
    #[must_use]
    pub fn part_name(&mut self) -> Result<&str> {
        self.ensure_prepared()?;
        Ok(self.draft()?.before().part_name())
    }

    /// The staged applied-theme family metadata.
    #[must_use]
    pub fn family(&mut self) -> Result<Option<&Family>> {
        self.ensure_prepared()?;
        Ok(self.draft()?.family())
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let before = self.draft()?.before().clone();
        let draft = before.edit();
        self.draft = Some(draft);
        Ok(())
    }

    /// Create, replace or remove the complete typed Theme. Returns whether
    /// the staged state changed.
    ///
    /// # Errors
    ///
    /// Returns the Theme model validation errors; a new Theme needs a
    /// non-empty name.
    pub fn replace(&mut self, theme: Option<Theme>) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.replace(theme)
    }

    /// Create or replace the complete typed Theme.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::replace`].
    pub fn set_theme(&mut self, theme: Theme) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.set_theme(theme)
    }

    /// Change the Theme display name.
    ///
    /// # Errors
    ///
    /// Returns an error when the Theme is absent and the model validation
    /// errors.
    pub fn set_name(&mut self, name: impl Into<String>) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.set_name(name)
    }

    /// Replace the complete twelve-slot color palette.
    ///
    /// # Errors
    ///
    /// Returns an error when the Theme is absent and the model validation
    /// errors.
    pub fn set_palette(&mut self, colors: Palette) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.set_palette(colors)
    }

    /// Replace the complete major/minor font set.
    ///
    /// # Errors
    ///
    /// Returns an error when the Theme is absent and the model validation
    /// errors.
    pub fn set_fonts(&mut self, fonts: FontSet) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.set_fonts(fonts)
    }

    /// Set or replace the applied-theme family metadata.
    ///
    /// # Errors
    ///
    /// Returns the family validation errors.
    pub fn set_family(&mut self, family: Family) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.set_family(family)
    }

    /// Remove the applied-theme family metadata while retaining the Theme.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the verb shape.
    pub fn remove_family(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.remove_family()
    }

    /// Remove the Theme part and its Workbook owner edge.
    ///
    /// # Errors
    ///
    /// Infallible today; the `Result` keeps the verb shape.
    pub fn remove(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        self.draft_mut()?.remove()
    }

    /// Validate the staged Theme, publish it to a new package and read it
    /// back. An unchanged Theme returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook, a
    /// stale source closure, and the Theme codec and read-back errors.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let draft = self
            .draft
            .take()
            .ok_or(Error::InvalidFormat("Theme edit is not prepared"))?;
        let commit = draft.commit()?;
        let changed = commit.changed();
        let (_, patch) = commit.into_parts();
        let patch = Patch(patch);
        if !changed {
            return Ok(Commit::new(self.source, patch, Diagnostics::UNCHANGED));
        }
        crate::signature_policy::ensure_unsigned(self.source.opc_package())?;
        let mut candidate = self.source.opc_package().clone();
        patch.0.apply(&mut candidate)?;
        Ok(Commit::new(
            self.source.with_unparsed(candidate),
            patch,
            Diagnostics::CHANGED,
        ))
    }
}
