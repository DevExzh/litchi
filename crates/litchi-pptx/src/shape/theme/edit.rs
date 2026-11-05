//! Immutable theme edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_themes`](crate::Package::edit_themes) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_themes`](crate::Package::apply_themes) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).
//!
//! Theme parts are package-level: several masters, notes and handout masters
//! may share one, and an added theme has no owner until it is attached, so
//! a theme is named by its part name (`/ppt/theme/themeN.xml`, the
//! [`Authored`] value [`Edit::add`] returns). Masters are selected by
//! [`MasterId`](crate::slide::MasterId), override owners by
//! [`Owner`](crate::slide::Owner).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::package as raw;
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};
use litchi_drawingml::theme::Theme;
use litchi_opc::PackURI;

family!("theme");

/// A staged theme edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a theme edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_themes(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_themes")?,
        })
    }

    /// Apply a theme patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_themes(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_themes")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Read one theme part by URI.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn theme(&self, theme_name: &str) -> Result<Theme> {
        load(self.staged.opc(), theme_name)
    }

    /// Read the permissive summary view used by producer-facing theme inspection.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn summary(&self, theme_name: &str) -> Result<ThemeSummary> {
        raw::load_summary(self.staged.opc(), theme_name)
    }

    /// Read an optional override attached to a slide or layout.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn theme_override<'o>(
        &self,
        owner: impl Into<crate::slide::Owner<'o>>,
    ) -> Result<Option<Override>> {
        let owner_part_name = crate::master_layout::resolve_owner(self.staged.opc(), owner.into())?;
        load_override(self.staged.opc(), owner_part_name.as_str())
    }

    /// The theme part a presentation's colors and fonts come from: the
    /// presentation part's own theme relationship, else the first slide
    /// master's (record 0321; OfficeCLI's `GetThemePart` order). `None` for a
    /// presentation with neither.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed presentation part, master list or
    /// relationship target.
    pub fn presentation_theme(&self) -> Result<Option<PackURI>> {
        presentation_theme(self.staged.opc())
    }

    /// Add a validated theme part at the next free `/ppt/theme/themeN.xml` URI.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add(&mut self, name: &str, colors: &Palette, fonts: &FontSet) -> Result<Authored> {
        self.staged
            .update(move |opc| raw::add(opc, name, colors, fonts))
    }

    /// Attach one theme part to a slide master, replacing any previous theme link.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn attach(&mut self, master: crate::slide::MasterId, theme_name: &str) -> Result<String> {
        let master_part_name = crate::master_layout::resolve_master(self.staged.opc(), master)?;
        self.staged
            .update(move |opc| raw::attach(opc, master_part_name.as_str(), theme_name))
    }

    /// Replace only the color palette of an existing theme part.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_colors(&mut self, theme_name: &str, colors: &Palette) -> Result<()> {
        self.staged
            .update(move |opc| raw::put_colors(opc, theme_name, colors))
    }

    /// Replace only the font set of an existing theme part.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_fonts(&mut self, theme_name: &str, fonts: &FontSet) -> Result<()> {
        self.staged
            .update(move |opc| raw::put_fonts(opc, theme_name, fonts))
    }

    /// Create or replace a slide/layout theme override.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_override<'o>(
        &mut self,
        owner: impl Into<crate::slide::Owner<'o>>,
        value: &Override,
    ) -> Result<String> {
        let owner_part_name = crate::master_layout::resolve_owner(self.staged.opc(), owner.into())?;
        self.staged
            .update(move |opc| raw::put_override(opc, owner_part_name.as_str(), value))
    }

    /// Remove an override and delete its part when no other parent references it.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first with [`Package::unsign`](crate::Package::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_override<'o>(
        &mut self,
        owner: impl Into<crate::slide::Owner<'o>>,
    ) -> Result<bool> {
        let owner_part_name = crate::master_layout::resolve_owner(self.staged.opc(), owner.into())?;
        self.staged
            .update(move |opc| raw::remove_override(opc, owner_part_name.as_str()))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged changes to a new package.
    ///
    /// # Errors
    ///
    /// Returns the part-count limit of the patch.
    pub fn commit(self) -> Result<Commit> {
        Ok(commit_of(self.staged.commit()?))
    }
}
