//! The web-settings family (record 0355): the document's `webSettings` part.
//!
//! [`Package::web`](crate::Package::web) reads the typed settings and
//! [`load_snapshot`](super::load_snapshot) the source-preserving
//! [`Snapshot`](super::Snapshot); [`Package::edit_web`](crate::Package::edit_web)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_web`](crate::Package::apply_web) replays its [`Patch`].
//! The semantic verbs ([`Edit::set_target_screen_size`] and the rest) rewrite
//! only the elements they change; [`Edit::put`] replaces the whole part with
//! canonical XML.

use super::transaction::Transaction as Draft;
use super::{Border, BorderSide, Conformance, Key, Layout, Screen, Settings};
use crate::package::family::{Staged, graph_family};

graph_family!("web-settings", own_limits);

pub use super::limit::LimitKind;

/// A staged web-settings change over one package snapshot. The source is
/// never changed: [`Self::commit`] publishes a new package, and dropping
/// the edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    draft: Option<Draft>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package with the changes staged before the pending semantic
    /// verbs (those are published by [`Self::commit`]).
    #[must_use]
    pub const fn package(&self) -> &crate::Package {
        self.staged.working()
    }

    fn draft(&mut self) -> Result<&mut Draft> {
        if self.draft.is_none() {
            let snapshot =
                super::load_snapshot(self.staged.working().opc_package(), &super::Limits::DEFAULT)?
                    .ok_or_else(|| crate::Error::PartNotFound("web-settings part".into()))?;
            self.draft = Some(snapshot.edit());
        }
        Ok(self
            .draft
            .as_mut()
            .expect("web-settings draft was created above"))
    }

    /// Publish the pending semantic verbs into the working package.
    fn flush(&mut self) -> Result<()> {
        if let Some(draft) = self.draft.take() {
            let patch = draft.commit()?.into_patch();
            let mut working = self.staged.working().clone();
            working.edit_semantic_opc("web.edit", |opc| {
                super::package::apply_patch(opc, &patch).map(|_| ())
            })?;
            *self.staged.working_mut() = working;
        }
        Ok(())
    }

    /// Replace the web-output settings with `settings` in canonical XML;
    /// returns whether the package changed. Equal settings stage nothing.
    ///
    /// # Errors
    ///
    /// Returns an error for a conformance different from the package's and
    /// the signed-package refusal for a change to a signed package.
    pub fn put(&mut self, settings: Settings, conformance: Conformance) -> Result<bool> {
        self.flush()?;
        let mut working = self.staged.working().clone();
        let changed = working.put_web(settings, conformance)?;
        *self.staged.working_mut() = working;
        Ok(changed)
    }

    /// Remove the document-owned web-settings part; returns whether it was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.flush()?;
        let mut working = self.staged.working().clone();
        let changed = working.remove_web()?;
        *self.staged.working_mut() = working;
        Ok(changed)
    }

    /// Set or remove the target screen size.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part.
    pub fn set_target_screen_size(&mut self, value: Option<Screen>) -> Result<&mut Self> {
        self.draft()?.set_target_screen_size(value);
        Ok(self)
    }

    /// Remove the target screen size.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part.
    pub fn clear_target_screen_size(&mut self) -> Result<&mut Self> {
        self.draft()?.clear_target_screen_size();
        Ok(self)
    }

    /// Set the root frameset size.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part or no
    /// frameset, or for an invalid size.
    pub fn set_frameset_size(&mut self, value: impl Into<String>) -> Result<&mut Self> {
        self.draft()?.set_frameset_size(value)?;
        Ok(self)
    }

    /// Remove the root frameset size.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part.
    pub fn clear_frameset_size(&mut self) -> Result<&mut Self> {
        self.draft()?.clear_frameset_size();
        Ok(self)
    }

    /// Set or remove the root frameset layout.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part or no
    /// frameset.
    pub fn set_frameset_layout(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        self.draft()?.set_frameset_layout(value)?;
        Ok(self)
    }

    /// Remove the root frameset layout.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part or no
    /// frameset.
    pub fn clear_frameset_layout(&mut self) -> Result<&mut Self> {
        self.draft()?.clear_frameset_layout()?;
        Ok(self)
    }

    /// Set or remove one border side of the `div` selected by `key`.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no web-settings part or no such
    /// `div`.
    pub fn set_div_border(
        &mut self,
        key: impl Into<Key>,
        side: BorderSide,
        value: Option<Border>,
    ) -> Result<&mut Self> {
        self.draft()?.set_div_border(key, side, value)?;
        Ok(self)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Validate the staged changes and publish them as a new package.
    /// Nothing staged publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when a semantic rewrite cannot preserve the part, and
    /// the signed-package refusal for a change to a signed package.
    pub fn commit(mut self) -> Result<Commit> {
        self.flush()?;
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a web-settings change over this package. O(1): the edit shares
    /// the package until a verb changes it.
    #[must_use]
    pub fn edit_web(&self) -> Edit {
        Edit {
            staged: Staged::new(self),
            draft: None,
        }
    }

    /// Replay a web-settings patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_web(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_web")
    }
}
