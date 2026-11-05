//! The content-control family (record 0355): checksums, data-binding
//! checksums and formatting exceptions of the package's content controls.
//!
//! [`Package::content_control_snapshot`](crate::Package::content_control_snapshot)
//! reads every story's controls;
//! [`Package::edit_content_controls`](crate::Package::edit_content_controls)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_content_controls`](crate::Package::apply_content_controls)
//! replays its [`Patch`].

use litchi_opc::PackURI;

use super::package::PackageDraft;
use super::{Checksum, FormattingAllowed};
use crate::package::family::{Staged, graph_family};

graph_family!("content-control", own_limits);

pub use super::limit::LimitKind;

/// Options of a content-control edit (record 0355a).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<super::PackageLimits>,
}

impl EditOptions {
    /// Bound the edit's snapshot and its published stories (including
    /// `max_output_bytes`) by `limits` instead of the package's semantic
    /// content-control limits.
    #[must_use]
    pub fn limits(mut self, limits: super::PackageLimits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// The configured limits; `None` uses the package's semantic limits.
    #[must_use]
    pub const fn limits_value(&self) -> Option<&super::PackageLimits> {
        self.limits.as_ref()
    }
}

/// A staged content-control change over one package snapshot. The source is
/// never changed: [`Self::commit`] publishes a new package, and dropping the
/// edit discards it.
#[derive(Clone)]
pub struct Edit {
    staged: Staged,
    limits: super::PackageLimits,
    draft: Option<PackageDraft>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn draft(&mut self) -> Result<&mut PackageDraft> {
        if self.draft.is_none() {
            self.draft = Some(
                self.staged
                    .working()
                    .content_control_snapshot_bounded(self.limits.clone())?
                    .edit()?,
            );
        }
        Ok(self
            .draft
            .as_mut()
            .expect("content-control draft was created above"))
    }

    /// Set, repair or remove the checksum of the control at `occurrence`
    /// in the story `part`.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown story or occurrence, or a control
    /// that has no checksum slot.
    pub fn set_checksum(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        self.draft()?.set_checksum(part, occurrence, value)?;
        Ok(self)
    }

    /// Set or remove the checksum of one data binding of a control.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown story, occurrence or binding.
    pub fn set_binding_checksum(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        binding: usize,
        value: Option<Checksum>,
    ) -> Result<&mut Self> {
        self.draft()?
            .set_binding_checksum(part, occurrence, binding, value)?;
        Ok(self)
    }

    /// Recompute the checksum of one control from its bound data.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown story or occurrence, or unresolvable
    /// bound data.
    pub fn refresh_checksum(&mut self, part: &PackURI, occurrence: usize) -> Result<&mut Self> {
        self.draft()?.refresh_checksum(part, occurrence)?;
        Ok(self)
    }

    /// Recompute every bound control's checksum.
    ///
    /// # Errors
    ///
    /// Returns an error for unresolvable bound data.
    pub fn refresh_checksums(&mut self) -> Result<&mut Self> {
        self.draft()?.refresh_checksums()?;
        Ok(self)
    }

    /// Set or remove the formatting exception on one exact lock.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown story or occurrence.
    pub fn set_formatting_allowed(
        &mut self,
        part: &PackURI,
        occurrence: usize,
        value: Option<FormattingAllowed>,
    ) -> Result<&mut Self> {
        self.draft()?
            .set_formatting_allowed(part, occurrence, value)?;
        Ok(self)
    }

    /// Discard every staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.staged.reset();
        self.draft = None;
    }

    /// Publish the staged changes as a new package with their exact patch.
    ///
    /// # Errors
    ///
    /// Returns an error when a story rewrite cannot preserve its markup, and
    /// the signed-package refusal for a change to a signed package.
    pub fn commit(mut self) -> Result<Commit> {
        if let Some(draft) = self.draft.take() {
            let planned = draft.commit()?;
            let mut working = self.staged.working().clone();
            working.publish_content_control_draft(&planned)?;
            *self.staged.working_mut() = working;
        }
        commit_staged(self.staged)
    }
}

impl crate::Package {
    /// Stage a content-control change over this package. O(1).
    #[must_use]
    pub fn edit_content_controls(&self) -> Edit {
        self.edit_content_controls_with_options(EditOptions::default())
    }

    /// Stage a content-control change over this package under `options`.
    /// O(1).
    #[must_use]
    pub fn edit_content_controls_with_options(&self, options: EditOptions) -> Edit {
        Edit {
            staged: Staged::new(self),
            limits: options
                .limits
                .unwrap_or_else(|| self.semantic_limits().content_controls_value().clone()),
            draft: None,
        }
    }

    /// Replay a content-control patch onto this package, which must hold the
    /// patch's before state.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] for a package that does not
    /// hold the before state.
    pub fn apply_content_controls(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch, "apply_content_controls")
    }
}
