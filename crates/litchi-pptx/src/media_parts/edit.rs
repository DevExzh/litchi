//! Immutable slide media edits of a [`Package`](crate::Package) (record 0356a).
//!
//! [`Package::edit_media`](crate::Package::edit_media) starts an [`Edit`]
//! over a staged copy of the package; its verbs run on the staged graph.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_media`](crate::Package::apply_media) replays a patch
//! onto an equal package. The source package never changes, an exact no-op
//! keeps every byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]

use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("slide media");

/// A staged slide media edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a slide media edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_media(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_media")?,
        })
    }

    /// Apply a slide media patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_media(&self, patch: &Patch) -> Result<Commit> {
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_media",
            |package, target, patch| {
                let limits = package
                    .semantic_limits()
                    .media_parts()
                    .constrained_by_source(target.read_limits());
                let mut changed_resources = Vec::new();
                let mut changed_slides = Vec::new();
                patch.validate_changed_parts(target, |name, part| {
                    validate_replay_part(target, name, part, limits)?;
                    if part.content_type() == litchi_opc::constants::content_type::PML_SLIDE {
                        changed_slides
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "media replay changed slide names",
                                source,
                            })?;
                        changed_slides.push(name.clone());
                    } else if name.as_str().starts_with("/ppt/media/") {
                        changed_resources
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "media replay changed resource names",
                                source,
                            })?;
                        changed_resources.push(name.clone());
                    }
                    Ok(())
                })?;
                let index = patch.replay_index_after(package, target)?;
                validate_replay_references(
                    target,
                    &index,
                    &changed_resources,
                    &changed_slides,
                    limits,
                )
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    fn limits(&self) -> Limits {
        self.staged
            .semantic_limits()
            .media_parts()
            .constrained_by_source(self.staged.opc().read_limits())
    }

    /// The staged media list of the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn media<'s>(&self, slide: impl Into<crate::slide::Selector<'s>>) -> Result<List> {
        let slide_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        package::load_with_limits(self.staged.opc(), &slide_name, self.limits())
    }

    /// Adds a new set of media pictures and their inert internal resources to a Slide part.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        value: &List,
        conformance: Conformance,
    ) -> Result<()> {
        let slide_name = crate::package::resolve_slide(self.staged.opc(), slide.into())?;
        let limits = self.limits();
        self.staged.update(move |opc| {
            store_with_limits(opc, &slide_name, value, conformance, false, limits)
        })
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
