//! Immutable media caption-track edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_caption_tracks`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of one media object's caption tracks. [`Edit::commit`] returns a
//! [`Commit`] with the new package, the reversible [`Patch`] and the
//! [`Diagnostics`]; [`Package::apply_caption_tracks`] replays a patch onto an equal
//! package. The source package never changes, an exact no-op keeps every
//! byte, and a changed edit of a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy).

#![allow(
    clippy::missing_errors_doc,
    reason = "each verb documents the raw operation it stages"
)]

use super::Draft;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::{Package, Result};

family!("media caption-track");

/// A staged media caption-track edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
}

impl Package {
    /// Start a media caption-track edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_caption_tracks(&self) -> Result<Edit<'_>> {
        Ok(Edit {
            staged: Staged::new(self, "edit_caption_tracks")?,
        })
    }

    /// Apply a media caption-track patch to a new package; this one is unchanged. An
    /// empty patch returns an equal package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state, and the signature refusal for a changed patch of a
    /// signed package.
    pub fn apply_caption_tracks(&self, patch: &Patch) -> Result<Commit> {
        Ok(commit_of((
            patch.0.apply(self, "apply_caption_tracks")?,
            patch.0.clone(),
        )))
    }
}

impl Edit<'_> {
    /// Stage one checked [`Draft`] of one media object's caption tracks: `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<T>(
        &mut self,
        key: &MediaKey,
        operation: impl FnOnce(&mut Draft) -> Result<T>,
    ) -> Result<T> {
        self.staged.update(move |opc| {
            let mut draft = load_media(opc, key)?
                .ok_or_else(|| {
                    Error::refused(
                        litchi_core::edit::RefusalKind::NotFound,
                        "update",
                        "the media object has no caption tracks",
                    )
                })?
                .edit();
            let value = operation(&mut draft)?;
            package::apply_media_commit(opc, draft.commit()?)?;
            Ok(value)
        })
    }

    /// Discover one media picture's typed TracksInfo/narration metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn media(&self, key: &MediaKey) -> Result<Option<Snapshot>> {
        load_media(self.staged.opc(), key)
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn caption_tracks(&self) -> Result<Vec<Track>> {
        load(self.staged.opc())
    }

    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Track) -> Result<()> {
        self.staged.update(move |opc| package::store(opc, value))
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
