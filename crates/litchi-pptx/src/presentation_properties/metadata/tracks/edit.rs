//! Immutable media caption-track edits of a [`Package`] (record 0351 §2.3).
//!
//! [`Package::edit_caption_tracks`] starts an [`Edit`] over a staged copy of the
//! package; its verbs run on the staged graph and [`Edit::update`] stages
//! one checked [`Draft`] of the caption tracks of one media picture (by slide and shape ID). [`Edit::commit`] returns a
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
use super::codec::Limits;
#[allow(
    unused_imports,
    reason = "the delegated signatures name the module's values"
)]
use super::*;
use crate::part_edit::{Staged, family};
use crate::slide::Selector;
use crate::{Package, Result};

family!("media caption-track");

/// Per-edit bounds for media caption-track operations.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    #[must_use]
    pub const fn limits_value(&self) -> Option<Limits> {
        self.limits
    }
}

/// A staged media caption-track edit of one [`Package`]. Dropping it discards it.
#[derive(Debug)]
pub struct Edit<'a> {
    staged: Staged<'a>,
    limits: Limits,
}

impl Package {
    /// Start a media caption-track edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose typed
    /// mutations are refused.
    pub fn edit_caption_tracks(&self) -> Result<Edit<'_>> {
        self.edit_caption_tracks_with_options(EditOptions::default())
    }

    /// Start a media caption-track edit under explicit operation options.
    pub fn edit_caption_tracks_with_options(&self, options: EditOptions) -> Result<Edit<'_>> {
        let staged = Staged::new(self, "edit_caption_tracks")?;
        let limits = options
            .limits_value()
            .unwrap_or_else(|| staged.tracks_limits())
            .constrained_by_source(self.opc.read_limits());
        Ok(Edit { staged, limits })
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
        let semantic_limits = self.semantic_limits();
        let package = patch.0.apply_with_replay_validation(
            self,
            "apply_caption_tracks",
            |_, target, patch| {
                patch.validate_changed_parts(target, |name, part| {
                    crate::presentation_properties::metadata::validate_replay_part_policy(
                        target,
                        crate::presentation_properties::metadata::ReplayFamily::Tracks,
                        name,
                        part,
                        semantic_limits,
                    )
                })
            },
        )?;
        Ok(commit_of((package, patch.0.clone())))
    }
}

impl Edit<'_> {
    /// Stage one checked [`Draft`] of the caption tracks of one media picture (by slide and shape ID): `operation` edits the
    /// draft and the validated result replaces the staged source. Any error
    /// leaves the staged graph unchanged.
    ///
    /// # Errors
    ///
    /// Returns the operation's error, a validation or serialization error of
    /// the draft, or a changed edit of a signed package.
    pub fn update<'s, T>(
        &mut self,
        slide: impl Into<Selector<'s>>,
        shape_id: u32,
        operation: impl FnOnce(&mut Draft) -> Result<T>,
    ) -> Result<T> {
        let key = media_key(self.staged.opc(), slide.into(), shape_id)?;
        let limits = self.limits;
        self.staged.update(move |opc| {
            let mut draft = load_media(opc, &key, limits)?
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

    /// Discover the typed TracksInfo/narration metadata of the media picture
    /// with the shape ID `shape_id` (its `p:cNvPr/@id`) on the selected slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn media<'s>(
        &self,
        slide: impl Into<Selector<'s>>,
        shape_id: u32,
    ) -> Result<Option<Snapshot>> {
        let key = media_key(self.staged.opc(), slide.into(), shape_id)?;
        load_media(self.staged.opc(), &key, self.limits)
    }

    /// The staged caption tracks of the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn caption_tracks(&self) -> Result<Vec<Track>> {
        load(self.staged.opc(), self.limits)
    }

    /// Create or replace one caption track.
    ///
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn set(&mut self, value: &Track) -> Result<()> {
        let limits = self.limits;
        self.staged
            .update(move |opc| package::store(opc, value, limits))
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

/// The media key of the picture `shape_id` on the selected slide.
fn media_key(opc: &litchi_opc::OpcPackage, slide: Selector<'_>, shape_id: u32) -> Result<MediaKey> {
    Ok(MediaKey {
        slide_part_name: crate::package::resolve_slide(opc, slide)?
            .as_str()
            .to_owned(),
        shape_id,
    })
}
