//! Coalesced direct picture removals inside one transaction (record 0280).
//!
//! `Edit::remove("/body/p[N]/picture[M]")` published one native commit per
//! call, so 64 removals cost 52–56× the native batch (record 0275). A run
//! keeps the handle it started from and the native indices of the pictures
//! its commands removed, and publishes them as one native removal, the
//! commit a query `remove` of the same pictures makes.
//!
//! Each command's path is resolved as it would be against the candidate
//! after the earlier removals: removing a picture changes no paragraph and
//! no other picture's anchor, so the `M`th picture of paragraph `N` is the
//! `M`th of the paragraph's original pictures the run has not removed. Any
//! other target, or a path that names no remaining picture, closes the run
//! and takes the ordinary path, which reports its canonical error.

use litchi_core::selector::Path;
use litchi_rtf::edit::MAX_PICTURE_REMOVAL_OPERATIONS;

use super::{Commit, Rtf, coalesce, destination};
use crate::error::Result;
use crate::target::TargetArg;
use crate::transaction::SemanticLimits;

/// Pending direct picture removals against one handle.
pub(crate) struct PictureRun {
    base: Rtf,
    /// Native picture indices of `base`, in command order.
    removed: Vec<usize>,
    first_index: Option<usize>,
}

impl PictureRun {
    pub(crate) fn new(base: &Rtf) -> Self {
        Self {
            base: base.clone(),
            removed: Vec::new(),
            first_index: None,
        }
    }

    /// The transaction index of the first staged command.
    pub(crate) const fn first_index(&self) -> Option<usize> {
        self.first_index
    }

    /// Whether the run already holds the most pictures one native removal
    /// accepts.
    pub(crate) fn is_full(&self) -> bool {
        self.removed.len() >= MAX_PICTURE_REMOVAL_OPERATIONS
    }

    /// The native index of `base` that the exact picture path `target` names
    /// after the staged removals, or `None` when `target` is not such a path
    /// or names no remaining picture.
    pub(crate) fn plan(&self, target: &TargetArg<'_>) -> Option<usize> {
        let path = target.path()?;
        if path.raw
            || !matches!(
                path.origin(),
                crate::path::arg::Origin::Typed | crate::path::arg::Origin::Text(_)
            )
        {
            return None;
        }
        let (paragraph, ordinal) = picture_path(path.path())?;
        let paragraph = self.base.0.package.body().paragraphs().nth(paragraph)?;
        crate::format::rtf::stories::pictures(
            &self.base,
            crate::format::rtf::stories::anchor_range(paragraph),
        )
        .filter(|index| !self.removed.contains(index))
        .nth(ordinal)
    }

    /// Stages the removal of the native picture `picture` for the command at
    /// transaction index `index`.
    pub(crate) fn stage(&mut self, picture: usize, index: usize) {
        self.first_index.get_or_insert(index);
        self.removed.push(picture);
    }

    /// The base handle and the one native removal of every staged picture.
    ///
    /// # Errors
    ///
    /// The native removal's error, as a direct removal reports it.
    pub(crate) fn commit(self, limits: &SemanticLimits) -> Result<(Rtf, Commit)> {
        self.base.admit_edit()?;
        let mut indices = self.removed;
        indices.sort_unstable();
        let commit = destination::remove_pictures(&self.base, &indices, limits)?;
        Ok((self.base, commit))
    }
}

/// `(paragraph, picture)` of the exact path `/body/p[N]/picture[M]` (or
/// `image`, `img`), both zero-based.
fn picture_path(path: &Path<'_>) -> Option<(usize, usize)> {
    let [body, paragraph, picture] = path.steps() else {
        return None;
    };
    if !path.is_absolute() || !coalesce::body_step(body) {
        return None;
    }
    let paragraph_named = ["p", "paragraph"]
        .iter()
        .any(|name| paragraph.name().eq_ignore_ascii_case(name));
    let picture_named = ["picture", "image", "img"]
        .iter()
        .any(|name| picture.name().eq_ignore_ascii_case(name));
    if !(paragraph_named && picture_named) {
        return None;
    }
    Some((paragraph.index()?.get(), picture.index()?.get()))
}
