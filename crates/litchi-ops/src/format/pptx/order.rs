//! PPTX slide-order mutations through the native transaction.
//!
//! The unified adapter resolves only structural slide paths and placement
//! vocabulary.  The native eager or source-primary edit owns the
//! presentation XML, positional-owner checks, immutable patch, and
//! publication proof.

use super::node::Item;
use super::structure::{bound, error};
use super::write::{
    Commit, Proof, native_snapshot_budget, opened_edit_options, source_edit_options,
};
use super::{Editable, Nav, Pptx};
use crate::path::PathArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::{Code, Placement, Result};
use litchi_core::{FileFormat, Position};
use std::sync::Arc;

/// One selected slide position is the bounded action charge for a slide-order
/// operation.  The native patch's retained bytes are accounted separately by
/// the transaction snapshot limit.
const ORDER_TARGET_BYTES: usize = size_of::<usize>();

impl Pptx {
    // Admit the owner without promoting a source-backed presentation.  The
    // native source transaction retains the positional source and returns its
    // own exact-source proof.
    fn admit_order_owner(&self) -> Result<()> {
        if self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Pptx));
        }
        if !self.0.retain_source {
            return Err(super::source_not_retained());
        }
        Ok(())
    }

    /// Move one slide to a final sibling position.
    pub(crate) fn move_to<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
    ) -> Result<Commit> {
        self.move_to_with_limits(path, placement, &SemanticLimits::default())
    }

    /// Move one slide with the caller's semantic expansion and staging limits.
    pub(crate) fn move_to_with_limits<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.check_source()?;
        let result = self
            .move_inner(path, placement, limits)
            .map_err(|error| error.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }

    fn move_inner<'p, 'q>(
        &self,
        path: PathArg<'p>,
        placement: Placement<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_order_owner()?;
        let nav = Nav::new(self);
        let selected = crate::path::resolve::get(&nav, &path)?;
        let from = slide_position(selected.item())?;
        let count = self.slides_in_catalog()?;
        let to = move_destination(&nav, from, count, &placement)?;

        admit_order_targets(1, limits)?;
        let source = self.editable()?;
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                edit.move_slide(Position::new(from), Position::new(to))?;
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                edit.move_slide(Position::new(from), Position::new(to))?;
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        let after = if !changed {
            self.clone()
        } else {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        };
        Ok(Commit::new(after, source, target, proof, 1))
    }

    /// Exchange two slides in one native transaction.
    pub(crate) fn swap<'p, 'q>(&self, left: PathArg<'p>, right: PathArg<'q>) -> Result<Commit> {
        self.swap_with_limits(left, right, &SemanticLimits::default())
    }

    /// Exchange two slides with the caller's semantic limits.
    pub(crate) fn swap_with_limits<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.check_source()?;
        let result = self
            .swap_inner(left, right, limits)
            .map_err(|error| error.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }

    fn swap_inner<'p, 'q>(
        &self,
        left: PathArg<'p>,
        right: PathArg<'q>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.admit_order_owner()?;
        let nav = Nav::new(self);
        let left = slide_position(crate::path::resolve::get(&nav, &left)?.item())?;
        let right = slide_position(crate::path::resolve::get(&nav, &right)?.item())?;
        let target_count = if left == right { 1 } else { 2 };
        admit_order_targets(target_count, limits)?;

        let source = self.editable()?;
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                edit.swap_slides(Position::new(left), Position::new(right))?;
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                edit.swap_slides(Position::new(left), Position::new(right))?;
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        let after = if !changed {
            self.clone()
        } else {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        };
        Ok(Commit::new(after, source, target, proof, 1))
    }
}

fn slide_position(item: &Item) -> Result<usize> {
    let Item::Slide(slide) = item else {
        return Err(error(
            Code::UnsupportedType,
            "PPTX slide ordering requires slide targets",
        ));
    };
    Ok(slide.position as usize)
}

fn move_destination(
    nav: &Nav<'_>,
    from: usize,
    count: usize,
    placement: &Placement<'_>,
) -> Result<usize> {
    if count == 0 {
        return Err(error(Code::NotFound, "the PPTX presentation has no slides"));
    }
    let last = count - 1;
    match placement {
        Placement::Start => Ok(0),
        Placement::End => Ok(last),
        Placement::Position(position) => Ok(position.get()),
        Placement::Before(path) => {
            let anchor = slide_position(crate::path::resolve::get(nav, path)?.item())?;
            if from == anchor {
                Ok(from)
            } else if from < anchor {
                Ok(anchor - 1)
            } else {
                Ok(anchor)
            }
        },
        Placement::After(path) => {
            let anchor = slide_position(crate::path::resolve::get(nav, path)?.item())?;
            if from == anchor {
                Ok(from)
            } else if from < anchor {
                Ok(anchor)
            } else {
                anchor
                    .checked_add(1)
                    .ok_or_else(|| error(Code::OutOfRange, "slide position overflow"))
            }
        },
        Placement::Into(path, position) => {
            let parent = crate::path::resolve::get(nav, path)?;
            if !matches!(parent.item(), Item::Presentation) {
                return Err(error(
                    Code::UnsupportedType,
                    "PPTX slides can only be reordered within the presentation root",
                ));
            }
            Ok(position.map_or(last, Position::get))
        },
    }
}

fn admit_order_targets(count: usize, limits: &SemanticLimits) -> Result<()> {
    bound(
        LimitKind::ExpandedTargets,
        count,
        limits.max_expanded_targets(),
    )?;
    bound(
        LimitKind::StagedBytes,
        count.saturating_mul(ORDER_TARGET_BYTES),
        limits.max_staged_bytes(),
    )
}
