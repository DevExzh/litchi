//! Unified PPTX slide and shape removal through the native source/eager transaction.
//!
//! Target resolution stays in the facade so paths and scoped queries have the
//! same diagnostics as the other mutation verbs.  The native opened-package
//! editor remains the authority for relationship ownership, dependency checks,
//! final-slide policy, and the immutable publication proof.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_core::{FileFormat, Position};

use super::node::Item;
use super::structure::{bound, error};
use super::write::{
    Commit, Proof, native_snapshot_budget, opened_edit_options, source_edit_options,
};
use super::{Editable, Nav, Pptx};
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{Code, Error, Result};

/// One resolved native target is the bounded action charge for removal.
/// Native snapshot and patch retention is charged separately from the proof.
const REMOVE_TARGET_BYTES: usize = size_of::<Target>();

// Variant order makes descending iteration remove shapes before slides;
// positions within each slide descend so earlier shape ordinals stay valid.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
enum Target {
    Slide(usize),
    Shape { slide: usize, position: usize },
}

impl Pptx {
    /// Remove slides or top-level shapes selected by a path or scoped query.
    pub(crate) fn remove<'p>(&self, target: TargetArg<'p>) -> Result<Commit> {
        self.remove_with_limits(target, &SemanticLimits::default())
    }

    /// Remove selected elements in one failure-atomic native transaction.
    pub(crate) fn remove_with_limits<'p>(
        &self,
        target: TargetArg<'p>,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        self.check_source()?;
        let result = self
            .remove_inner(target, limits)
            .map_err(|error| error.with_format(FileFormat::Pptx));
        self.check_source()?;
        result
    }

    fn remove_inner<'p>(&self, target: TargetArg<'p>, limits: &SemanticLimits) -> Result<Commit> {
        self.admit_remove_owner()?;
        if let Some(commit) = self.try_remove_smartart_node(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_column(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_row(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_hyperlink(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_text(&target, limits)? {
            return Ok(commit);
        }
        if let Some(commit) = self.try_remove_comment(&target, limits)? {
            return Ok(commit);
        }
        let mut positions = resolve_targets(self, &target, limits)?;
        admit_remove_targets(positions.len(), limits)?;
        reject_overlaps(&positions)?;
        self.check_source()?;

        // Positions are resolved against one immutable snapshot.  Removing in
        // descending order keeps every remaining position valid while one
        // native edit preserves atomic failure and patch ownership.
        positions.sort_unstable();
        let operations = positions.len();
        let source = self.editable_owner()?;
        let native_limit = native_snapshot_budget(limits);
        let (target, proof, changed) = match &source {
            Editable::Source(package) => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                for position in positions.iter().rev().copied() {
                    match position {
                        Target::Slide(slide) => {
                            edit.remove_slide(Position::new(slide))?;
                        },
                        Target::Shape { slide, position } => {
                            edit.remove_shape(Position::new(slide), position)?;
                        },
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (Editable::Source(target), Proof::Source(patch), changed)
            },
            Editable::Eager(package) => {
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_limit))
                    .map_err(super::write::native_error)?;
                for position in positions.iter().rev().copied() {
                    match position {
                        Target::Slide(slide) => {
                            edit.remove_slide(Position::new(slide))?;
                        },
                        Target::Shape { slide, position } => {
                            edit.remove_shape(Position::new(slide), position)?;
                        },
                    }
                }
                let (target, patch) = edit.commit()?.into_parts();
                let changed = !patch.is_noop();
                (
                    Editable::Eager(Arc::new(target)),
                    Proof::Eager(patch),
                    changed,
                )
            },
        };
        let after = if changed {
            match &target {
                Editable::Source(package) => self.adopt_source(package.clone())?,
                Editable::Eager(package) => self.adopt_eager(Arc::clone(package)),
            }
        } else {
            self.clone()
        };
        Ok(Commit::new(after, source, target, proof, operations))
    }

    fn admit_remove_owner(&self) -> Result<()> {
        if self.signature() == super::Signature::Signed {
            return Err(super::signature::refusal(FileFormat::Pptx));
        }
        if !self.0.retain_source {
            return Err(super::source_not_retained());
        }
        Ok(())
    }

    /// Whether this owner can enter native top-level shape removal.
    /// Per-shape dependency and topology admission runs during staging.
    pub(crate) fn supports_shape_removal(&self) -> bool {
        self.0.retain_source && self.signature() != super::Signature::Signed
    }

    /// Whether this owner can enter the native dependency-free slide-removal
    /// path. Per-slide relationship and final-slide admission still happens
    /// during the operation, so this is only an owner-level capability answer.
    pub(crate) fn supports_slide_removal(&self) -> bool {
        self.0.retain_source && self.signature() != super::Signature::Signed
    }
}

fn resolve_targets(
    file: &Pptx,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Target>> {
    let nav = Nav::new(file);
    let mut positions = Vec::new();
    let mut seen = HashSet::new();

    if let Some(path) = target.path() {
        let node = crate::path::resolve::get(&nav, path)?;
        append_target(&mut positions, &mut seen, &node, limits)?;
        return Ok(positions);
    }

    let Some(query) = target.query() else {
        return Err(error(Code::InvalidArgument, "mutation target is empty"));
    };
    let walk = crate::query::run(nav, query)?;
    if !walk.plan().diagnostics().is_empty() {
        return Err(error(
            Code::InvalidSelector,
            "mutation selector contains an unsupported predicate",
        ));
    }

    let mut matched = false;
    for node in walk {
        let node = node?;
        matched = true;
        append_target(&mut positions, &mut seen, &node, limits)?;
    }
    if !matched {
        return Err(error(
            Code::ZeroMatches,
            "mutation selector matched no element",
        ));
    }
    Ok(positions)
}

fn append_target(
    positions: &mut Vec<Target>,
    seen: &mut HashSet<Target>,
    node: &super::Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    let position = match node.item() {
        Item::Slide(slide) => Target::Slide(slide.position as usize),
        Item::Tree(tree) => {
            let Some(slide) = tree.container.slide() else {
                return Err(error(
                    Code::UnsupportedCapability,
                    "master and layout shape removal is not integrated",
                ));
            };
            if tree.handle.parent().is_some() {
                return Err(error(
                    Code::UnsupportedCapability,
                    "shape removal requires a top-level slide shape",
                ));
            }
            Target::Shape {
                slide: slide as usize,
                position: tree.handle.position().get(),
            }
        },
        _ => {
            return Err(error(
                Code::UnsupportedType,
                "PPTX removal requires slide or top-level shape targets",
            ));
        },
    };
    if seen.contains(&position) {
        return Ok(());
    }
    let observed = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| target_limit(usize::MAX, limits.max_expanded_targets()))?;
    admit_remove_targets(observed, limits)?;
    positions
        .try_reserve(1)
        .map_err(|source| allocation("PPTX removal targets", source))?;
    seen.try_reserve(1)
        .map_err(|source| allocation("PPTX removal target identities", source))?;
    seen.insert(position);
    positions.push(position);
    Ok(())
}

fn reject_overlaps(targets: &[Target]) -> Result<()> {
    let slide_count = targets
        .iter()
        .filter(|target| matches!(target, Target::Slide(_)))
        .count();
    if slide_count == 0 {
        return Ok(());
    }
    let mut slides = HashSet::new();
    slides
        .try_reserve(slide_count)
        .map_err(|source| allocation("PPTX removal slide identities", source))?;
    for target in targets {
        if let Target::Slide(slide) = target {
            slides.insert(*slide);
        }
    }
    if targets.iter().any(|target| {
        matches!(target,
        Target::Shape { slide, .. } if slides.contains(slide))
    }) {
        return Err(error(
            Code::InvalidArgument,
            "PPTX removal cannot select both a slide and one of its shapes",
        ));
    }
    Ok(())
}

fn admit_remove_targets(count: usize, limits: &SemanticLimits) -> Result<()> {
    bound(
        LimitKind::ExpandedTargets,
        count,
        limits.max_expanded_targets(),
    )?;
    let staged = count.saturating_mul(REMOVE_TARGET_BYTES);
    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())
}

fn target_limit(observed: usize, maximum: usize) -> Error {
    TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    }
    .into()
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    litchi_core::Error::Allocation { resource, source }.into()
}

#[cfg(test)]
mod tests {
    use crate::{File, IntoTarget};

    #[test]
    fn source_removal_uses_the_source_revision_without_unsigned_promotion() {
        let mut builder = litchi_pptx::builder::Builder::new().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        builder.presentation().unwrap().add_slide().unwrap();
        let bytes = builder.build().unwrap().to_bytes().unwrap();
        let File::Pptx(source) = crate::open_bytes(bytes).unwrap() else {
            panic!("PPTX owner")
        };
        let removed = source
            .remove("/slide[1]".into_target().unwrap())
            .expect("source removal");
        assert_eq!(removed.after.slide_count().unwrap(), 1);
        assert!(removed.after.is_source_backed());
        assert!(matches!(removed.proof, super::Proof::Source(_)));
        assert!(source.0.unsigned_promotion.get().is_none());
    }
}
