//! Placeholder and speaker-notes additions (record 0346).
//!
//! `add(/slide[N], placeholder, {phType, text})` appends a placeholder bound
//! to the slide layout's slot by OfficeCLI's rules (the leaf's
//! `add_placeholder`) and returns its canonical `/slide[N]/shape[K]` path;
//! `add(/slide[N], notes, {text})` creates the slide's notes page when it
//! has none and gives it the text, as OfficeCLI's `AddNotes` does, and
//! returns `/slide[N]/notes`. Both stage on the eager editor; a source
//! owner is promoted, as notes and comment writes promote it.

use super::node::Item;
use super::spec::{Kind, PLACEHOLDER_ADD_TABLE, PLACEHOLDER_TYPE, PlaceholderSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options, validate_text};
use super::{Editable, Nav, Pptx, steps};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Result};
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use litchi_pptx::PlaceholderKind;
use std::sync::Arc;

/// The leaf kind of a canonical `phType` token, in [`PLACEHOLDER_TYPE`]'s
/// order.
fn placeholder_kind(token: &str) -> Option<PlaceholderKind> {
    Some(match PLACEHOLDER_TYPE.get(token)? {
        "title" => PlaceholderKind::Title,
        "ctrTitle" => PlaceholderKind::CenteredTitle,
        "body" => PlaceholderKind::Body,
        "subtitle" => PlaceholderKind::Subtitle,
        "date" => PlaceholderKind::DateTime,
        "footer" => PlaceholderKind::Footer,
        "slidenum" => PlaceholderKind::SlideNumber,
        "header" => PlaceholderKind::Header,
        "obj" => PlaceholderKind::Object,
        "chart" => PlaceholderKind::Chart,
        "table" => PlaceholderKind::Table,
        "clipart" => PlaceholderKind::ClipArt,
        "diagram" => PlaceholderKind::Diagram,
        "media" => PlaceholderKind::Media,
        "picture" => PlaceholderKind::Picture,
        "sldImg" => PlaceholderKind::SlideImage,
        _ => return None,
    })
}

/// The slide a structural addition names as its parent.
fn slide_parent(file: &Pptx, parent: &PathArg<'_>, what: &'static str) -> Result<u32> {
    let parent = crate::path::resolve::get(&Nav::new(file), parent)?;
    match parent.item() {
        Item::Slide(slide) => Ok(slide.position),
        _ => Err(error(Code::UnsupportedType, what)),
    }
}

fn end_only(node: &NewNode<'_>, what: &'static str) -> Result<()> {
    if matches!(node.placement(), None | Some(Placement::End)) {
        Ok(())
    } else {
        Err(error(Code::UnsupportedCapability, what)
            .with_suggestion("omit placement or use Placement::End"))
    }
}

impl Pptx {
    pub(super) fn add_placeholder(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let slide = slide_parent(self, &parent, "placeholders require a slide parent")?;
        end_only(&node, "placeholder insertion supports end placement")?;
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut kind = None;
        let mut text = None;
        let mut assigned = 0_u8;
        for (key, value) in node.props().iter() {
            let (index, entry) = PLACEHOLDER_ADD_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u8 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a placeholder property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match (setter, scalar) {
                    (PlaceholderSetter::Kind, Some(Scalar::Token(token))) => {
                        kind = placeholder_kind(token);
                    },
                    (PlaceholderSetter::Text, Some(Scalar::Text(value))) => {
                        bound(
                            LimitKind::StagedBytes,
                            value.len(),
                            limits.max_staged_bytes(),
                        )?;
                        validate_text(value)?;
                        text = Some(value.to_owned());
                    },
                    (PlaceholderSetter::Text, None) => text = Some(String::new()),
                    _ => unreachable!("placeholder property coercion"),
                }
                Ok(())
            })?;
        }
        let kind = kind.ok_or_else(|| {
            error(
                Code::MissingProperty,
                "placeholder insertion requires its type",
            )
            .with_property("phType")
            .with_suggestion("pass phType=title, body, subtitle, date, footer, slidenum, ...")
        })?;
        let scene = self
            .slide_scene(slide as usize)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        // The new shape's canonical path is its `shape[K]` ordinal, as
        // OfficeCLI returns it and as queries print placeholders.
        let mut ordinal = 0_usize;
        let mut title = false;
        for handle in scene.roots() {
            if steps::tree_kind(&handle) != Some(Kind::Shape) {
                continue;
            }
            ordinal += 1;
            if let Some(placeholder) = handle.view().placeholder() {
                title |= matches!(placeholder.kind(), Some("title" | "ctrTitle"));
            }
        }
        drop(scene);
        if title
            && matches!(
                kind,
                PlaceholderKind::Title | PlaceholderKind::CenteredTitle
            )
        {
            return Err(error(
                Code::InvalidArgument,
                "the slide already has a title placeholder",
            )
            .with_property("phType")
            .with_suggestion("set the existing title placeholder's text instead"));
        }
        let source = self.editable_owner()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let _ = edit.add_placeholder(slide as usize, kind, text.as_deref().unwrap_or(""))?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide as usize))),
            Step::new("shape").with(Predicate::Index(Position::new(ordinal))),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }

    pub(super) fn add_notes(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let slide = slide_parent(self, &parent, "speaker notes require a slide parent")?;
        end_only(&node, "notes take no placement")?;
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let text = super::write::notes_add_text(node.props(), limits)?;
        let source = self.editable_owner()?;
        let base = self.eager_base(&source)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let changed = edit.set_notes_text(slide as usize, &text)?;
        let (target, patch) = edit.commit()?.into_parts();
        let changed = changed && !patch.is_noop();
        let target = Arc::new(target);
        let after = if changed {
            self.adopt_eager(Arc::clone(&target))
        } else {
            self.clone()
        };
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide as usize))),
            Step::new("notes"),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }
}
