//! Typed text-box additions through the native immutable edit.
use super::node::Item;
use super::spec::{Kind, SHAPE_WRITE_TABLE, ShapeSetter};
use super::structure::{bound, error};
use super::write::text::TextValues;
use super::write::{
    Commit, Proof, native_snapshot_budget, opened_edit_options, source_edit_options, validate_text,
};
use super::{Editable, Nav, Pptx, steps};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::transaction::{Error as TxError, LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Result};
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use std::sync::Arc;

impl Pptx {
    pub(super) fn add_text_box(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        source: Editable,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let parent = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "text boxes require a slide parent",
            ));
        };
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "text-box insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        const GEOMETRY_BYTES: usize = 4 * size_of::<i64>();
        bound(
            LimitKind::StagedBytes,
            GEOMETRY_BYTES,
            limits.max_staged_bytes(),
        )?;
        let mut text = None;
        let mut geometry = [None; 4];
        // Run and paragraph formatting of the new body (`size`, `bold`, ...).
        let mut format = TextValues::default();
        // The click action (record 0325), written in the same native edit.
        let mut click = super::links::ClickInput::default();
        for (key, value) in node.props().iter() {
            let (_, entry) = SHAPE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match setter {
                    ShapeSetter::Text => {
                        if text.is_some() {
                            return Err(error(
                                Code::InvalidArgument,
                                "text was assigned more than once",
                            )
                            .with_property("text"));
                        }
                        let value = match scalar {
                            Some(Scalar::Text(value)) => value,
                            None => "",
                            _ => unreachable!("text coercion"),
                        };
                        bound(
                            LimitKind::StagedBytes,
                            value.len().saturating_add(GEOMETRY_BYTES),
                            limits.max_staged_bytes(),
                        )?;
                        validate_text(value)?;
                        let mut owned = String::new();
                        owned
                            .try_reserve_exact(value.len())
                            .map_err(|source| crate::Error::from(TxError::Allocation { source }))?;
                        owned.push_str(value);
                        text = Some(owned);
                    },
                    ShapeSetter::Format(setter) => {
                        if format.has(setter) {
                            return Err(error(
                                Code::InvalidArgument,
                                "a text property was assigned more than once",
                            )
                            .with_property(key.as_str().to_owned()));
                        }
                        format.stage(setter, scalar, limits)?;
                    },
                    ShapeSetter::Click(key) => click.stage(key, scalar, limits)?,
                    setter => {
                        let (index, name) = match setter {
                            ShapeSetter::X => (0, "x"),
                            ShapeSetter::Y => (1, "y"),
                            ShapeSetter::Width => (2, "width"),
                            ShapeSetter::Height => (3, "height"),
                            ShapeSetter::Text | ShapeSetter::Format(_) | ShapeSetter::Click(_) => {
                                unreachable!("text, formatting and clicks are matched above")
                            },
                        };
                        if geometry[index].is_some() {
                            return Err(error(
                                Code::InvalidArgument,
                                "geometry was assigned more than once",
                            )
                            .with_property(name));
                        }
                        let Some(Scalar::Length(value)) = scalar else {
                            unreachable!("geometry coercion")
                        };
                        geometry[index] = Some(value.emu());
                    },
                }
                Ok(())
            })?;
        }
        let required = |index: usize, name: &'static str| {
            geometry[index].ok_or_else(|| {
                error(
                    Code::MissingProperty,
                    "text boxes require explicit positive width and height",
                )
                .with_property(name)
            })
        };
        let bounds = (
            geometry[0].unwrap_or(0),
            geometry[1].unwrap_or(0),
            required(2, "width")?,
            required(3, "height")?,
        );
        let slide = slide.position as usize;
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        // Ordinary shape paths count direct p:sp siblings, not mixed drawing
        // elements or group descendants. Reuse the read path's classification.
        let ordinal = scene
            .roots()
            .filter(|shape| steps::tree_kind(shape) == Some(Kind::Shape))
            .count();
        // The native text target counts every shape in pre-order; the new
        // box is appended after all of them.
        let scene_index = Position::new(scene.len());
        let properties = format.properties();
        let body = [litchi_pptx::opened::TextAssignment::new(
            litchi_pptx::opened::TextTarget::Shape { shape: scene_index },
            properties,
        )];
        let body: &[litchi_pptx::opened::TextAssignment<'_>] =
            if format.is_empty() { &[] } else { &body };
        let link = click.for_add(self.slides_in_catalog()?)?;
        let (target, proof, after) = match &source {
            // A click action needs the eager editor's hyperlink verbs.
            Editable::Source(package) if link.is_none() => {
                let mut edit = package
                    .edit_with_options(source_edit_options(native_snapshot_budget(limits)))
                    .map_err(super::write::native_error)?;
                // The source editor returns the native non-visual shape ID.
                // The unified path is ordinal, so the detached scene above
                // supplies the stable direct-shape position independently.
                let _ =
                    edit.add_text_box(Position::new(slide), text.as_deref().unwrap_or(""), bounds)?;
                if !body.is_empty() {
                    let _ = edit.set_text_properties(Position::new(slide), body)?;
                }
                let (target, patch) = edit.commit()?.into_parts();
                let after = self.adopt_source(target.clone())?;
                (Editable::Source(target), Proof::Source(patch), after)
            },
            Editable::Eager(_) | Editable::Source(_) => {
                let package = super::smartart::eager(&source, self.0.ooxml_limits)?;
                let mut edit = package
                    .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
                    .map_err(super::write::native_error)?;
                let _ =
                    edit.add_text_box(Position::new(slide), text.as_deref().unwrap_or(""), bounds)?;
                if !body.is_empty() {
                    let _ = edit.set_text_properties(Position::new(slide), body)?;
                }
                if let Some(link) = &link {
                    let _ = edit.set_shape_hyperlink(slide, scene_index.get(), Some(link))?;
                }
                let (target, patch) = edit.commit()?.into_parts();
                let target = Arc::new(target);
                let after = self.adopt_eager(Arc::clone(&target));
                (Editable::Eager(target), Proof::Eager(patch), after)
            },
        };
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("shape").with(Predicate::Index(Position::new(ordinal))),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                target,
                proof,
                1 + usize::from(!body.is_empty()) + usize::from(link.is_some()),
            ),
            path,
        ))
    }
}
