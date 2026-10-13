//! Typed text-box additions through the native immutable edit.
use super::node::Item;
use super::spec::{Kind, SHAPE_WRITE_TABLE, ShapeSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, validate_text};
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
                    setter => {
                        let (index, name) = match setter {
                            ShapeSetter::X => (0, "x"),
                            ShapeSetter::Y => (1, "y"),
                            ShapeSetter::Width => (2, "width"),
                            ShapeSetter::Height => (3, "height"),
                            ShapeSetter::Text | ShapeSetter::Format(_) => {
                                unreachable!("text and formatting are not add setters")
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
        let (target, proof, after) = match &source {
            Editable::Source(package) => {
                let mut edit = package.edit()?;
                // The source editor returns the native non-visual shape ID.
                // The unified path is ordinal, so the detached scene above
                // supplies the stable direct-shape position independently.
                let _ =
                    edit.add_text_box(Position::new(slide), text.as_deref().unwrap_or(""), bounds)?;
                let (target, patch) = edit.commit()?.into_parts();
                let after = self.adopt_source(target.clone())?;
                (Editable::Source(target), Proof::Source(patch), after)
            },
            Editable::Eager(package) => {
                let mut edit = package.edit()?;
                let _ =
                    edit.add_text_box(Position::new(slide), text.as_deref().unwrap_or(""), bounds)?;
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
            Commit {
                after,
                source,
                target,
                proof,
                operations: 1,
            },
            path,
        ))
    }
}
