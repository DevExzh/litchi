//! Office Math equations.
//!
//! `/slide[N]/shape[M]/equation[K]` is the K-th equation in the text of a
//! shape; `/slide[N]/equation[K]` the K-th of the slide. Each reads
//! `latex`, its OMML converted to LaTeX (absent without the `equation`
//! feature or when the OMML does not convert), the raw `omml`, and its
//! text. Equations are read from the slide's own XML (PowerPoint writes them in a
//! markup-compatibility choice whose fallback, a picture, is what the
//! shape scene reads).
//!
//! With the `equation` feature, `add(/slide[N], equation, {latex | omml,
//! x, y, size})` adds an equation text box in PowerPoint's form (the OMML
//! with a PNG picture fallback), `add(/slide[N]/shape[M], equation, {latex |
//! omml})` adds one at the end of a shape's text, `set` replaces an
//! equation's `latex` or `omml`, and `remove` deletes it (with its shape
//! when it was the shape's whole text), through `opened::Edit`'s
//! `add_equation`, `add_equation_to_shape`, `set_equation` and
//! `remove_equation`. A shape whose text mixes equations with other text
//! gets a text fallback with each equation in linear form. LaTeX that
//! loses formatting in the conversion is refused, so the caller can write
//! OMML instead.

use std::borrow::Cow;
use std::sync::Arc;

#[cfg(feature = "equation")]
use litchi_core::Position;
use litchi_core::selector::Path;
#[cfg(feature = "equation")]
use litchi_core::selector::{Predicate, Step};
use litchi_pptx::shape::math::Equation;

#[cfg(feature = "equation")]
use super::Editable;
use super::node::{Item, TreeItem};
use super::spec::{EQUATION_WRITE_TABLE, EquationKey, EquationSetter};
use super::structure::{bound, error};
use super::write::Commit;
#[cfg(feature = "equation")]
use super::write::{Proof, native_snapshot_budget, opened_edit_options};
use super::{Nav, Pptx, PptxOwner};
use crate::node::Value;
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Result};

/// One equation and where it is addressed.
#[derive(Clone)]
pub(super) struct EquationItem {
    pub(super) slide: u32,
    /// The shape whose text holds it, when it is addressed through it.
    pub(super) shape: Option<TreeItem>,
    /// Its position in its scope (the shape's or the slide's equations).
    pub(super) index: u32,
    /// Its position among the slide's equations.
    pub(super) ordinal: u32,
    pub(super) equation: Arc<Equation>,
}

/// The value of the equation property at `index`.
pub(super) fn value(item: &EquationItem, index: usize) -> Option<Value<'_>> {
    match EquationKey::at(index)? {
        EquationKey::Latex => latex(&item.equation).map(|latex| Value::Text(Cow::Owned(latex))),
        EquationKey::Omml => Some(Value::Text(Cow::Borrowed(item.equation.omml()))),
    }
}

/// The LaTeX of an equation; `None` without the `equation` feature and for
/// OMML that does not convert.
fn latex(equation: &Equation) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        equation.latex().ok()
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = equation;
        None
    }
}

impl Pptx {
    /// The equations of the slide at `position`, in document order.
    pub(super) fn slide_equations(&self, position: usize) -> Result<Vec<Equation>> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                let presentation = revision.presentation()?;
                if position >= presentation.slide_count()? {
                    return Ok(Vec::new());
                }
                return Ok(presentation.slide(position)?.equations()?);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.equations()?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.equations()?)
                },
            }
        })
    }
}

/// The equations of a slide (`shape: None`) or of one of its shapes.
pub(super) fn items(file: &Pptx, slide: u32, shape: Option<&TreeItem>) -> Result<Vec<Item>> {
    let id = shape.and_then(|tree| tree.handle.view().id());
    if shape.is_some() && id.is_none() {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    for (ordinal, equation) in file
        .slide_equations(slide as usize)?
        .into_iter()
        .enumerate()
    {
        if shape.is_some() && equation.shape_id() != id {
            continue;
        }
        out.push(Item::Equation(Box::new(EquationItem {
            slide,
            shape: shape.cloned(),
            index: u32::try_from(out.len()).unwrap_or(u32::MAX),
            ordinal: u32::try_from(ordinal).unwrap_or(u32::MAX),
            equation: Arc::new(equation),
        })));
    }
    Ok(out)
}

/// The equation a `set` or `remove` path names; `None` for other paths.
fn target(file: &Pptx, target: &TargetArg<'_>) -> Result<Option<EquationItem>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    if !path.path().steps().last().is_some_and(|step| {
        step.name().eq_ignore_ascii_case("equation") || step.name().eq_ignore_ascii_case("math")
    }) {
        return Ok(None);
    }
    let node = crate::path::resolve::get(&Nav::new(file), path)?;
    Ok(match node.item() {
        Item::Equation(item) => Some(EquationItem::clone(item)),
        _ => None,
    })
}

/// What an `add` or `set` of an equation assigns.
#[derive(Default)]
struct Input {
    content: Option<(EquationSetter, String)>,
    x: Option<i64>,
    y: Option<i64>,
    size: Option<f64>,
}

fn input(props: &crate::Props<'_>, operation: Operation, limits: &SemanticLimits) -> Result<Input> {
    let mut input = Input::default();
    let mut assigned = 0_u32;
    for (key, value) in props.iter() {
        let (index, entry) = EQUATION_WRITE_TABLE
            .resolve(key.as_str())
            .map_err(|e| e.with_property(key.as_str().to_owned()))?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "an equation property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry.prepare(operation, value, |setter, scalar| {
            match (setter, scalar) {
                (EquationSetter::Latex | EquationSetter::Omml, Some(Scalar::Text(text))) => {
                    bound(
                        LimitKind::StagedBytes,
                        text.len(),
                        limits.max_staged_bytes(),
                    )?;
                    if input.content.replace((setter, text.to_owned())).is_some() {
                        return Err(error(
                            Code::InvalidArgument,
                            "an equation takes one of latex and omml",
                        ));
                    }
                },
                (EquationSetter::X, Some(Scalar::Length(length))) => input.x = Some(length.emu()),
                (EquationSetter::Y, Some(Scalar::Length(length))) => input.y = Some(length.emu()),
                (EquationSetter::Size, Some(Scalar::Length(length))) => {
                    input.size = Some(length.points());
                },
                _ => unreachable!("equation property coercion"),
            }
            Ok(())
        })?;
    }
    Ok(input)
}

/// The leaf source of an admitted `latex` or `omml`.
#[cfg(feature = "equation")]
fn source(content: &(EquationSetter, String)) -> litchi_pptx::shape::math::EquationSource<'_> {
    use litchi_pptx::shape::math::EquationSource;
    match content {
        (EquationSetter::Omml, omml) => EquationSource::Omml(omml),
        (_, latex) => EquationSource::Latex(latex),
    }
}

/// Refuse a LaTeX conversion that lost formatting: the caller decides
/// whether to write OMML instead.
#[cfg(feature = "equation")]
fn refuse_losses(losses: &[String]) -> Result<()> {
    match losses.first() {
        None => Ok(()),
        Some(loss) => Err(crate::Error::new(
            Code::InvalidValue,
            format!("the LaTeX equation does not convert to OMML without loss: {loss}"),
        )
        .with_format(litchi_core::FileFormat::Pptx)
        .with_property("latex")
        .with_suggestion("write the equation as omml")),
    }
}

#[cfg(not(feature = "equation"))]
fn unavailable() -> crate::Error {
    error(
        Code::UnsupportedCapability,
        "writing equations needs the equation feature",
    )
    .with_suggestion("enable the litchi equation feature")
}

impl Pptx {
    /// `add(/slide[N], equation, {latex | omml, x, y, size})`.
    pub(super) fn add_equation(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let parent = crate::path::resolve::get(&Nav::new(self), &parent)?;
        // A slide takes a new text box, a shape the equation at the end of
        // its text.
        let (slide, shape) = match parent.item() {
            Item::Slide(slide) => (slide.position as usize, None),
            Item::Tree(tree) if tree.kind == super::spec::Kind::Shape => {
                let (Some(slide), Some(id)) = (tree.container.slide(), tree.handle.view().id())
                else {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "equations are added to the shapes of slides only",
                    ));
                };
                (slide as usize, Some(id))
            },
            _ => {
                return Err(error(
                    Code::UnsupportedType,
                    "equations are added to a slide, as a new text box, or to a shape's text",
                )
                .with_suggestion("add the equation to /slide[N] or /slide[N]/shape[M]"));
            },
        };
        drop(parent);
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "equation insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let input = input(node.props(), Operation::Add, limits)?;
        let Some(content) = input.content.as_ref() else {
            return Err(
                error(Code::MissingProperty, "an equation needs latex or omml")
                    .with_property("latex"),
            );
        };
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, slide, shape);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let source = self.editable_owner()?;
            let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
            let mut edit = base
                .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
                .map_err(super::write::native_error)?;
            let written = match shape {
                None => edit.add_equation(
                    slide,
                    self::source(content),
                    (input.x.unwrap_or(0), input.y.unwrap_or(0)),
                    input.size,
                ),
                Some(id) => {
                    if input.x.is_some() || input.y.is_some() || input.size.is_some() {
                        return Err(error(
                            Code::UnsupportedProperty,
                            "an equation added to a shape's text takes no position or size",
                        ));
                    }
                    edit.add_equation_to_shape(slide, id, self::source(content))
                },
            }
            .map_err(|e| crate::Error::from(e).with_property(content_key(content)))?;
            refuse_losses(&written.losses)?;
            let (target, patch) = edit.commit()?.into_parts();
            let target = Arc::new(target);
            let after = self.adopt_eager(Arc::clone(&target));
            let path = Path::from_steps([
                Step::new("slide").with(Predicate::Index(Position::new(slide))),
                Step::new("equation").with(Predicate::Index(Position::new(written.index))),
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

    /// A `set` of an equation's `latex` or `omml`; `None` for other
    /// targets.
    pub(super) fn try_set_equation(
        &self,
        target: &TargetArg<'_>,
        props: &crate::Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self::target(self, target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let input = input(props, Operation::Set, limits)?;
        let Some(content) = input.content.as_ref() else {
            return Err(error(
                Code::InvalidArgument,
                "no equation property was assigned",
            ));
        };
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, item.slide, item.ordinal);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let source = self.editable_owner_for_write()?;
            let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
            let mut edit = base
                .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
                .map_err(super::write::native_error)?;
            let written = edit
                .set_equation(
                    item.slide as usize,
                    item.ordinal as usize,
                    self::source(content),
                )
                .map_err(|e| crate::Error::from(e).with_property(content_key(content)))?;
            if let Some(written) = &written {
                refuse_losses(&written.losses)?;
            }
            let (target, patch) = edit.commit()?.into_parts();
            let target = Arc::new(target);
            let after = if patch.is_noop() {
                self.clone()
            } else {
                self.adopt_eager(Arc::clone(&target))
            };
            Ok(Some(Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                usize::from(written.is_some()),
            )))
        }
    }

    /// A `remove` of an equation with its shape and picture; `None` for
    /// other targets.
    pub(super) fn try_remove_equation(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(item) = self::target(self, target)? else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        #[cfg(not(feature = "equation"))]
        {
            let _ = (item.slide, item.ordinal);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            let source = self.editable_owner()?;
            let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
            let mut edit = base
                .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
                .map_err(super::write::native_error)?;
            edit.remove_equation(item.slide as usize, item.ordinal as usize)?;
            let (target, patch) = edit.commit()?.into_parts();
            let target = Arc::new(target);
            let after = self.adopt_eager(Arc::clone(&target));
            Ok(Some(Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            )))
        }
    }
}

#[cfg(feature = "equation")]
fn content_key(content: &(EquationSetter, String)) -> &'static str {
    match content.0 {
        EquationSetter::Omml => "omml",
        _ => "latex",
    }
}
