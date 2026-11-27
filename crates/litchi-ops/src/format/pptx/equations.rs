//! Equations: Office Math in slide text and OLE equation objects.
//!
//! `/slide[N]/shape[M]/equation[K]` is the K-th Office Math equation in the
//! text of a shape. `/slide[N]/equation[K]` is the K-th equation of the
//! slide: its Office Math equations in document order, then its OLE
//! equation objects (MathType `Equation.DSMT4`, Equation Editor 3.0
//! `Equation.3`) in document order, so every Office Math path stays the
//! same whatever OLE objects a slide holds. Each reads `latex`, its OMML or
//! MTEF converted to LaTeX (absent without the `equation` feature or when
//! the source does not convert); an Office Math equation also reads its raw
//! `omml` and its text, and an OLE object its `mtef` (the `Equation Native`
//! stream, base64), its `progId` and `error`, why its embedding cannot be
//! read. An OLE object's frame is also listed as `/slide[N]/ole[K]`, as PPT
//! lists its equation frames. Office Math equations are read from the
//! slide's own XML (PowerPoint writes them in a markup-compatibility choice
//! whose fallback, a picture, is what the shape scene reads).
//!
//! With the `equation` feature, `add(/slide[N], equation, {latex | omml,
//! x, y, size})` adds an equation text box in PowerPoint's form (the OMML
//! with a PNG picture fallback), `add(/slide[N]/shape[M], equation, {latex |
//! omml})` adds one at the end of a shape's text, and
//! `add(/slide[N], equation, {mtef | progId: "Equation.DSMT4", latex | omml,
//! x, y, size})` adds a MathType OLE object with a PNG preview. `set`
//! replaces an equation's `latex` or `omml`, or an OLE object's `latex`,
//! `omml` or `mtef` keeping its kind, and `remove` deletes it (an Office
//! Math equation with its shape when it was the shape's whole text; an OLE
//! object with its embedding, preview and legacy VML shape), through
//! `opened::Edit`'s equation verbs. A shape whose text mixes equations with
//! other text gets a text fallback with each equation in linear form. A
//! conversion that loses formatting is refused, so the caller can write
//! OMML or MTEF instead.

use std::borrow::Cow;
use std::sync::Arc;

#[cfg(feature = "equation")]
use litchi_core::Position;
use litchi_core::selector::Path;
#[cfg(feature = "equation")]
use litchi_core::selector::{Predicate, Step};
use litchi_pptx::shape::math::{Equation, OleEquation};

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

/// The ProgID a new OLE equation object takes.
const MATHTYPE_PROG_ID: &str = "Equation.DSMT4";

/// What an equation node is.
#[derive(Clone)]
pub(super) enum Content {
    /// An Office Math equation of slide text.
    Omml(Arc<Equation>),
    /// An OLE equation object.
    Ole(Arc<OleEquation>),
}

/// One equation and where it is addressed.
#[derive(Clone)]
pub(super) struct EquationItem {
    pub(super) slide: u32,
    /// The shape whose text holds it, when it is addressed through it.
    pub(super) shape: Option<TreeItem>,
    /// Its position in its scope (the shape's or the slide's equations).
    pub(super) index: u32,
    /// Its position among the slide's equations of its kind (Office Math
    /// or OLE).
    pub(super) ordinal: u32,
    pub(super) content: Content,
}

impl EquationItem {
    /// Whether it sits in the text of a shape (and is listed under it).
    pub(super) fn in_shape(&self) -> bool {
        matches!(&self.content, Content::Omml(equation) if equation.shape_id().is_some())
    }

    /// Its text: an Office Math equation's `m:t` text, an OLE object's
    /// LaTeX.
    pub(super) fn text(&self) -> Option<Cow<'_, str>> {
        match &self.content {
            Content::Omml(equation) => Some(Cow::Borrowed(equation.text())),
            Content::Ole(object) => ole_latex(object).map(Cow::Owned),
        }
    }
}

/// The value of the equation property at `index`.
pub(super) fn value(item: &EquationItem, index: usize) -> Option<Value<'_>> {
    let key = EquationKey::at(index)?;
    match &item.content {
        Content::Omml(equation) => match key {
            EquationKey::Latex => latex(equation).map(|latex| Value::Text(Cow::Owned(latex))),
            EquationKey::Omml => Some(Value::Text(Cow::Borrowed(equation.omml()))),
            EquationKey::Mtef | EquationKey::ProgId | EquationKey::Error => None,
        },
        Content::Ole(object) => match key {
            EquationKey::Latex => ole_latex(object).map(|latex| Value::Text(Cow::Owned(latex))),
            EquationKey::Omml => None,
            EquationKey::ProgId => Some(Value::Text(Cow::Borrowed(object.prog_id()))),
            EquationKey::Mtef => mtef(object).map(|mtef| Value::Text(Cow::Owned(mtef))),
            EquationKey::Error => unreadable(object).map(|reason| Value::Text(Cow::Owned(reason))),
        },
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

/// The LaTeX of an OLE equation; `None` without the `equation` feature and
/// for MTEF that cannot be read or does not convert.
fn ole_latex(object: &OleEquation) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        object.latex(&litchi_equation::Limits::DEFAULT).ok()
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = object;
        None
    }
}

/// An OLE equation's `Equation Native` stream as base64.
fn mtef(object: &OleEquation) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        use base64::Engine as _;
        object
            .equation_native()
            .ok()
            .map(|native| base64::engine::general_purpose::STANDARD.encode(native))
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = object;
        None
    }
}

/// Why an OLE equation's embedding cannot be read.
fn unreadable(object: &OleEquation) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        object
            .equation_native()
            .err()
            .map(|error| error.reason().to_owned())
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = object;
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

    /// The OLE equation objects of the slide at `position`, in document
    /// order.
    pub(super) fn slide_ole_equations(&self, position: usize) -> Result<Vec<OleEquation>> {
        self.finish(|| {
            if let Some(revision) = &self.0.revision {
                let presentation = revision.presentation()?;
                if position >= presentation.slide_count()? {
                    return Ok(Vec::new());
                }
                return Ok(presentation.slide(position)?.ole_equations()?);
            }
            match &self.0.owner {
                PptxOwner::Source(presentation) => {
                    self.check_owner()?;
                    let Some(slide) = presentation.slide(position) else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.ole_equations()?)
                },
                PptxOwner::Eager(package) => {
                    let Some(slide) = super::eager_slide(self, package, position)? else {
                        return Ok(Vec::new());
                    };
                    Ok(slide.ole_equations()?)
                },
            }
        })
    }
}

/// The equations of a slide (`shape: None`: its Office Math equations, then
/// its OLE equation objects) or the Office Math equations of one of its
/// shapes.
pub(super) fn items(file: &Pptx, slide: u32, shape: Option<&TreeItem>) -> Result<Vec<Item>> {
    let id = shape.and_then(|tree| tree.handle.view().id());
    if shape.is_some() && id.is_none() {
        return Ok(Vec::new());
    }
    let mut out = Vec::new();
    let position = |count: usize| u32::try_from(count).unwrap_or(u32::MAX);
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
            index: position(out.len()),
            ordinal: position(ordinal),
            content: Content::Omml(Arc::new(equation)),
        })));
    }
    if shape.is_none() {
        for (ordinal, object) in file
            .slide_ole_equations(slide as usize)?
            .into_iter()
            .enumerate()
        {
            out.push(Item::Equation(Box::new(EquationItem {
                slide,
                shape: None,
                index: position(out.len()),
                ordinal: position(ordinal),
                content: Content::Ole(Arc::new(object)),
            })));
        }
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

/// What an equation is written from.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the content is written with the equation feature")
)]
enum Source {
    Latex(String),
    Omml(String),
    Mtef(Vec<u8>),
}

impl Source {
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "the content is written with the equation feature")
    )]
    const fn key(&self) -> &'static str {
        match self {
            Self::Latex(_) => "latex",
            Self::Omml(_) => "omml",
            Self::Mtef(_) => "mtef",
        }
    }
}

/// What an `add` or `set` of an equation assigns.
#[derive(Default)]
struct Input {
    content: Option<Source>,
    /// Whether `progId` asked for an OLE equation object.
    ole: bool,
    x: Option<i64>,
    y: Option<i64>,
    size: Option<f64>,
}

fn decode_mtef(text: &str) -> Result<Vec<u8>> {
    use base64::Engine as _;
    base64::engine::general_purpose::STANDARD
        .decode(text.trim())
        .map_err(|_error| {
            error(
                Code::InvalidValue,
                "mtef is the base64 of an Equation Native stream",
            )
            .with_property("mtef")
        })
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
                (
                    EquationSetter::Latex | EquationSetter::Omml | EquationSetter::Mtef,
                    Some(Scalar::Text(text)),
                ) => {
                    bound(
                        LimitKind::StagedBytes,
                        text.len(),
                        limits.max_staged_bytes(),
                    )?;
                    let content = match setter {
                        EquationSetter::Latex => Source::Latex(text.to_owned()),
                        EquationSetter::Omml => Source::Omml(text.to_owned()),
                        _ => Source::Mtef(decode_mtef(text)?),
                    };
                    if input.content.replace(content).is_some() {
                        return Err(error(
                            Code::InvalidArgument,
                            "an equation takes one of latex, omml and mtef",
                        ));
                    }
                },
                (EquationSetter::ProgId, Some(Scalar::Text(text))) => {
                    if !text.trim().eq_ignore_ascii_case(MATHTYPE_PROG_ID) {
                        return Err(error(
                            Code::InvalidValue,
                            "a new equation object is a MathType object (Equation.DSMT4)",
                        )
                        .with_property("progId")
                        .with_suggestion("set progId to Equation.DSMT4, or omit it"));
                    }
                    input.ole = true;
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
    if matches!(input.content, Some(Source::Mtef(_))) {
        input.ole = true;
    }
    Ok(input)
}

/// The leaf source of admitted content.
#[cfg(feature = "equation")]
fn source(content: &Source) -> Result<litchi_pptx::shape::math::EquationSource<'_>> {
    use litchi_pptx::shape::math::EquationSource;
    match content {
        Source::Latex(latex) => Ok(EquationSource::Latex(latex)),
        Source::Omml(omml) => Ok(EquationSource::Omml(omml)),
        Source::Mtef(_) => Err(mtef_on_omml()),
    }
}

/// The leaf source of admitted content for an OLE equation object.
#[cfg(feature = "equation")]
fn ole_source(content: &Source) -> litchi_pptx::shape::math::OleEquationSource<'_> {
    use litchi_pptx::shape::math::OleEquationSource;
    match content {
        Source::Latex(latex) => OleEquationSource::Latex(latex),
        Source::Omml(omml) => OleEquationSource::Omml(omml),
        Source::Mtef(mtef) => OleEquationSource::Mtef(mtef),
    }
}

/// Refuse a LaTeX conversion that lost formatting: the caller decides
/// whether to write OMML (or MTEF, for an OLE object) instead.
#[cfg(feature = "equation")]
fn refuse_losses(losses: &[String], content: &Source, ole: bool) -> Result<()> {
    match losses.first() {
        None => Ok(()),
        Some(loss) if ole => Err(crate::Error::new(
            Code::InvalidValue,
            format!("the equation does not convert to MTEF without loss: {loss}"),
        )
        .with_format(litchi_core::FileFormat::Pptx)
        .with_property(content.key())
        .with_suggestion("write the equation as mtef")),
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

/// The refusal of MTEF for an Office Math equation.
fn mtef_on_omml() -> crate::Error {
    error(
        Code::UnsupportedProperty,
        "an Office Math equation has no MTEF; only OLE equation objects do",
    )
    .with_property("mtef")
    .with_suggestion("write the equation as latex or omml")
}

impl Pptx {
    /// `add(/slide[N], equation, {latex | omml | mtef, progId, x, y,
    /// size})` and `add(/slide[N]/shape[M], equation, {latex | omml})`.
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
            return Err(error(
                Code::MissingProperty,
                "an equation needs latex, omml or mtef",
            )
            .with_property("latex"));
        };
        if input.ole && shape.is_some() {
            return Err(error(
                Code::UnsupportedType,
                "an OLE equation object is added to a slide, not to a shape's text",
            )
            .with_suggestion("add the equation to /slide[N]"));
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = (content, slide, shape);
            Err(unavailable())
        }
        #[cfg(feature = "equation")]
        {
            if input.ole
                && let Source::Mtef(bytes) = content
                && litchi_equation::read_mtef(bytes, &litchi_equation::Limits::DEFAULT)
                    .ok()
                    .and_then(|read| read.info().map(|info| info.version))
                    .is_some_and(|version| version != 5)
                && node
                    .props()
                    .iter()
                    .any(|(key, _)| key.as_str().eq_ignore_ascii_case("progId"))
            {
                return Err(error(
                    Code::InvalidValue,
                    "a MathType object (Equation.DSMT4) holds MTEF 5",
                )
                .with_property("mtef")
                .with_suggestion("omit progId to add an Equation Editor 3.0 object"));
            }
            let source = self.editable_owner()?;
            let base = super::smartart::eager(&source, self.0.ooxml_limits)?;
            let mut edit = base
                .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
                .map_err(super::write::native_error)?;
            let origin = (input.x.unwrap_or(0), input.y.unwrap_or(0));
            let (written, offset) = if input.ole {
                let written = edit
                    .add_ole_equation(slide, ole_source(content), origin, input.size)
                    .map_err(|e| crate::Error::from(e).with_property(content.key()))?;
                (written, self.slide_equations(slide)?.len())
            } else {
                if matches!(content, Source::Mtef(_)) {
                    return Err(mtef_on_omml());
                }
                let written = match shape {
                    None => edit.add_equation(slide, self::source(content)?, origin, input.size),
                    Some(id) => {
                        if input.x.is_some() || input.y.is_some() || input.size.is_some() {
                            return Err(error(
                                Code::UnsupportedProperty,
                                "an equation added to a shape's text takes no position or size",
                            ));
                        }
                        edit.add_equation_to_shape(slide, id, self::source(content)?)
                    },
                }
                .map_err(|e| crate::Error::from(e).with_property(content.key()))?;
                (written, 0)
            };
            refuse_losses(&written.losses, content, input.ole)?;
            let (target, patch) = edit.commit()?.into_parts();
            let target = Arc::new(target);
            let after = self.adopt_eager(Arc::clone(&target));
            let path = Path::from_steps([
                Step::new("slide").with(Predicate::Index(Position::new(slide))),
                Step::new("equation").with(Predicate::Index(Position::new(
                    written.index.saturating_add(offset),
                ))),
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

    /// A `set` of an equation's `latex`, `omml` or (on an OLE object)
    /// `mtef`; `None` for other targets.
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
        let ole = matches!(item.content, Content::Ole(_));
        if !ole && matches!(content, Source::Mtef(_)) {
            return Err(mtef_on_omml());
        }
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
            let (slide, ordinal) = (item.slide as usize, item.ordinal as usize);
            let written = if ole {
                edit.set_ole_equation(slide, ordinal, ole_source(content))
            } else {
                edit.set_equation(slide, ordinal, self::source(content)?)
            }
            .map_err(|e| crate::Error::from(e).with_property(content.key()))?;
            if let Some(written) = &written {
                refuse_losses(&written.losses, content, ole)?;
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

    /// A `remove` of an equation with its shape and picture, or of an OLE
    /// equation object with its parts; `None` for other targets.
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
            let (slide, ordinal) = (item.slide as usize, item.ordinal as usize);
            match item.content {
                Content::Omml(_) => edit.remove_equation(slide, ordinal)?,
                Content::Ole(_) => edit.remove_ole_equation(slide, ordinal)?,
            };
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
