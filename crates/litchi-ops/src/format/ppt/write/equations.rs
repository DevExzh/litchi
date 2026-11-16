//! Unified writes of PPT equations.
//!
//! `add(/slide[N], equation, {latex | omml | mtef, x, y, size})` adds a
//! MathType equation object (an Equation Editor 3.0 one for MTEF 3 input)
//! with its PNG preview at (`x`, `y`), sized for text of `size` (24 pt by
//! default); `set` replaces an equation's `latex`, `omml` or `mtef`, keeping
//! the object's kind; `remove` deletes the equation with its object,
//! storage and preview. Each call is one native equation edit
//! (`litchi_ppt::Package::edit_equations`), which needs the `equation`
//! feature. A LaTeX or OMML equation that MTEF cannot hold without loss is
//! refused, so the caller can write MTEF instead.
//!
//! PPT edits are append-only saves: a replaced or removed equation's old
//! storage stays in the file, reachable through earlier persist
//! directories. A removal is refused while an animation, a build list or a
//! connector of the slide still names the equation's frame.

use litchi_core::FileFormat;

use super::super::Ppt;
use super::super::elements::Element;
use super::super::node::Node;
use super::super::spec::{EQUATION_WRITE_TABLE, EquationSetter, Kind};
use super::{WriteCommit, bound, master_from_emu};
use crate::error::{Code, Error, Result};
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::Operation;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The equation content of one write.
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the content is written with the equation feature")
)]
enum Content {
    Latex(String),
    Omml(String),
    Mtef(Vec<u8>),
}

#[cfg(feature = "equation")]
impl Content {
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
    content: Option<Content>,
    x: Option<i32>,
    y: Option<i32>,
    size: Option<f64>,
}

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Ppt)
}

fn input(props: &Props<'_>, operation: Operation, limits: &SemanticLimits) -> Result<Input> {
    let mut input = Input::default();
    let mut assigned = 0_u32;
    for (key, value) in props.iter() {
        let (index, entry) = EQUATION_WRITE_TABLE.resolve(key.as_str()).map_err(|e| {
            e.with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u32 << index;
        if assigned & bit != 0 {
            return Err(error(
                Code::InvalidArgument,
                "an equation property was assigned more than once",
            )
            .with_property(entry.spec.name));
        }
        assigned |= bit;
        entry
            .prepare(operation, value, |setter, scalar| {
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
                            EquationSetter::Latex => Content::Latex(text.to_owned()),
                            EquationSetter::Omml => Content::Omml(text.to_owned()),
                            _ => Content::Mtef(decode_mtef(text)?),
                        };
                        if input.content.replace(content).is_some() {
                            return Err(error(
                                Code::InvalidArgument,
                                "an equation takes one of latex, omml and mtef",
                            ));
                        }
                    },
                    (EquationSetter::X | EquationSetter::Y, Some(Scalar::Length(length))) => {
                        let master = master_from_emu(length.emu(), entry.spec.name)?;
                        let master = i32::try_from(master).map_err(|_error| {
                            error(
                                Code::InvalidValue,
                                "the equation position is outside the native range",
                            )
                            .with_property(entry.spec.name)
                        })?;
                        if setter == EquationSetter::X {
                            input.x = Some(master);
                        } else {
                            input.y = Some(master);
                        }
                    },
                    (EquationSetter::Size, Some(Scalar::Length(length))) => {
                        input.size = Some(length.points());
                    },
                    _ => unreachable!("PPT equation coercion returned an unexpected scalar"),
                }
                Ok(())
            })
            .map_err(|e| e.with_format(FileFormat::Ppt))?;
    }
    Ok(input)
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

/// The slide position and frame shape id of one equation node.
fn target(node: &Node<'_>) -> Result<(usize, u32)> {
    match node.element_item() {
        Some(Element::Equation {
            slide, equation, ..
        }) => Ok((*slide, equation.shape_id())),
        _ => Err(error(
            Code::InternalError,
            "the PPT equation node has no equation",
        )),
    }
}

#[cfg(not(feature = "equation"))]
fn unavailable() -> Error {
    error(
        Code::UnsupportedCapability,
        "writing equations needs the equation feature",
    )
    .with_suggestion("enable the litchi equation feature")
}

/// Replace the content of every selected equation in one native edit.
pub(super) fn set_equations(
    file: &Ppt,
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    bound(
        LimitKind::ExpandedTargets,
        nodes.len(),
        limits.max_expanded_targets(),
    )?;
    let input = input(props, Operation::Set, limits)?;
    let Some(content) = input.content else {
        return Err(error(
            Code::InvalidArgument,
            "no equation property was assigned",
        ));
    };
    let targets = nodes.iter().map(target).collect::<Result<Vec<_>>>()?;
    #[cfg(not(feature = "equation"))]
    {
        let _ = (file, content, targets);
        Err(unavailable())
    }
    #[cfg(feature = "equation")]
    {
        native::publish(file, limits, targets.len(), |edit| {
            for (slide, shape) in &targets {
                let written = edit
                    .set(
                        litchi_core::Position::new(*slide),
                        *shape,
                        native::source(&content),
                        &litchi_equation::Limits::DEFAULT,
                    )
                    .map_err(|e| Error::from(e).with_property(content.key()))?;
                native::refuse_losses(written.losses(), &content)?;
            }
            Ok(())
        })
    }
}

/// Remove every selected equation in one native edit.
pub(super) fn remove_equations(
    file: &Ppt,
    nodes: &[Node<'_>],
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    bound(
        LimitKind::ExpandedTargets,
        nodes.len(),
        limits.max_expanded_targets(),
    )?;
    let targets = nodes.iter().map(target).collect::<Result<Vec<_>>>()?;
    #[cfg(not(feature = "equation"))]
    {
        let _ = (file, targets);
        Err(unavailable())
    }
    #[cfg(feature = "equation")]
    {
        native::publish(file, limits, targets.len(), |edit| {
            for (slide, shape) in &targets {
                edit.remove(litchi_core::Position::new(*slide), *shape)?;
            }
            Ok(())
        })
    }
}

/// Add one equation to a slide.
pub(super) fn add_equation(
    file: &Ppt,
    parent: &Node<'_>,
    node: &crate::NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, crate::path::Path<'static>)> {
    if parent.element() != Kind::Slide {
        return Err(
            error(Code::UnsupportedType, "PPT equations are added to a slide")
                .with_suggestion("add the equation to /slide[N]"),
        );
    }
    if !matches!(node.placement(), None | Some(Placement::End)) {
        return Err(error(
            Code::UnsupportedCapability,
            "PPT equation insertion supports end placement",
        )
        .with_suggestion("omit placement or use Placement::End"));
    }
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let input = input(node.props(), Operation::Add, limits)?;
    let Some(content) = input.content else {
        return Err(error(
            Code::MissingProperty,
            "an equation needs latex, omml or mtef",
        )
        .with_property("latex"));
    };
    let slide = parent
        .slide_position()
        .ok_or_else(|| error(Code::InternalError, "PPT slide parent has no position"))?;
    #[cfg(not(feature = "equation"))]
    {
        let _ = (file, content, slide, input.x, input.y, input.size);
        Err(unavailable())
    }
    #[cfg(feature = "equation")]
    {
        let mut shape = 0;
        let commit = native::publish(file, limits, 1, |edit| {
            let written = edit
                .add(
                    litchi_core::Position::new(slide),
                    input.x.unwrap_or(0),
                    input.y.unwrap_or(0),
                    input.size.unwrap_or(native::DEFAULT_SIZE_PT),
                    native::source(&content),
                    &litchi_equation::Limits::DEFAULT,
                )
                .map_err(|e| Error::from(e).with_property(content.key()))?;
            native::refuse_losses(written.losses(), &content)?;
            shape = written.shape_id();
            Ok(())
        })?;
        let index = commit
            .after
            .reader()
            .slide_equations(slide)?
            .iter()
            .position(|equation| equation.shape_id() == shape)
            .ok_or_else(|| {
                error(
                    Code::InternalError,
                    "the added PPT equation did not read back",
                )
            })?;
        Ok((
            commit,
            crate::path::Path::from_steps([
                crate::path::Step::new("slide").with(crate::path::Predicate::Index(
                    litchi_core::Position::new(slide),
                )),
                crate::path::Step::new("equation").with(crate::path::Predicate::Index(
                    litchi_core::Position::new(index),
                )),
            ]),
        ))
    }
}

#[cfg(feature = "equation")]
mod native {
    use std::mem::size_of;

    use litchi_core::FileFormat;
    use litchi_ppt::equation::Source;

    use super::super::{SidePatch, StepPatch, WriteCommit};
    use super::{Content, Error};
    use crate::error::{Code, Result};
    use crate::format::ppt::Ppt;
    use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

    /// The text size a new equation's frame is sized for.
    pub(super) const DEFAULT_SIZE_PT: f64 = 24.0;

    pub(super) fn source(content: &Content) -> Source<'_> {
        match content {
            Content::Latex(latex) => Source::Latex(latex),
            Content::Omml(omml) => Source::Omml(omml),
            Content::Mtef(mtef) => Source::Mtef(mtef),
        }
    }

    /// Refuse a conversion that lost part of the equation.
    pub(super) fn refuse_losses(losses: &[String], content: &Content) -> Result<()> {
        match losses.first() {
            None => Ok(()),
            Some(loss) => Err(Error::new(
                Code::InvalidValue,
                format!("the equation does not convert to MTEF without loss: {loss}"),
            )
            .with_format(FileFormat::Ppt)
            .with_property(content.key())
            .with_suggestion("write the equation as mtef")),
        }
    }

    /// Stage `operations` equation operations in one native edit and
    /// publish it.
    pub(super) fn publish(
        file: &Ppt,
        limits: &SemanticLimits,
        operations: usize,
        stage: impl FnOnce(&mut litchi_ppt::equation::Edit) -> Result<()>,
    ) -> Result<WriteCommit> {
        if operations > limits.max_operations() {
            return Err(TransactionError::Limit {
                kind: LimitKind::Operations,
                observed: operations,
                maximum: limits.max_operations(),
            }
            .into());
        }
        file.preflight_snapshot_bytes(
            limits.max_snapshot_bytes(),
            size_of::<litchi_ppt::equation::Edit>(),
        )?;
        let package = file.editable_package_for_write()?;
        let mut edit = package.edit_equations()?;
        stage(&mut edit)?;
        let commit = edit.commit()?;
        let changed = commit.diagnostics().changed();
        if changed && file.signature() == crate::format::signature::Signature::Signed {
            return Err(crate::format::signature::refusal(FileFormat::Ppt));
        }
        let (after, _patch) = commit.into_parts();
        let (patch, snapshot_bytes) = if changed {
            let before = package.to_bytes()?;
            let after = after.to_bytes()?;
            let bytes = crate::format::ole_metadata::snapshot_bytes(&before, &after);
            (
                StepPatch::Metadata(SidePatch::Equation { before, after }),
                bytes,
            )
        } else {
            // An unchanged edit keeps the original package's identity, as an
            // empty native edit proves it.
            let original = file.original_package_for_noop()?;
            let (_, patch) = original
                .edit()
                .map_err(|e| Error::from(e).with_format(FileFormat::Ppt))?
                .commit()
                .map_err(|e| Error::from(e).with_format(FileFormat::Ppt))?
                .into_parts();
            let bytes = patch
                .before()
                .len()
                .saturating_add(patch.after().len())
                .saturating_add(size_of::<litchi_ppt::Patch>());
            (StepPatch::Primary(patch), bytes)
        };
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(TransactionError::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: snapshot_bytes,
                maximum: limits.max_snapshot_bytes(),
            }
            .into());
        }
        file.check_source_for_write()?;
        let after = if changed {
            file.with_package(after)?
        } else {
            file.clone()
        };
        Ok(WriteCommit {
            before: file.clone(),
            after,
            patch,
            operations,
            snapshot_bytes,
        })
    }
}
