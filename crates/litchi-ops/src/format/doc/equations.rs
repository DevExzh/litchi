//! DOC equations: the MathType and Equation Editor objects a paragraph
//! shows inline, as `equation[M]` nodes of the paragraph.
//!
//! An equation is the `EMBED` field of an OLE object whose `ObjectPool`
//! storage holds an `Equation Native` stream. Its node reads `mode` (always
//! `inline`), `latex` (the MTEF converted through `litchi-equation`) and
//! `mtef` (the whole `Equation Native` stream, its 28-byte header and the
//! MTEF, in base64). `add` inserts a new MathType object from `latex`,
//! `omml` or `mtef`; `set` replaces an object's content in the object's own
//! kind (MTEF 3 stays MTEF 3); `remove` deletes the field with its storage
//! and preview picture. Writes need the `equation` feature, which renders
//! the preview picture Word shows for every equation.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::selector::Path;
#[cfg(feature = "equation")]
use litchi_core::selector::{Predicate, Step};
use litchi_core::{FileFormat, Position};

use super::node::Item;
use super::spec::Kind;
use super::{Doc, Nav, Node, WriteCommit};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::node::Value;
use crate::placement::Placement;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};

/// One equation object of a paragraph.
#[derive(Debug)]
pub(crate) struct Equation {
    /// The CP of the field begin, for a main-story equation whose field
    /// the field table holds; the leaf edit address.
    pub(crate) start_cp: Option<u32>,
    /// The `ObjectPool` storage number.
    pub(crate) storage_id: u32,
    /// The MTEF as LaTeX, when it decodes.
    latex: Option<Arc<str>>,
    /// The size of the object's text, in points: the size a new preview
    /// is drawn at.
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "previews are drawn by the equation feature")
    )]
    font_size_pt: f64,
}

/// The size Word gives text that names none, in points.
const DEFAULT_FONT_SIZE_PT: f64 = 10.0;

/// The size of `run`'s text in points.
fn font_size_pt(run: Option<&litchi_doc::Run>) -> f64 {
    run.and_then(litchi_doc::Run::font_size)
        .filter(|size| *size > 0)
        .map_or(DEFAULT_FONT_SIZE_PT, |size| f64::from(size) / 2.0)
}

/// The equation objects shown in `paragraph`, in order: the runs that are
/// an OLE object's field separator and whose storage has an
/// `Equation Native` stream.
pub(super) fn equations(doc: &Doc, paragraph: &litchi_doc::Paragraph) -> Result<Vec<Equation>> {
    let reader = doc.reader()?;
    let mut found = Vec::new();
    let mut fields = None;
    for run in paragraph.runs()? {
        let properties = run.properties();
        let Some(storage_id) = properties.pic_offset.filter(|_| properties.is_ole2) else {
            continue;
        };
        if reader.equation_native(storage_id).is_none() {
            continue;
        }
        if fields.is_none() {
            fields = Some(super::marks::fields(doc)?);
        }
        let separator = run.cp_range().map(|range| range.start());
        let start_cp = fields.as_ref().and_then(|fields| {
            fields
                .iter()
                .find(|field| separator.is_some() && field.field.separator_cp == separator)
                .map(|field| field.field.start_cp)
        });
        found.push(Equation {
            start_cp,
            storage_id,
            latex: run.mtef_equation_latex().map(Arc::from),
            font_size_pt: font_size_pt(Some(&run)),
        });
    }
    Ok(found)
}

/// The equation nodes of a paragraph node.
pub(super) fn equation_nodes<'f>(
    parent: &Node<'f>,
    paragraph: &litchi_doc::Paragraph,
) -> Result<Vec<Node<'f>>> {
    Ok(equations(parent.doc, paragraph)?
        .into_iter()
        .enumerate()
        .map(|(index, equation)| {
            parent.child(
                Kind::Equation,
                Item::Equation(Arc::new(equation)),
                Position::new(index),
            )
        })
        .collect())
}

/// The value at `index` of the equation table.
pub(super) fn equation_value<'a>(
    doc: &'a Doc,
    equation: &'a Equation,
    index: usize,
) -> Option<Value<'a>> {
    match index {
        0 => Some(Value::Token("inline")),
        1 => equation
            .latex
            .as_deref()
            .map(|latex| Value::Text(Cow::Borrowed(latex))),
        2 => mtef_base64(doc, equation.storage_id).map(|text| Value::Text(Cow::Owned(text))),
        _ => None,
    }
}

#[cfg(feature = "equation")]
fn mtef_base64(doc: &Doc, storage_id: u32) -> Option<String> {
    use base64::Engine as _;
    let native = doc.reader().ok()?.equation_native(storage_id)?;
    Some(base64::engine::general_purpose::STANDARD.encode(native))
}

#[cfg(not(feature = "equation"))]
fn mtef_base64(_doc: &Doc, _storage_id: u32) -> Option<String> {
    None
}

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Doc)
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(crate::transaction::Error::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

/// The content an `add` or `set` writes: one of `latex`, `omml` and
/// `mtef`.
enum Source<'a> {
    Latex(&'a str),
    Omml(&'a str),
    Mtef(&'a str),
}

impl Source<'_> {
    const fn key(&self) -> &'static str {
        match self {
            Self::Latex(_) => "latex",
            Self::Omml(_) => "omml",
            Self::Mtef(_) => "mtef",
        }
    }
}

fn source<'a>(props: &'a Props<'_>, limits: &SemanticLimits) -> Result<Source<'a>> {
    let mut content = None;
    for (key, value) in props.iter() {
        let text = || {
            value.as_text().ok_or_else(|| {
                error(
                    Code::InvalidValue,
                    "the DOC equation property requires text",
                )
                .with_property(key.as_str().to_owned())
            })
        };
        let source = match key.as_str() {
            "latex" => Source::Latex(text()?),
            "omml" => Source::Omml(text()?),
            "mtef" => Source::Mtef(text()?),
            other => {
                return Err(error(
                    Code::UnsupportedProperty,
                    "DOC equations write `latex`, `omml` or `mtef`",
                )
                .with_property(other.to_owned()));
            },
        };
        let (Source::Latex(text) | Source::Omml(text) | Source::Mtef(text)) = source;
        if text.trim().is_empty() {
            return Err(
                error(Code::InvalidValue, "the DOC equation content is empty")
                    .with_property(source.key()),
            );
        }
        limit(
            LimitKind::StagedBytes,
            text.len(),
            limits.max_staged_bytes(),
        )?;
        if content.replace(source).is_some() {
            return Err(error(
                Code::InvalidArgument,
                "a DOC equation takes one of latex, omml and mtef",
            ));
        }
    }
    content.ok_or_else(|| {
        error(
            Code::InvalidArgument,
            "a DOC equation write needs latex, omml or mtef",
        )
    })
}

/// The leaf content of `source` for an object of `target`'s kind, refusing
/// a conversion that lost formatting.
#[cfg(feature = "equation")]
fn content(
    source: &Source<'_>,
    target: &litchi_equation::MtefTarget,
    font_size_pt: f64,
) -> Result<litchi_doc::EquationContent> {
    use base64::Engine as _;
    let limits = litchi_equation::Limits::DEFAULT;
    // The converter's own message says what is wrong with the input.
    let invalid = |error: litchi_doc::Error| -> Error {
        let code = match &error {
            litchi_doc::Error::Equation(
                litchi_equation::Error::InputTooLarge { .. }
                | litchi_equation::Error::LimitExceeded { .. }
                | litchi_equation::Error::Render(_),
            ) => Code::LimitExceeded,
            litchi_doc::Error::Equation(litchi_equation::Error::Internal(_)) => Code::InternalError,
            _ => Code::InvalidValue,
        };
        Error::new(
            code,
            format!(
                "the DOC equation `{}` does not convert: {error}",
                source.key()
            ),
        )
        .with_format(FileFormat::Doc)
        .with_property(source.key())
    };
    let (content, losses) = match source {
        Source::Latex(latex) => {
            litchi_doc::EquationContent::from_latex(latex, target, font_size_pt, &limits)
                .map_err(invalid)?
        },
        Source::Omml(omml) => {
            litchi_doc::EquationContent::from_omml(omml, target, font_size_pt, &limits)
                .map_err(invalid)?
        },
        Source::Mtef(text) => {
            let bytes = base64::engine::general_purpose::STANDARD
                .decode(text.trim())
                .map_err(|_| {
                    error(Code::InvalidValue, "the DOC equation `mtef` is not base64")
                        .with_property("mtef")
                })?;
            let content = litchi_doc::EquationContent::from_mtef(bytes, font_size_pt, &limits)
                .map_err(invalid)?;
            (content, Vec::new())
        },
    };
    if let Some(loss) = losses.first() {
        return Err(Error::new(
            Code::InvalidValue,
            format!("the equation does not convert to MTEF without loss: {loss}"),
        )
        .with_format(FileFormat::Doc)
        .with_property(source.key())
        .with_suggestion("write the equation as mtef"));
    }
    Ok(content)
}

#[cfg(not(feature = "equation"))]
fn unavailable() -> Error {
    error(
        Code::UnsupportedCapability,
        "writing DOC equations needs the equation feature",
    )
    .with_suggestion("enable the litchi equation feature")
}

/// Whether the last step of `path` may name an equation.
fn names_equation(path: &crate::path::PathArg<'_>) -> bool {
    path.path().last().is_some_and(|step| {
        ["equation", "math", "omath"]
            .iter()
            .any(|name| step.name().eq_ignore_ascii_case(name))
    })
}

/// The equation node `target` names, if it names one.
fn target_equation<'f>(doc: &'f Doc, target: &TargetArg<'_>) -> Result<Option<Node<'f>>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    if !names_equation(path) {
        return Ok(None);
    }
    let node = crate::path::resolve::get(&Nav::new(doc), path)?;
    Ok((node.element_kind() == Kind::Equation).then_some(node))
}

/// The leaf edit address of an equation node.
fn start_cp(node: &Node<'_>) -> Result<u32> {
    match node.item() {
        Item::Equation(equation) => equation.start_cp.ok_or_else(|| {
            error(
                Code::UnsupportedType,
                "DOC equations are written in ordinary main-story paragraphs",
            )
        }),
        _ => Err(error(Code::PathNotFound, "the DOC equation is gone")),
    }
}

impl Doc {
    /// `set` of `latex`, `omml` or `mtef` on an equation; `None` for every
    /// other target.
    pub(super) fn set_equation(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(node) = target_equation(self, target)? else {
            return Ok(None);
        };
        limit(LimitKind::Operations, 1, limits.max_operations())?;
        let source = source(props, limits)?;
        let start = start_cp(&node)?;
        #[cfg(feature = "equation")]
        {
            let Item::Equation(equation) = node.item() else {
                return Err(error(Code::PathNotFound, "the DOC equation is gone"));
            };
            // The object keeps its kind: write the version and product of the
            // MTEF it holds.
            let info = self
                .reader()?
                .equation_native(equation.storage_id)
                .and_then(|native| {
                    litchi_equation::read_mtef(native, &litchi_equation::Limits::DEFAULT).ok()
                })
                .and_then(|read| read.info().cloned())
                .ok_or_else(|| {
                    error(
                        Code::UnsupportedType,
                        "the DOC equation's MTEF does not read, so its kind is unknown",
                    )
                })?;
            let kind = litchi_doc::EquationKind::for_mtef_version(info.version);
            let content = content(
                &source,
                &litchi_equation::MtefTarget::SameAs(info),
                equation.font_size_pt,
            )?;
            if litchi_doc::EquationKind::for_mtef_version(content.mtef_version()) != kind {
                return Err(Error::new(
                    Code::InvalidValue,
                    format!(
                        "the DOC equation is an {} object, which holds MTEF {}; the `mtef` given is MTEF {}",
                        kind.prog_id(),
                        if kind == litchi_doc::EquationKind::EquationEditor3 { 3 } else { 5 },
                        content.mtef_version(),
                    ),
                )
                .with_format(FileFormat::Doc)
                .with_property("mtef")
                .with_suggestion("an equation keeps its kind; remove it and add the new one instead"));
            }
            let mut edit = self.native_edit(limits)?;
            edit.set_equation(start, &content)?;
            self.publish(edit, 1, limits).map(Some)
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = (source, start);
            Err(unavailable())
        }
    }

    /// `add(/body/p[N], equation, {latex | omml | mtef})`: a new MathType
    /// object inline in the paragraph, at its end or beside one of its runs.
    /// `None` for every other kind.
    pub(super) fn add_equation(
        &self,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<(WriteCommit, Path<'static>)>> {
        let kind = node.kind();
        if !["equation", "math", "omath"]
            .iter()
            .any(|name| kind.eq_ignore_ascii_case(name))
        {
            return Ok(None);
        }
        limit(LimitKind::Operations, 1, limits.max_operations())?;
        let Some(position) = parent.paragraph_position() else {
            return Err(error(
                Code::UnsupportedType,
                "DOC equations are added to ordinary main-story paragraphs",
            ));
        };
        let source = source(node.props(), limits)?;
        let paragraph = parent
            .native_paragraph()?
            .ok_or_else(|| error(Code::PathNotFound, "the DOC paragraph is gone"))?;
        let range = paragraph.content_cp_range();
        let start = range.map_or(0, |range| range.start());
        let length = range.map_or(0, |range| range.end().saturating_sub(range.start()));
        let nav = Nav::new(self);
        let run_offset = |path: &crate::path::PathArg<'_>, end: bool| -> Result<u32> {
            let run = crate::path::resolve::get(&nav, path)?;
            match run.body_run() {
                Some((at, from, to)) if at == position => Ok(if end { to } else { from }),
                _ => Err(error(
                    Code::UnsupportedType,
                    "a DOC equation is placed before or after a run of its paragraph",
                )),
            }
        };
        let offset = match node.placement() {
            None | Some(Placement::End) => length,
            Some(Placement::Start) => 0,
            Some(Placement::Before(path)) => run_offset(path, false)?,
            Some(Placement::After(path)) => run_offset(path, true)?,
            Some(Placement::Position(_) | Placement::Into(..)) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "DOC equations are placed at the start or end of a paragraph, or beside one of its runs",
                ));
            },
        };
        // The new equation's ordinal among the paragraph's equations.
        let at = start.saturating_add(offset);
        // The preview is drawn at the size of the text it is inserted
        // before, else of the text before it, as Word draws an equation.
        let runs = paragraph.runs()?;
        let covers = |cp: u32| {
            runs.iter().find(|run| {
                run.cp_range()
                    .is_some_and(|range| range.start() <= cp && cp < range.end())
            })
        };
        let size = font_size_pt(covers(at).or_else(|| at.checked_sub(1).and_then(covers)));
        let index = equations(self, paragraph)?
            .iter()
            .filter(|equation| equation.start_cp.is_some_and(|cp| cp < at))
            .count();
        #[cfg(feature = "equation")]
        {
            // A new equation is a MathType object, except raw MTEF, whose
            // version decides the kind of object that holds it.
            let content = content(&source, &litchi_equation::MtefTarget::MathType, size)?;
            let kind = match source {
                Source::Mtef(_) => {
                    litchi_doc::EquationKind::for_mtef_version(content.mtef_version())
                },
                Source::Latex(_) | Source::Omml(_) => litchi_doc::EquationKind::MathType,
            };
            let mut edit = self.native_edit(limits)?;
            edit.insert_equation(position, offset, kind, &content)?;
            let path = Path::from_steps([
                Step::new("body"),
                Step::new("p").with(Predicate::Index(position)),
                Step::new("equation").with(Predicate::Index(Position::new(index))),
            ]);
            Ok(Some((self.publish(edit, 1, limits)?, path)))
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = (source, index, size);
            Err(unavailable())
        }
    }

    /// `remove` of an equation: its field, storage and preview picture.
    /// `None` for every other target.
    pub(super) fn remove_equation(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<WriteCommit>> {
        let Some(node) = target_equation(self, target)? else {
            return Ok(None);
        };
        limit(LimitKind::Operations, 1, limits.max_operations())?;
        let start = start_cp(&node)?;
        let mut edit = self.native_edit(limits)?;
        edit.remove_equation(start)?;
        self.publish(edit, 1, limits).map(Some)
    }
}
