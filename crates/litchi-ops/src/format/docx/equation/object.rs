//! OLE equation objects of a paragraph: MathType (`Equation.DSMT4`) and
//! Equation Editor 3.0 (`Equation.3`) objects, `…/p[N]/equation[M]`.
//!
//! M counts the paragraph's equation objects in document order, apart from
//! its Office Math (`oMath[M]`, `oMathPara[M]`), as DOC counts the equation
//! objects of a paragraph (`/body/p[N]/equation[M]`). A node reads `mode`
//! (`inline`: Word shows an object in a run), `latex` (its MTEF converted to
//! LaTeX), `mtef` (its `Equation Native` stream, base64), `progId` and
//! `error` (why the object's embedding cannot be read; `latex` and `mtef`
//! are then absent). The embedding is read on first use, bounded by
//! `litchi-equation`'s object limits; `latex`, `mtef` and `error` need the
//! `equation` feature.
//!
//! `add` under a paragraph with `mtef` or `progId` (only `Equation.DSMT4`)
//! inserts a new MathType object (an Equation Editor 3.0 one for raw MTEF
//! 3) beside the paragraph's runs, from `latex`, `omml` or `mtef`; `set
//! latex|omml|mtef` replaces an object's equation in its own kind; `remove`
//! deletes the object with its embedding, preview and relationships. Writes
//! use the package editor (a source-backed owner is promoted) and need the
//! `equation` feature. Objects of a header, footer or note are written the
//! same way, their parts related from the story's part. A new or replaced
//! object's preview is drawn at its run's text size, resolved through the
//! styles.

use std::sync::{Arc, OnceLock};

use litchi_core::FileFormat;
use litchi_docx::object::EquationObject;

use super::super::Docx;
use super::super::node::{Address, Node, Step, Value};
use super::super::spec::Kind;
use super::EquationProps;
use crate::error::{Code, Error, Result};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// An equation object of a paragraph and, once read, its `Equation Native`
/// stream.
pub(crate) struct ObjectMath {
    pub(in super::super) object: EquationObject,
    native: OnceLock<std::result::Result<Arc<Vec<u8>>, String>>,
}

impl ObjectMath {
    /// The `Equation Native` stream of the object of `node`, read once;
    /// `Err` says why it cannot be read. `None` without the `equation`
    /// feature.
    fn native(&self, node: &Node<'_>) -> Option<&std::result::Result<Arc<Vec<u8>>, String>> {
        #[cfg(feature = "equation")]
        {
            Some(self.native.get_or_init(|| read_native(node, &self.object)))
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = (node, &self.native);
            None
        }
    }

    /// The object's equation as LaTeX.
    pub(in super::super) fn latex(&self, node: &Node<'_>) -> Option<String> {
        #[cfg(feature = "equation")]
        {
            let native = self.native(node)?.as_ref().ok()?;
            litchi_equation::mtef_to_latex(native, &litchi_equation::Limits::DEFAULT).ok()
        }
        #[cfg(not(feature = "equation"))]
        {
            let _ = node;
            None
        }
    }

    /// The node's properties.
    pub(in super::super) fn props(&self, node: &Node<'_>) -> EquationProps {
        let (mtef, error) = match self.native(node) {
            Some(Ok(native)) => {
                use base64::Engine as _;
                (
                    Some(base64::engine::general_purpose::STANDARD.encode(native.as_slice())),
                    None,
                )
            },
            Some(Err(message)) => (None, Some(message.clone())),
            None => (None, None),
        };
        EquationProps {
            mode: "inline",
            latex: self.latex(node),
            omml: None,
            mtef,
            prog_id: Some(self.object.prog_id().to_owned()),
            error,
        }
    }
}

/// The node of the equation object at `ordinal` of the paragraph at
/// `paragraph`.
pub(in super::super) fn object_node<'f>(
    docx: &'f Docx,
    paragraph: &Address,
    ordinal: u32,
    object: EquationObject,
) -> Node<'f> {
    Node::new(
        docx,
        Kind::Equation,
        Value::Object(Arc::new(ObjectMath {
            object,
            native: OnceLock::new(),
        })),
        paragraph.with(Step::equation_object(ordinal)),
    )
}

#[cfg(feature = "equation")]
/// The part whose relationships the object of `node` names: `None` for the
/// main document, else its header, footer or notes part.
fn source_part(node: &Node<'_>) -> Result<Option<String>> {
    let steps = node.address.steps();
    let skip = usize::from(
        steps
            .first()
            .is_some_and(|step| matches!(step.kind, Kind::Footnotes | Kind::Endnotes)),
    );
    let Some(story) = steps.get(skip) else {
        return Ok(None);
    };
    if !matches!(
        story.kind,
        Kind::Header | Kind::Footer | Kind::Footnote | Kind::Endnote
    ) {
        return Ok(None);
    }
    let (key, _) = super::super::parts::story_of(node.docx, story.kind, story.ordinal as usize)?;
    match key {
        super::super::element::StoryKey::Part(name) => Ok(Some(name)),
        super::super::element::StoryKey::Note(kind, _) => {
            use litchi_docx::document::substitute::StoryKind;
            let wanted = if kind == litchi_docx::document::NoteKind::Footnote {
                StoryKind::Footnotes
            } else {
                StoryKind::Endnotes
            };
            let docx = node.docx;
            let parts = match docx.owner() {
                super::super::DocxOwner::Eager(package) => {
                    package.story_parts().map_err(Error::from)
                },
                super::super::DocxOwner::Source(package) => {
                    package.story_parts().map_err(Error::from)
                },
            };
            let parts = docx.owner().finish(parts)?;
            Ok(parts
                .into_iter()
                .find(|part| part.kind() == wanted)
                .map(|part| part.name().to_owned()))
        },
    }
}

/// Read the `Equation Native` stream of `object`, the object of `node`.
#[cfg(feature = "equation")]
fn read_native(
    node: &Node<'_>,
    object: &EquationObject,
) -> std::result::Result<Arc<Vec<u8>>, String> {
    if object
        .object_type()
        .is_some_and(|kind| kind.eq_ignore_ascii_case("Link"))
    {
        return Err("the object is linked: its data is in another file".to_owned());
    }
    let id = object
        .object_relationship_id()
        .ok_or_else(|| "the object names no embedded part".to_owned())?;
    let part = source_part(node).map_err(|error| error.to_string())?;
    let docx = node.docx;
    let payload = match docx.owner() {
        super::super::DocxOwner::Eager(package) => package
            .related_part_payload(part.as_deref(), id)
            .map_err(Error::from),
        super::super::DocxOwner::Source(package) => package
            .related_part_payload(part.as_deref(), id)
            .map_err(Error::from),
    };
    let payload = docx
        .owner()
        .finish(payload)
        .map_err(|error| error.to_string())?
        .ok_or_else(|| "the object's embedded part is missing".to_owned())?;
    litchi_equation::equation_object_native(&payload)
        .map(Arc::new)
        .map_err(|error| error.to_string())
}

/// The canonical path of the equation object at `ordinal` of the paragraph
/// `parent`.
#[cfg(feature = "equation")]
pub(in super::super) fn object_path(
    parent: &Node<'_>,
    ordinal: litchi_core::Position,
) -> litchi_core::selector::Path<'static> {
    use litchi_core::selector::{Path, Predicate, Step as PathStep};
    let mut steps = super::super::element::path_with(parent, None)
        .steps()
        .to_vec();
    steps.push(PathStep::new("equation").with(Predicate::Index(ordinal)));
    Path::from_steps(steps)
}

/// Refuse a ProgID other than the one new objects take.
pub(in super::super) fn check_prog_id(prog_id: &str) -> Result<()> {
    if prog_id.eq_ignore_ascii_case("Equation.DSMT4") {
        Ok(())
    } else {
        Err(error(
            Code::InvalidValue,
            "a new equation object is a MathType one: progId is Equation.DSMT4",
        )
        .with_property("progId"))
    }
}

/// The base64 `mtef` of a write, decoded.
pub(in super::super) fn decode_mtef(text: &str) -> Result<Vec<u8>> {
    use base64::Engine as _;
    base64::engine::general_purpose::STANDARD
        .decode(text.trim())
        .map_err(|_| {
            error(
                Code::InvalidValue,
                "mtef is the base64 of an Equation Native stream or of bare MTEF",
            )
            .with_property("mtef")
        })
}

/// The content of an object write, converted for `kind` (a new MathType
/// object for `None`); a conversion that loses part of the equation is
/// refused.
#[cfg(feature = "equation")]
pub(in super::super) fn content(
    source: &super::Source,
    kind: Option<&litchi_equation::MtefTarget>,
    font_size_pt: f64,
) -> Result<litchi_docx::document::EquationContent> {
    use litchi_docx::document::{EquationContent, EquationSource};
    let (source, key) = match source {
        super::Source::Latex(latex) => (EquationSource::Latex(latex), "latex"),
        super::Source::Omml(omml) => (EquationSource::Omml(omml), "omml"),
        super::Source::Mtef(mtef) => (EquationSource::Mtef(mtef), "mtef"),
    };
    let content = EquationContent::new(
        source,
        kind,
        font_size_pt,
        &litchi_equation::Limits::DEFAULT,
    )
    .map_err(|error| {
        Error::new(
            Code::InvalidValue,
            format!("the equation does not convert to MTEF: {error}"),
        )
        .with_property(key)
        .with_format(FileFormat::Docx)
    })?;
    if let Some(loss) = content.losses().first() {
        return Err(Error::new(
            Code::InvalidValue,
            format!("the equation does not convert to MTEF without loss: {loss}"),
        )
        .with_property(key)
        .with_format(FileFormat::Docx)
        .with_suggestion("write the equation as mtef"));
    }
    Ok(content)
}

/// Writing an equation object without the `equation` feature.
#[cfg(not(feature = "equation"))]
pub(in super::super) fn unavailable() -> Error {
    error(
        Code::UnsupportedCapability,
        "writing an equation object needs the `equation` cargo feature",
    )
    .with_suggestion("enable the `equation` cargo feature")
}
