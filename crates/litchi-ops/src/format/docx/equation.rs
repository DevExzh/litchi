//! Office Math equations.
//!
//! | Path | Equation |
//! |---|---|
//! | `/body/oMathPara[N]`, `…/tc[K]/oMathPara[N]`, `/header[N]/oMathPara[M]` | display math of a block container: a `w:p` holding one `m:oMathPara`, or a bare `m:oMathPara` (read only) |
//! | `/body/p[N]/oMath[M]`, `/footnote[N]/p[M]/oMath[K]` | inline math of a paragraph: its M-th `m:oMath` outside a math paragraph, in document order |
//! | `/body/p[N]/oMathPara[M]` | display math of a paragraph that also holds other content (a bookmark, text, inline math) |
//!
//! Each reads `mode` (`display` or `inline`), `latex` (its OMML converted to
//! LaTeX; absent without the `equation` feature or when the OMML does not
//! convert) and `omml` (the `m:oMathPara` of a display equation or the
//! `m:oMath` of an inline one, as the main-document reader holds it).
//!
//! `add` under a paragraph inserts inline math beside the paragraph's runs;
//! under the body, a cell, a content control, a header, a footer or a note
//! it inserts a display-equation paragraph. `set` replaces an equation's
//! content from `latex` or `omml`, keeping the paragraph's properties and a
//! display equation's math-paragraph properties; `remove` removes an inline
//! equation's `m:oMath`, a paragraph's math paragraph, or a display
//! equation's paragraph (only its math when the paragraph is the last of a
//! cell, a content control or a story). LaTeX is converted through
//! `litchi-equation` and needs the `equation` feature; a conversion the
//! converter reports as dropping formatting is refused (it reports no loss
//! for a LaTeX command it does not know, which it writes as text).

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_docx::document::{
    ElementAnchor, ElementPath, ElementStep, ParagraphEquation, raw::Tree,
};
use litchi_docx::{OfficeMath, OfficeMathParagraph};

use super::element::{StoryKey, element_path, path_with, run_locations, story_path};
use super::node::{Address, Node, Step, Value};
use super::setters::Setter;
use super::spec::{EQUATION_REGISTRY, Kind};
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::node::Value as NodeValue;
use crate::props::Props;
use crate::registry::Operation;
use crate::value::Scalar;

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// Math of a paragraph: an inline `m:oMath`, or an `m:oMathPara` of a
/// paragraph with other content, as the reader holds it.
pub(super) struct InlineMath {
    omml: String,
    pub(super) display: bool,
}

impl InlineMath {
    pub(super) fn omml(&self) -> &str {
        &self.omml
    }
}

/// The node of the inline equation (`display == false`) or math paragraph
/// at `ordinal` of the paragraph at `paragraph`.
pub(super) fn inline_node<'f>(
    docx: &'f Docx,
    paragraph: &Address,
    ordinal: u32,
    omml: String,
    display: bool,
) -> Node<'f> {
    let step = if display {
        Step::math_paragraph(ordinal)
    } else {
        Step::new(Kind::Equation, ordinal)
    };
    Node::new(
        docx,
        Kind::Equation,
        Value::Math(std::sync::Arc::new(InlineMath { omml, display })),
        paragraph.with(step),
    )
}

/// An equation's properties, in the order of the equation table.
pub(crate) struct EquationProps {
    pub(super) mode: &'static str,
    pub(super) latex: Option<String>,
    pub(super) omml: Option<String>,
}

/// The value of the equation property at `index`.
pub(super) fn value<'a>(props: &EquationProps, index: usize) -> Option<NodeValue<'a>> {
    match index {
        0 => Some(NodeValue::Token(props.mode)),
        1 => props
            .latex
            .as_ref()
            .map(|latex| NodeValue::Text(latex.clone().into())),
        2 => props
            .omml
            .as_ref()
            .map(|omml| NodeValue::Text(omml.clone().into())),
        _ => None,
    }
}

/// The LaTeX of an OMML equation; `None` without the `equation` feature
/// and for OMML that does not convert.
pub(super) fn latex(omml: &str) -> Option<String> {
    #[cfg(feature = "equation")]
    {
        litchi_equation::omml_to_latex(omml, &litchi_equation::Limits::DEFAULT).ok()
    }
    #[cfg(not(feature = "equation"))]
    {
        let _ = omml;
        None
    }
}

/// The `m:oMathPara` of a display equation, as the reader holds it.
pub(super) fn display_omml(block: &super::node::Block) -> Result<Option<String>> {
    let indexed = block.indexed()?;
    let Some(paragraph) = indexed.paragraph() else {
        return Ok(Some(
            String::from_utf8_lossy(&indexed.self_contained_xml()?).into_owned(),
        ));
    };
    Ok(paragraph.math_paragraphs()?.into_iter().next())
}

// ---------------------------------------------------------------------------
// Writes
// ---------------------------------------------------------------------------

/// Where new or replaced content comes from.
enum Source {
    Latex(String),
    Omml(String),
}

/// An admitted `add` or `set`: the requested mode and the content.
struct Request {
    mode: Option<Mode>,
    source: Source,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Mode {
    Display,
    Inline,
}

/// Admit an equation's `add` or `set` properties: exactly one of `latex`
/// and `omml`, and, for `add`, the `mode`.
fn request(props: &Props<'_>, operation: Operation) -> Result<Request> {
    let mut mode = None;
    let mut source = None;
    for (key, value) in props.iter() {
        let (_, entry) = EQUATION_REGISTRY
            .resolve(key.as_str())
            .map_err(|error| error.with_property(key.as_str().to_owned()))?;
        entry
            .prepare(operation, value, |setter, value| {
                let text = match value {
                    Some(Scalar::Text(text) | Scalar::Token(text)) => text.to_owned(),
                    Some(scalar) => scalar.to_string(),
                    None => String::new(),
                };
                let content = match setter {
                    Setter::EquationMode => {
                        mode = Some(if text == "inline" {
                            Mode::Inline
                        } else {
                            Mode::Display
                        });
                        return Ok(());
                    },
                    Setter::EquationLatex => Source::Latex(text),
                    Setter::EquationOmml => Source::Omml(text),
                    _ => {
                        return Err(Error::new(
                            Code::UnsupportedProperty,
                            "the property does not write a DOCX equation",
                        ));
                    },
                };
                if source.replace(content).is_some() {
                    return Err(Error::new(
                        Code::InvalidArgument,
                        "an equation takes its content from one of latex and omml",
                    ));
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    let source = source.ok_or_else(|| {
        Error::new(
            Code::MissingProperty,
            "a DOCX equation needs its content: latex or omml",
        )
        .with_property("latex")
        .with_format(FileFormat::Docx)
    })?;
    Ok(Request { mode, source })
}

/// One `m:oMath` from LaTeX, refused when the converter reports that the
/// conversion drops formatting.
#[cfg(feature = "equation")]
fn from_latex(latex: &str) -> Result<OfficeMath> {
    // Every conversion failure is a property of the given LaTeX: it does
    // not parse, nests too deeply, is too large, or holds characters XML
    // cannot carry.
    let (math, losses) =
        OfficeMath::from_latex(latex, &litchi_equation::Limits::DEFAULT).map_err(|error| {
            Error::new(
                Code::InvalidValue,
                format!("the LaTeX does not convert to OMML: {error}"),
            )
            .with_property("latex")
            .with_format(FileFormat::Docx)
        })?;
    if !losses.is_empty() {
        return Err(Error::new(
            Code::InvalidValue,
            format!(
                "the LaTeX holds what OMML cannot carry: {}",
                losses.join("; ")
            ),
        )
        .with_property("latex")
        .with_format(FileFormat::Docx));
    }
    Ok(math)
}

#[cfg(not(feature = "equation"))]
fn from_latex(_latex: &str) -> Result<OfficeMath> {
    Err(Error::new(
        Code::UnsupportedCapability,
        "converting LaTeX to an equation needs the `equation` cargo feature",
    )
    .with_suggestion("enable the `equation` cargo feature, or give the equation as omml")
    .with_property("latex")
    .with_format(FileFormat::Docx))
}

fn invalid_omml(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message)
        .with_property("omml")
        .with_format(FileFormat::Docx)
}

/// The content of an inline equation: one `m:oMath`.
fn inline_math(source: &Source) -> Result<OfficeMath> {
    match source {
        Source::Latex(latex) => from_latex(latex),
        Source::Omml(omml) => OfficeMath::from_xml(omml.as_str())
            .map_err(|_| invalid_omml("inline math is one well-formed m:oMath element")),
    }
}

/// The content of a display equation: one `m:oMath`, or a whole
/// `m:oMathPara` given as OMML.
enum Display {
    Equation(OfficeMath),
    Paragraph(OfficeMathParagraph),
}

fn display_math(source: &Source) -> Result<Display> {
    match source {
        Source::Latex(latex) => from_latex(latex).map(Display::Equation),
        Source::Omml(omml) => match OfficeMathParagraph::from_xml(omml.as_str()) {
            Ok(paragraph) => Ok(Display::Paragraph(paragraph)),
            Err(_) => OfficeMath::from_xml(omml.as_str())
                .map(Display::Equation)
                .map_err(|_| {
                    invalid_omml("display math is one well-formed m:oMath or m:oMathPara element")
                }),
        },
    }
}

/// A leaf error of an equation write. New markup that holds content an
/// equation does not own is the caller's invalid `omml`, and says why.
fn leaf(error: litchi_docx::document::Error) -> Error {
    match error {
        litchi_docx::document::Error::Refused {
            reason: reason @ litchi_docx::document::Refusal::EquationContent,
            ..
        } => Error::new(Code::InvalidValue, format!("the OMML is refused: {reason}"))
            .with_property("omml")
            .with_format(FileFormat::Docx),
        error => Error::from(error),
    }
}

/// The leaf verbs an equation write uses, on the main story's editor and on
/// a story's.
trait Equations {
    fn insert_inline(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        math: &OfficeMath,
    ) -> Result<Position>;
    fn insert_display(
        &mut self,
        container: &ElementPath,
        anchor: ElementAnchor,
        math: &OfficeMathParagraph,
    ) -> Result<Position>;
    fn replace(
        &mut self,
        paragraph: &ElementPath,
        which: ParagraphEquation,
        math: &OfficeMath,
    ) -> Result<bool>;
    fn replace_paragraph(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMathParagraph,
    ) -> Result<bool>;
    fn replace_paragraph_equations(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMath,
    ) -> Result<bool>;
    fn remove(&mut self, paragraph: &ElementPath, which: ParagraphEquation) -> Result<()>;
    fn remove_paragraph(&mut self, paragraph: &ElementPath, index: Position) -> Result<()>;
    fn remove_element(&mut self, path: &ElementPath) -> Result<()>;
}

macro_rules! equations {
    ($type:ty) => {
        impl Equations for $type {
            fn insert_inline(
                &mut self,
                paragraph: &ElementPath,
                at: Option<Position>,
                math: &OfficeMath,
            ) -> Result<Position> {
                self.insert_equation_at(paragraph, at, math).map_err(leaf)
            }

            fn insert_display(
                &mut self,
                container: &ElementPath,
                anchor: ElementAnchor,
                math: &OfficeMathParagraph,
            ) -> Result<Position> {
                self.insert_display_equation_at(container, anchor, math)
                    .map_err(leaf)
            }

            fn replace(
                &mut self,
                paragraph: &ElementPath,
                which: ParagraphEquation,
                math: &OfficeMath,
            ) -> Result<bool> {
                self.replace_equation_at(paragraph, which, math)
                    .map_err(leaf)
            }

            fn replace_paragraph(
                &mut self,
                paragraph: &ElementPath,
                index: Position,
                math: &OfficeMathParagraph,
            ) -> Result<bool> {
                self.replace_math_paragraph_at(paragraph, index, math)
                    .map_err(leaf)
            }

            fn replace_paragraph_equations(
                &mut self,
                paragraph: &ElementPath,
                index: Position,
                math: &OfficeMath,
            ) -> Result<bool> {
                self.replace_math_paragraph_equations_at(paragraph, index, math)
                    .map_err(leaf)
            }

            fn remove(&mut self, paragraph: &ElementPath, which: ParagraphEquation) -> Result<()> {
                self.remove_equation_at(paragraph, which).map_err(leaf)?;
                Ok(())
            }

            fn remove_paragraph(&mut self, paragraph: &ElementPath, index: Position) -> Result<()> {
                self.remove_math_paragraph_at(paragraph, index)
                    .map_err(leaf)?;
                Ok(())
            }

            fn remove_element(&mut self, path: &ElementPath) -> Result<()> {
                self.remove_element_at(path).map_err(leaf)?;
                Ok(())
            }
        }
    };
}

equations!(Tree);
equations!(litchi_docx::document::Edit);
equations!(litchi_docx::document::Edit<litchi_docx::source_backed::Package>);

/// Run `apply` on the main story's editor, or on the story `key`'s.
fn in_story<T>(
    edit: &mut WorkEdit,
    story: Option<&StoryKey>,
    apply: impl FnOnce(&mut dyn Equations) -> Result<T>,
) -> Result<T> {
    match (story, edit) {
        (None, WorkEdit::Eager(edit)) => apply(edit.as_mut()),
        (None, WorkEdit::Source(edit)) => apply(edit.as_mut()),
        (Some(key), WorkEdit::Eager(edit)) => {
            edit.edit_story(super::story::target(key), |story| apply(story))
        },
        (Some(_), WorkEdit::Source(_)) => Err(error(
            Code::InternalError,
            "a story edit needs the package editor",
        )),
    }
}

/// Which equation of a paragraph a node is.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Which {
    /// The inline equation at this position.
    Inline(Position),
    /// The math paragraph at this position among the paragraph's.
    MathParagraph {
        index: Position,
        /// What a removal takes: the whole `w:p` of a display-equation
        /// block, or only the math paragraph (in a paragraph with other
        /// content, or in the last paragraph of a cell, a content control
        /// or a story, which keeps it).
        whole: bool,
    },
}

/// An equation node's paragraph, its story and which of its equations it
/// is.
struct Target {
    story: Option<StoryKey>,
    paragraph: ElementPath,
    which: Which,
}

impl Target {
    fn of(docx: &Docx, node: &Node<'_>) -> Result<Self> {
        let located = |node: &Node<'_>| -> Result<(Option<StoryKey>, ElementPath)> {
            Ok(match story_path(node)? {
                Some((key, path)) => (Some(key), path),
                None => (None, element_path(node)?),
            })
        };
        match &node.value {
            Value::Math(math) => {
                let ordinal = node
                    .address
                    .steps()
                    .last()
                    .map(|step| Position::new(step.ordinal as usize))
                    .ok_or_else(|| {
                        error(Code::InternalError, "the equation node has no address")
                    })?;
                let parent = super::nested::parent_node(docx, node)?;
                let (story, paragraph) = located(&parent)?;
                Ok(Self {
                    story,
                    paragraph,
                    which: if math.display {
                        Which::MathParagraph {
                            index: ordinal,
                            whole: false,
                        }
                    } else {
                        Which::Inline(ordinal)
                    },
                })
            },
            Value::Equation(block) => {
                if block.indexed()?.paragraph().is_none() {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "a math paragraph outside a w:p is read but not edited",
                    ));
                }
                let (story, paragraph) = located(node)?;
                let (container, _) = paragraph
                    .split_last()
                    .ok_or_else(|| error(Code::InternalError, "an equation has no container"))?;
                // Every container but the main body keeps a paragraph: a
                // cell, a content control (whose cell or story would
                // otherwise lose its last paragraph) and a story.
                let keeps = story.is_some() || !container.steps().is_empty();
                let paragraphs = block
                    .container
                    .iter()
                    .filter(|entry| entry.paragraph().is_some())
                    .count();
                Ok(Self {
                    story,
                    paragraph,
                    which: Which::MathParagraph {
                        index: Position::new(0),
                        whole: !(keeps && paragraphs == 1),
                    },
                })
            },
            _ => Err(error(Code::InternalError, "the node is not an equation")),
        }
    }

    /// The order removals run in: later paragraphs and equations first.
    fn key(&self) -> (Option<String>, Vec<(u8, usize)>, u8, usize) {
        let (kind, position) = match self.which {
            Which::Inline(position) => (0, position.get()),
            Which::MathParagraph { index, .. } => (1, index.get()),
        };
        (
            self.story.as_ref().map(|story| format!("{story:?}")),
            super::nested::order_key(&self.paragraph),
            kind,
            position,
        )
    }
}

impl Docx {
    /// `add` of an equation under a paragraph (inline) or a block container
    /// (display).
    pub(super) fn add_equation(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let mut request = request(node.props(), Operation::Add)?;
        // `oMathPara` names display math and `oMath` inline math, as their
        // paths do; `equation` and `math` leave it to `mode`.
        let named = if node.kind().eq_ignore_ascii_case("oMathPara") {
            Some(Mode::Display)
        } else if node.kind().eq_ignore_ascii_case("oMath") {
            Some(Mode::Inline)
        } else {
            None
        };
        if let (Some(named), Some(mode)) = (named, request.mode)
            && named != mode
        {
            return Err(error(
                Code::InvalidArgument,
                "the equation's element name and its mode disagree",
            )
            .with_property("mode"));
        }
        request.mode = request.mode.or(named);
        match parent.element_kind() {
            Kind::Paragraph => {
                if request.mode == Some(Mode::Display) {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "display math is added under the body, a cell, a content control, a header, a footer or a note",
                    ));
                }
                self.add_inline(editor, parent, node, &request.source)
            },
            Kind::Body
            | Kind::Cell
            | Kind::Sdt
            | Kind::Header
            | Kind::Footer
            | Kind::Footnote
            | Kind::Endnote => {
                if request.mode == Some(Mode::Inline) {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "inline math is added under a paragraph",
                    ));
                }
                self.add_display(editor, parent, node, &request.source)
            },
            _ => Err(error(
                Code::UnsupportedCapability,
                "an equation is added under a paragraph or a block container",
            )),
        }
    }

    fn add_inline(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        source: &Source,
    ) -> Result<(Commit, Path<'static>)> {
        let Value::Paragraph { paragraph, .. } = &parent.value else {
            return Err(error(
                Code::InternalError,
                "the equation parent holds no paragraph",
            ));
        };
        let (story, path) = match story_path(parent)? {
            Some((key, path)) => (Some(key), path),
            None => (None, element_path(parent)?),
        };
        // The new equation goes before a direct run, as a picture does.
        let locations = run_locations(self, paragraph)?;
        let index = super::nested::sibling_position(
            self,
            parent,
            Kind::Run,
            locations.len(),
            node.placement(),
        )?;
        let at = match index {
            None => None,
            Some(index) if index >= locations.len() => None,
            Some(index) => match locations.get(index).map(Vec::as_slice) {
                Some([ElementStep::Run(run)]) => Some(*run),
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "an equation can only be added beside a direct run of the paragraph",
                    ));
                },
            },
        };
        let math = inline_math(source)?;
        self.owner().check_source()?;
        let mut edit = self.equation_edit(editor, story.is_some())?;
        let ordinal = in_story(&mut edit, story.as_ref(), |edit| {
            edit.insert_inline(&path, at, &math)
        })?;
        let commit = edit.commit(self)?;
        Ok((commit, path_with(parent, Some((Kind::Equation, ordinal)))))
    }

    fn add_display(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        source: &Source,
    ) -> Result<(Commit, Path<'static>)> {
        let math = match display_math(source)? {
            Display::Equation(math) => OfficeMathParagraph::from_equation(math),
            Display::Paragraph(paragraph) => paragraph,
        };
        let story = matches!(
            parent.element_kind(),
            Kind::Header | Kind::Footer | Kind::Footnote | Kind::Endnote
        );
        if story {
            let key = super::story::story_key(parent)?;
            let anchor = self.story_anchor(parent, node.placement())?;
            let blocks = super::nav::container_of(parent)?
                .ok_or_else(|| error(Code::InternalError, "a story holds no blocks"))?;
            self.owner().check_source()?;
            let mut edit = self.equation_edit(editor, true)?;
            let raw = in_story(&mut edit, Some(&key), |edit| {
                edit.insert_display(&ElementPath::body(), anchor, &math)
            })?;
            let ordinal = blocks
                .iter()
                .filter(|block| Kind::of_block(block.kind()) == Some(Kind::Equation))
                .filter(|block| {
                    block
                        .direct_position()
                        .is_some_and(|position| position < raw)
                })
                .count();
            let commit = edit.commit(self)?;
            return Ok((
                commit,
                path_with(parent, Some((Kind::Equation, Position::new(ordinal)))),
            ));
        }
        let container = super::nested::container_path(parent)?;
        let anchor = super::nested::block_anchor(self, parent, &container, node.placement(), None)?;
        self.owner().check_source()?;
        let mut edit = self.equation_edit(editor, false)?;
        let raw = in_story(&mut edit, None, |edit| {
            edit.insert_display(&container, anchor, &math)
        })?;
        let commit = edit.commit(self)?;
        let ordinal =
            super::nested::visible_child(self, commit.file(), parent, Kind::Equation, raw)?;
        Ok((commit, path_with(parent, Some((Kind::Equation, ordinal)))))
    }

    /// The editor of an equation write: a story's is the owning package's.
    fn equation_edit(&self, editor: DocxEditor, story: bool) -> Result<WorkEdit> {
        if story {
            WorkEdit::from(self.eager_editor(editor)?)
        } else {
            WorkEdit::from(editor)
        }
    }

    /// A `set` whose targets are equations; `Err` hands the editor back for
    /// other targets.
    pub(super) fn try_set_equation(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty() || nodes.iter().any(|node| node.kind != Kind::Equation) {
            return Ok(Err(editor));
        }
        let request = request(props, Operation::Set)?;
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            targets.push(Target::of(self, node)?);
        }
        let inline = targets
            .iter()
            .any(|target| matches!(target.which, Which::Inline(_)))
            .then(|| inline_math(&request.source))
            .transpose()?;
        let display = targets
            .iter()
            .any(|target| matches!(target.which, Which::MathParagraph { .. }))
            .then(|| display_math(&request.source))
            .transpose()?;
        self.owner().check_source()?;
        let stories = targets.iter().any(|target| target.story.is_some());
        let mut edit = self.equation_edit(editor, stories)?;
        for target in &targets {
            in_story(&mut edit, target.story.as_ref(), |edit| {
                match (target.which, &inline, &display) {
                    (Which::Inline(position), Some(math), _) => {
                        edit.replace(&target.paragraph, ParagraphEquation::Inline(position), math)?;
                    },
                    (Which::MathParagraph { index, .. }, _, Some(Display::Paragraph(math))) => {
                        edit.replace_paragraph(&target.paragraph, index, math)?;
                    },
                    // The math paragraph keeps its properties and holds the
                    // one new equation.
                    (Which::MathParagraph { index, .. }, _, Some(Display::Equation(math))) => {
                        edit.replace_paragraph_equations(&target.paragraph, index, math)?;
                    },
                    _ => {
                        return Err(error(
                            Code::InternalError,
                            "an equation write has no content for its target",
                        ));
                    },
                }
                Ok(())
            })?;
        }
        edit.commit(self).map(Ok)
    }

    /// Remove equations selected by their paths.
    pub(super) fn remove_equations(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<Commit> {
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            targets.push(Target::of(self, node)?);
        }
        // Later equations first, so no removal shifts another.
        targets.sort_by_key(|target| std::cmp::Reverse(target.key()));
        targets.dedup_by(|left, right| left.key() == right.key());
        self.owner().check_source()?;
        let stories = targets.iter().any(|target| target.story.is_some());
        let mut edit = self.equation_edit(editor, stories)?;
        for target in &targets {
            in_story(&mut edit, target.story.as_ref(), |edit| {
                match target.which {
                    Which::Inline(position) => {
                        edit.remove(&target.paragraph, ParagraphEquation::Inline(position))
                    },
                    // The paragraph exists for its equation.
                    Which::MathParagraph { whole: true, .. } => {
                        edit.remove_element(&target.paragraph)
                    },
                    Which::MathParagraph {
                        index,
                        whole: false,
                    } => edit.remove_paragraph(&target.paragraph, index),
                }
            })?;
        }
        edit.commit(self)
    }
}
