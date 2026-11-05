//! Paragraph and run insertion and removal in existing text bodies
//! (record 0333).
//!
//! A [`TextBody`] names a shape's text body (`p:txBody`), a table cell's
//! (`a:tc/a:txBody`) or a notes-page shape's. Paragraph positions count the
//! body's direct `a:p` children and run positions a paragraph's
//! `a:r`/`a:br`/`a:fld` children, exactly as [`super::TextTarget`] counts
//! them, so a position read from the text-body reader addresses the same
//! element here.
//!
//! Every edit is a byte splice: the inserted element is new markup in the
//! body's own `DrawingML` prefix, a removed element's bytes are cut, and
//! every other byte of the part is kept. A new paragraph or run carries only
//! the formatting the caller assigns (no language or template is copied
//! from its neighbours, unlike `PowerPoint`'s editor). The candidate part is
//! read back as a complete shape scene with unchanged shape identities.

use litchi_core::Position;

use super::{
    Node, Splices, TextProperties, child, children, invalid, is_run, qualified, rewrite_cell_frame,
    rewrite_paragraph, rewrite_rpr, rewrite_run, root,
};
use crate::{Error, Result};

/// The text body a paragraph or run edit changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextBody {
    /// The text body of a slide shape (`p:sp`, by positional scene index).
    Shape {
        /// Positional scene index of the shape.
        shape: Position,
    },
    /// The text body of one cell of a table graphic frame.
    Cell {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
    },
    /// The text body of a shape of the slide's notes page (by positional
    /// scene index of the notes page).
    Notes {
        /// Positional scene index of the notes-page shape.
        shape: Position,
    },
}

impl TextBody {
    pub(crate) const fn shape(self) -> Position {
        match self {
            Self::Shape { shape } | Self::Cell { shape, .. } | Self::Notes { shape } => shape,
        }
    }
}

/// One structural change of a text body.
#[derive(Debug, Clone, Copy)]
pub(crate) enum Change<'a> {
    /// Insert a paragraph before `position` (the paragraph count appends).
    InsertParagraph {
        position: usize,
        properties: TextProperties<'a>,
    },
    /// Remove the paragraph at `paragraph`; a body keeps one paragraph.
    RemoveParagraph { paragraph: usize },
    /// Insert a text run before run child `position` of `paragraph` (the
    /// run count appends).
    InsertRun {
        paragraph: usize,
        position: usize,
        properties: TextProperties<'a>,
    },
    /// Remove run child `run` of `paragraph`.
    RemoveRun { paragraph: usize, run: usize },
}

fn operation(change: &Change<'_>) -> &'static str {
    match change {
        Change::InsertParagraph { .. } => "insert_paragraph",
        Change::RemoveParagraph { .. } => "remove_paragraph",
        Change::InsertRun { .. } => "insert_run",
        Change::RemoveRun { .. } => "remove_run",
    }
}

/// Check the assignments a new paragraph or run may carry.
fn validate(change: &Change<'_>, max_text_bytes: usize) -> Result<()> {
    let (properties, run) = match change {
        Change::InsertParagraph { properties, .. } => (properties, false),
        Change::InsertRun { properties, .. } => (properties, true),
        Change::RemoveParagraph { .. } | Change::RemoveRun { .. } => return Ok(()),
    };
    if properties.has_cell_properties()
        || properties.shape_fill.is_some()
        || properties.geometry.is_some()
        || (run && properties.align.is_some())
    {
        return Err(Error::Invalid(
            if run {
                "a new run takes text and character formatting only"
            } else {
                "a new paragraph takes text, character formatting and alignment only"
            }
            .into(),
        ));
    }
    let text = properties.text.unwrap_or("");
    if text.len() > max_text_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::TextBytes),
            observed: text.len(),
            maximum: max_text_bytes,
        });
    }
    if run && text.contains(['\n', '\r']) {
        return Err(Error::Invalid(
            "run text cannot contain a line break".into(),
        ));
    }
    // The property checks of a formatting edit (font size range, typeface
    // bounds, XML characters) apply to the new element's formatting.
    let check = TextProperties {
        text: properties.text.filter(|text| !text.is_empty()),
        ..*properties
    };
    if !check.is_empty() {
        super::validate(&super::TextAssignment::new(
            super::TextTarget::Shape {
                shape: Position::new(0),
            },
            check,
        ))?;
    }
    Ok(())
}

/// The `DrawingML` prefix of a text body: that of its `bodyPr` or first
/// paragraph.
fn body_prefix(body: &[u8], kids: &[Node]) -> Result<Vec<u8>> {
    child(body, kids, b"bodyPr")
        .or_else(|| child(body, kids, b"p"))
        .map(|node| node.prefix(body).to_vec())
        .ok_or_else(|| invalid("the text body has no drawing children"))
}

/// The markup of a new paragraph: its alignment, then one text run per line
/// (line breaks between them), or, for empty text, the end-of-paragraph
/// properties carrying the character formatting (as `PowerPoint` sizes an
/// empty line by them).
fn new_paragraph(prefix: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let p = qualified(prefix, "p");
    let mut empty = b"<".to_vec();
    empty.extend_from_slice(&p);
    empty.extend_from_slice(b"></");
    empty.extend_from_slice(&p);
    empty.push(b'>');
    let text = properties.text.filter(|text| !text.is_empty());
    let formatting = TextProperties {
        text,
        ..*properties
    };
    let mut paragraph = if formatting.is_empty() {
        empty
    } else {
        rewrite_paragraph(&empty, None, None, &formatting)?
    };
    if text.is_none() && properties.has_run_properties() {
        let mut end = b"<".to_vec();
        end.extend_from_slice(&qualified(prefix, "endParaRPr"));
        end.extend_from_slice(b"/>");
        let end = rewrite_rpr(&end, prefix, properties)?;
        let close = paragraph.len() - (p.len() + 3);
        paragraph.splice(close..close, end);
    }
    Ok(paragraph)
}

/// The markup of a new text run.
fn new_run(prefix: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let r = qualified(prefix, "r");
    let t = qualified(prefix, "t");
    let mut run = b"<".to_vec();
    run.extend_from_slice(&r);
    run.extend_from_slice(b"><");
    run.extend_from_slice(&t);
    run.extend_from_slice(b"/></");
    run.extend_from_slice(&r);
    run.push(b'>');
    let text = properties.text.unwrap_or("");
    let formatting = TextProperties {
        text: None,
        ..*properties
    };
    rewrite_run(&run, &formatting, Some(text))
}

/// Apply `change` to the bytes of one text body element.
fn change_body(body: &[u8], change: &Change<'_>) -> Result<Vec<u8>> {
    let node = root(body)?;
    if node.empty {
        return Err(invalid("an empty text body holds no paragraphs"));
    }
    let kids = children(body, &node)?;
    let paragraphs: Vec<&Node> = kids.iter().filter(|kid| kid.local(body) == b"p").collect();
    let out_of_bounds = |index: usize, len: usize| Error::IndexOutOfBounds { index, len };
    let mut splices = Splices::new();
    match *change {
        Change::InsertParagraph {
            position,
            ref properties,
        } => {
            if position > paragraphs.len() {
                return Err(out_of_bounds(position, paragraphs.len()));
            }
            let prefix = body_prefix(body, &kids)?;
            let at = match (paragraphs.get(position), paragraphs.last()) {
                (Some(next), _) => next.span.start,
                (None, Some(last)) => last.span.end,
                (None, None) => node.close_start,
            };
            splices.insert(at, new_paragraph(&prefix, properties)?);
        },
        Change::RemoveParagraph { paragraph } => {
            let target = paragraphs
                .get(paragraph)
                .ok_or_else(|| out_of_bounds(paragraph, paragraphs.len()))?;
            if paragraphs.len() == 1 {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "remove_paragraph",
                    "a text body keeps at least one paragraph",
                ));
            }
            splices.replace(target.span.clone(), Vec::new());
        },
        Change::InsertRun {
            paragraph,
            position,
            ref properties,
        } => {
            let target = paragraphs
                .get(paragraph)
                .ok_or_else(|| out_of_bounds(paragraph, paragraphs.len()))?;
            let element = &body[target.span.clone()];
            let replaced = insert_run(element, position, properties)?;
            splices.replace(target.span.clone(), replaced);
        },
        Change::RemoveRun { paragraph, run } => {
            let target = paragraphs
                .get(paragraph)
                .ok_or_else(|| out_of_bounds(paragraph, paragraphs.len()))?;
            let element = &body[target.span.clone()];
            let local = root(element)?;
            let inner = children(element, &local)?;
            let runs: Vec<&Node> = inner
                .iter()
                .filter(|kid| is_run(kid.local(element)))
                .collect();
            let removed = runs
                .get(run)
                .ok_or_else(|| out_of_bounds(run, runs.len()))?;
            let mut paragraph_splices = Splices::new();
            paragraph_splices.replace(removed.span.clone(), Vec::new());
            splices.replace(target.span.clone(), paragraph_splices.apply(element)?);
        },
    }
    splices.apply(body)
}

/// Insert a new text run into one paragraph element.
fn insert_run(
    paragraph: &[u8],
    position: usize,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(paragraph)?;
    let prefix = node.prefix(paragraph).to_vec();
    let run = new_run(&prefix, properties)?;
    if node.empty {
        if position > 0 {
            return Err(Error::IndexOutOfBounds {
                index: position,
                len: 0,
            });
        }
        // `<a:p/>` → `<a:p>run</a:p>`.
        let mut output = paragraph[..paragraph.len() - 2].to_vec();
        output.push(b'>');
        output.extend_from_slice(&run);
        output.extend_from_slice(b"</");
        output.extend_from_slice(node.qname(paragraph));
        output.push(b'>');
        return Ok(output);
    }
    let kids = children(paragraph, &node)?;
    let runs: Vec<&Node> = kids
        .iter()
        .filter(|kid| is_run(kid.local(paragraph)))
        .collect();
    if position > runs.len() {
        return Err(Error::IndexOutOfBounds {
            index: position,
            len: runs.len(),
        });
    }
    // Before the run at the position; else after the last run; else after
    // the paragraph properties; and always before `a:endParaRPr`.
    let at = match (runs.get(position), runs.last()) {
        (Some(next), _) => next.span.start,
        (None, Some(last)) => last.span.end,
        (None, None) => child(paragraph, &kids, b"pPr").map_or_else(
            || {
                child(paragraph, &kids, b"endParaRPr")
                    .map_or(node.close_start, |end| end.span.start)
            },
            |ppr| ppr.span.end,
        ),
    };
    let mut splices = Splices::new();
    splices.insert(at, run);
    splices.apply(paragraph)
}

pub(crate) fn stage_text_structure_with_limits(
    xml: &[u8],
    body: TextBody,
    change: &Change<'_>,
    max_text_bytes: usize,
    max_output_bytes: usize,
    shape_limits: crate::shape::Limits,
) -> Result<Vec<u8>> {
    let operation = operation(change);
    validate(change, max_text_bytes)?;
    let original = crate::shape::Scene::read(xml, shape_limits)?;
    if original.is_rewritten() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            operation,
            "text structure edits do not support markup-compatibility branch selection",
        ));
    }
    let shape = original.at(body.shape().get())?;
    let span = shape.common().span()?.range(xml.len())?;
    let is_table = matches!(shape, crate::shape::Shape::Table(_));
    let element = &xml[span.clone()];
    let replacement = match body {
        TextBody::Cell { row, column, .. } => {
            if !is_table {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    operation,
                    "a table cell body requires a table graphic frame",
                ));
            }
            rewrite_cell_frame(element, row.get(), column.get(), |cell| {
                let node = root(cell)?;
                if node.empty {
                    return Err(invalid("an empty a:tc has no text body"));
                }
                let kids = children(cell, &node)?;
                let text = child(cell, &kids, b"txBody").ok_or(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    operation,
                    "the table cell has no text body",
                ))?;
                let rewritten = change_body(&cell[text.span.clone()], change)?;
                let mut splices = Splices::new();
                splices.replace(text.span.clone(), rewritten);
                splices.apply(cell)
            })?
        },
        TextBody::Shape { .. } | TextBody::Notes { .. } => {
            if is_table {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedTarget,
                    operation,
                    "a table frame's text lives in its cells",
                ));
            }
            let node = root(element)?;
            let kids = children(element, &node)?;
            let text = child(element, &kids, b"txBody").ok_or(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                operation,
                "the shape has no text body",
            ))?;
            let rewritten = change_body(&element[text.span.clone()], change)?;
            let mut splices = Splices::new();
            splices.replace(text.span.clone(), rewritten);
            splices.apply(element)?
        },
    };
    let length = xml.len() - span.len() + replacement.len();
    if length > max_output_bytes {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: length,
            maximum: max_output_bytes,
        });
    }
    let mut candidate = Vec::new();
    candidate
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation text structure output",
            source,
        })?;
    candidate.extend_from_slice(&xml[..span.start]);
    candidate.extend_from_slice(&replacement);
    candidate.extend_from_slice(&xml[span.end..]);
    let published = crate::shape::Scene::read(&candidate, shape_limits)?;
    if published.is_rewritten() || published.len() != original.len() {
        return Err(Error::Invalid(
            "opened-presentation text structure edit changed the semantic shape graph".into(),
        ));
    }
    for index in 0..original.len() {
        if original.at(index)?.id() != published.at(index)?.id() {
            return Err(Error::Invalid(
                "opened-presentation text structure edit changed shape identities".into(),
            ));
        }
    }
    Ok(candidate)
}
