//! PPTX template merge (record 0274): placeholders in every slide
//! paragraph, rewritten run by run through the native text edits, so each
//! run keeps its formatting; one native commit.

use std::ops::Range;

use litchi_core::Position;
use litchi_drawingml::text::body::RunKind;
use litchi_pptx::opened::{TextEdit, TextProperties, TextTarget};

use super::{Commit, Pptx};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::merge::{Options, Planner, Record, Summary, validate_value};
use crate::transaction::SemanticLimits;
use crate::{Code, Error, Result};

/// One run's planned text.
struct RunText {
    slide: usize,
    target: TextTarget,
    text: String,
}

/// Fill `record` into every slide paragraph. `None` means nothing matched.
pub(crate) fn merge(
    file: &Pptx,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    // Refuse a signed or released presentation before reading, as `set`
    // does.
    drop(file.editable_owner()?);
    let mut planner = Planner::new(record, options, limits);
    let mut runs: Vec<RunText> = Vec::new();
    let walk = crate::query::run(Nav::new(file), "paragraph")?;
    for node in walk {
        let node = node?;
        let Item::Paragraph(item) = node.item() else {
            continue;
        };
        let Some(slide) = item.shape.container.slide() else {
            continue;
        };
        let Some(paragraph) = item.paragraph() else {
            continue;
        };
        // The paragraph text; breaks and fields are fixed characters.
        let mut text = String::new();
        let mut spans: Vec<(usize, Range<usize>)> = Vec::new();
        let mut fixed: Vec<Range<usize>> = Vec::new();
        for (index, run) in paragraph.runs.iter().enumerate() {
            let start = text.len();
            match run.kind {
                RunKind::Text => {
                    text.push_str(&run.text);
                    spans.push((index, start..text.len()));
                },
                RunKind::Break => text.push('\n'),
                _ => text.push('\u{FFFC}'),
            }
            if !matches!(run.kind, RunKind::Text) {
                fixed.push(start..text.len());
            }
        }
        let mut planned: Vec<(Range<usize>, &str)> = Vec::new();
        planner.text(
            &text,
            |range| {
                !fixed
                    .iter()
                    .any(|fixed| fixed.start < range.end && range.start < fixed.end)
            },
            &[],
            |range, value| planned.push((range, value)),
        )?;
        if planned.is_empty() {
            continue;
        }
        for (_, value) in &planned {
            validate_value(value)?;
            if value.contains(['\n', '\r']) {
                return Err(Error::new(
                    Code::InvalidValue,
                    "a PPTX merge value cannot hold a line break",
                )
                .with_format(litchi_core::FileFormat::Pptx)
                .with_suggestion("remove line breaks from the data or split the paragraph"));
            }
        }
        for (index, span) in &spans {
            let Some(content) = rewrite(&text, span.clone(), &planned) else {
                continue;
            };
            runs.push(RunText {
                slide: slide as usize,
                target: TextTarget::Run {
                    shape: item.shape.handle.position(),
                    paragraph: Position::new(item.paragraph as usize),
                    run: Position::new(*index),
                },
                text: content,
            });
        }
    }
    let summary = planner.finish();
    if runs.is_empty() {
        return Ok((None, summary));
    }
    runs.sort_by_key(|run| run.slide);
    let mut groups: Vec<(usize, Vec<TextEdit<'_>>)> = Vec::new();
    for run in &runs {
        if groups.last().is_none_or(|(slide, _)| *slide != run.slide) {
            groups.push((run.slide, Vec::new()));
        }
        if let Some((_, edits)) = groups.last_mut() {
            edits.push(TextEdit::new(
                run.target,
                TextProperties::default().with_text(&run.text),
            ));
        }
    }
    let commit = file.commit_text_edits(&groups)?;
    Ok((Some(commit), summary))
}

/// The new text of the run covering `span`, or `None` when no replacement
/// touches it. A value goes to the run holding its placeholder's first
/// character.
fn rewrite(text: &str, span: Range<usize>, planned: &[(Range<usize>, &str)]) -> Option<String> {
    let mut overlapping = planned
        .iter()
        .filter(|(range, _)| range.start < span.end && range.end > span.start)
        .peekable();
    overlapping.peek()?;
    let mut content = String::new();
    let mut cursor = span.start;
    for (range, value) in overlapping {
        if range.start >= cursor {
            content.push_str(&text[cursor..range.start]);
        }
        if range.start >= span.start {
            content.push_str(value);
        }
        cursor = range.end.min(span.end);
    }
    content.push_str(&text[cursor..span.end]);
    Some(content)
}
