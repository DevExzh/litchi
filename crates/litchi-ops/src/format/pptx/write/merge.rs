//! PPTX template merge (record 0274): placeholders in every slide
//! paragraph, rewritten run by run through the native text edits, so each
//! run keeps its formatting; one native commit. The run rewrite is the one
//! find/replace uses (`find::rewrite_ranges`, record 0285).

use std::ops::Range;

use litchi_core::Position;
use litchi_drawingml::text::body::RunKind;
use litchi_pptx::opened::{TextEdit, TextProperties, TextTarget};

use super::{Commit, Pptx};
use crate::find::{Piece, Values, rewrite_ranges};
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::merge::{Options, Planner, Record, Summary};
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
        // The paragraph text; breaks and fields are fixed characters, one
        // piece each, so every run is one piece of the shared rewrite.
        let mut text = String::new();
        let mut fixed: Vec<Range<usize>> = Vec::new();
        let mut pieces: Vec<Piece<'_>> = Vec::with_capacity(paragraph.runs.len());
        for run in &paragraph.runs {
            let start = text.len();
            let piece = match run.kind {
                RunKind::Text => run.text.as_str(),
                RunKind::Break => "\n",
                _ => "\u{FFFC}",
            };
            text.push_str(piece);
            if !matches!(run.kind, RunKind::Text) {
                fixed.push(start..text.len());
            }
            pieces.push(Piece {
                text: piece,
                group: 0,
                scoped: true,
            });
        }
        let mut ranges: Vec<Range<usize>> = Vec::new();
        let mut values: Vec<&str> = Vec::new();
        planner.text(
            &text,
            |range| {
                !fixed
                    .iter()
                    .any(|fixed| fixed.start < range.end && range.start < fixed.end)
            },
            &[],
            |range, value| {
                ranges.push(range);
                values.push(value);
            },
        )?;
        if ranges.is_empty() {
            continue;
        }
        for value in &values {
            if value.contains(['\n', '\r']) {
                return Err(Error::new(
                    Code::InvalidValue,
                    "a PPTX merge value cannot hold a line break",
                )
                .with_format(litchi_core::FileFormat::Pptx)
                .with_suggestion("remove line breaks from the data or split the paragraph"));
            }
        }
        // The run rewrite find/replace uses (record 0285): a value takes the
        // place of its placeholder's first character, in that run.
        let rewrites = rewrite_ranges(
            &pieces,
            &ranges,
            Values::Each(&values),
            limits.max_staged_bytes(),
        )?;
        for rewrite in rewrites {
            let Some(content) = rewrite.text else {
                continue;
            };
            runs.push(RunText {
                slide: slide as usize,
                target: TextTarget::Run {
                    shape: item.shape.handle.position(),
                    paragraph: Position::new(item.paragraph as usize),
                    run: Position::new(rewrite.piece),
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
