//! PPTX template merge (record 0274): placeholders in every paragraph of
//! every slide shape, table cell and notes page, rewritten run by run
//! through the native text edits, so each run keeps its formatting; one
//! native commit. The text bodies and the run rewrite are the ones
//! find/replace uses (record 0285).

use std::ops::Range;

use litchi_drawingml::text::body::RunKind;
use litchi_pptx::opened::{TextAssignment, TextProperties};

use super::find::{Place, RunChange};
use super::{Commit, Pptx};
use crate::find::{Piece, Values, rewrite_ranges};
use crate::merge::{Options, Planner, Record, Summary};
use crate::transaction::SemanticLimits;
use crate::{Code, Error, Result};

/// Fill `record` into every paragraph of every slide shape, table cell
/// and notes page. `None` means nothing matched.
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
    let mut changes: Vec<RunChange> = Vec::new();
    let scopes = file.presentation_scopes(&|text: &str| text.contains("{{"))?;
    for scope in &scopes {
        for (index, paragraph) in scope.body.paragraphs.iter().enumerate() {
            // The paragraph text; breaks and fields are fixed characters,
            // one piece each, so every run is one piece of the shared
            // rewrite.
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
            if values.iter().any(|value| value.contains(['\n', '\r'])) {
                return Err(Error::new(
                    Code::InvalidValue,
                    "a PPTX merge value cannot hold a line break",
                )
                .with_format(litchi_core::FileFormat::Pptx)
                .with_suggestion("remove line breaks from the data or split the paragraph"));
            }
            // The run rewrite find/replace uses (record 0285): a value takes
            // the place of its placeholder's first character, in that run.
            let rewrites = rewrite_ranges(
                &pieces,
                &ranges,
                Values::Each(&values),
                limits.max_staged_bytes(),
            )?;
            for rewrite in rewrites {
                if rewrite.text.is_none() {
                    continue;
                }
                changes.push(RunChange {
                    locus: scope.locus,
                    paragraph: index,
                    run: rewrite.piece,
                    text: rewrite.text,
                    marks: Vec::new(),
                });
            }
        }
    }
    let summary = planner.finish();
    if changes.is_empty() {
        return Ok((None, summary));
    }
    let mut slides: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
    let mut notes: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
    for change in &changes {
        let groups = match change.locus.place {
            Place::Slide => &mut slides,
            Place::Notes => &mut notes,
        };
        if groups
            .last()
            .is_none_or(|(slide, _)| *slide != change.locus.slide)
        {
            groups.push((change.locus.slide, Vec::new()));
        }
        if let (Some((_, edits)), Some(text)) = (groups.last_mut(), &change.text) {
            edits.push(TextAssignment::new(
                change.target(None),
                TextProperties::default().with_text(text),
            ));
        }
    }
    let commit = file.commit_story_text_edits(&slides, &notes, limits)?;
    Ok((Some(commit), summary))
}
