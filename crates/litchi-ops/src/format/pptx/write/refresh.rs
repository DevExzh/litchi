//! PPTX refresh (record 0279): slide-number and date fields (`a:fld`) in
//! slide text, rewritten through the native run-text edits in one commit.

use litchi_core::Position;
use litchi_drawingml::text::body::RunKind;
use litchi_pptx::opened::{TextAssignment, TextProperties, TextTarget};

use super::{Commit, Editable, Pptx};
use crate::Result;
use crate::format::pptx::Nav;
use crate::format::pptx::node::Item;
use crate::refresh::{Clock, Field, Options, Status, Summary};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// One field's planned text.
struct RunText {
    slide: usize,
    target: TextTarget,
    text: String,
}

/// The PowerPoint `datetime` field formats (en-US).
fn date_picture(kind: &str) -> Option<&'static str> {
    Some(match kind {
        "datetime" | "datetime1" => "M/d/yyyy",
        "datetime2" => "dddd, MMMM d, yyyy",
        "datetime3" => "d MMMM yyyy",
        "datetime4" => "MMMM d, yyyy",
        "datetime5" => "d-MMM-yy",
        "datetime6" => "MMMM yy",
        "datetime7" => "MMM-yy",
        "datetime8" => "M/d/yyyy h:mm AM/PM",
        "datetime9" => "M/d/yyyy h:mm:ss AM/PM",
        "datetime10" => "H:mm",
        "datetime11" => "H:mm:ss",
        "datetime12" => "h:mm AM/PM",
        "datetime13" => "h:mm:ss AM/PM",
        _ => return None,
    })
}

fn result(
    kind: &str,
    slide: usize,
    first: u32,
    clock: Option<Clock>,
) -> std::result::Result<String, Status> {
    if kind == "slidenum" {
        return Ok((u64::from(first) + slide as u64).to_string());
    }
    let picture = date_picture(kind).ok_or(Status::Unsupported)?;
    clock
        .map(|clock| clock.format(picture))
        .ok_or(Status::NeedsClock)
}

/// Refresh every slide field. `None` means nothing changed.
pub(crate) fn refresh(
    file: &Pptx,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let first = match file.editable()? {
        Editable::Eager(package) => package.first_slide_number()?,
        Editable::Source(package) => package.first_slide_number()?,
    };
    let mut summary = Summary::default();
    let mut runs: Vec<RunText> = Vec::new();
    let mut staged = 0usize;
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
        for (index, run) in paragraph.runs.iter().enumerate() {
            let RunKind::Field(field) = &run.kind else {
                continue;
            };
            let kind = field.kind.as_deref().unwrap_or("").to_ascii_lowercase();
            let location = Some(format!("/slide[{}]", slide as usize + 1));
            let report = |status| Field::new(kind.to_ascii_uppercase(), location.clone(), status);
            let text = match result(&kind, slide as usize, first, options.clock_value()) {
                Ok(text) => text,
                Err(status) => {
                    summary.fields.push(report(status));
                    continue;
                },
            };
            if text == run.text {
                summary.fields.push(report(Status::Unchanged));
                continue;
            }
            staged = staged.saturating_add(text.len() + 64);
            if runs.len() + 1 > limits.max_expanded_targets() || staged > limits.max_staged_bytes()
            {
                return Err(TransactionError::Limit {
                    kind: if staged > limits.max_staged_bytes() {
                        LimitKind::StagedBytes
                    } else {
                        LimitKind::ExpandedTargets
                    },
                    observed: if staged > limits.max_staged_bytes() {
                        staged
                    } else {
                        runs.len() + 1
                    },
                    maximum: if staged > limits.max_staged_bytes() {
                        limits.max_staged_bytes()
                    } else {
                        limits.max_expanded_targets()
                    },
                }
                .into());
            }
            runs.push(RunText {
                slide: slide as usize,
                target: TextTarget::Run {
                    shape: item.shape.handle.position(),
                    paragraph: Position::new(item.paragraph as usize),
                    run: Position::new(index),
                },
                text,
            });
            summary.fields.push(report(Status::Refreshed));
        }
    }
    if runs.is_empty() {
        return Ok((None, summary));
    }
    runs.sort_by_key(|run| run.slide);
    let mut groups: Vec<(usize, Vec<TextAssignment<'_>>)> = Vec::new();
    for run in &runs {
        if groups.last().is_none_or(|(slide, _)| *slide != run.slide) {
            groups.push((run.slide, Vec::new()));
        }
        if let Some((_, edits)) = groups.last_mut() {
            edits.push(TextAssignment::new(
                run.target,
                TextProperties::default().with_text(&run.text),
            ));
        }
    }
    let commit = file.commit_text_edits(&groups, limits)?;
    Ok((Some(commit), summary))
}
