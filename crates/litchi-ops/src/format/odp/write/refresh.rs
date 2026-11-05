//! ODP field refresh (record 0295, record 0279 follow-up F1): the field
//! runs of every slide's and notes page's stories. A page number is the
//! slide's position (`text:select-page` previous or next shifting it, a
//! `style:num-format` of `i`, `I`, `a` or `A` applied) and a page count the
//! number of slides, as the PPTX `slidenum` field is computed; clock,
//! stored-time and document-property fields are computed as in ODT. Every
//! changed field is one leaf `SetFieldText` scene change of one commit.
//! Master pages are not visited (their fields are placeholders).

use litchi_core::FileFormat;
use litchi_odp::scene::{Change, Object, ObjectRef, RunKind, Scene, Scope, Story, StoryRef};

use super::super::Odp;
use super::{WriteCommit, commit_changes};
use crate::error::{Error, Result};
use crate::refresh::{Field, Options, Status, Summary, format_number, odf_field};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Refresh every slide's and notes page's field runs. `None` means
/// nothing changed.
pub(crate) fn refresh(
    file: &Odp,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<WriteCommit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        file.admit_edit()?;
        let scene = file.scene()?;
        let content = file.content_xml()?;
        let metadata = file.metadata()?.unwrap_or_default();
        let mut walk = Walk {
            scene,
            content,
            metadata: &metadata,
            options,
            pages: scene.pages().len(),
            report: Vec::new(),
            changes: Vec::new(),
            staged: 0,
            limits,
        };
        for (index, page) in scene.pages().iter().enumerate() {
            for (position, object) in page.objects().iter().enumerate() {
                walk.object(index, object, ObjectRef::new(index, position))?;
            }
            if let Some(notes) = page.notes() {
                for (position, object) in notes.objects().iter().enumerate() {
                    let at = ObjectRef {
                        page: index,
                        scope: Scope::Notes,
                        path: vec![position],
                    };
                    walk.object(index, object, at)?;
                }
            }
        }
        let Walk {
            report, changes, ..
        } = walk;
        let summary = Summary::from_fields(report);
        if changes.is_empty() {
            return Ok((None, summary));
        }
        Ok((Some(commit_changes(file, &changes, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Odp));
    file.check_source()?;
    result
}

struct Walk<'a> {
    scene: &'a Scene,
    content: &'a str,
    metadata: &'a litchi_core::Metadata,
    options: &'a Options,
    pages: usize,
    report: Vec<Field>,
    changes: Vec<Change>,
    staged: usize,
    limits: &'a SemanticLimits,
}

impl Walk<'_> {
    fn object(&mut self, page: usize, object: &Object, at: ObjectRef) -> Result<()> {
        if let Some(story) = object.story() {
            self.story(page, story, &StoryRef::object(at.clone()))?;
        }
        if let Some(table) = object.table() {
            for (row_index, row) in table.rows().iter().enumerate() {
                for (cell_index, cell) in row.cells().iter().enumerate() {
                    if cell.is_covered() {
                        continue;
                    }
                    let reference = StoryRef {
                        object: at.clone(),
                        cell: Some((row_index, cell_index)),
                    };
                    self.story(page, cell.story(), &reference)?;
                }
            }
        }
        for (position, child) in object.children().iter().enumerate() {
            let mut path = at.path.clone();
            path.push(position);
            self.object(
                page,
                child,
                ObjectRef {
                    page: at.page,
                    scope: at.scope,
                    path,
                },
            )?;
        }
        Ok(())
    }

    fn story(&mut self, page: usize, story: &Story, at: &StoryRef) -> Result<()> {
        let _ = self.scene;
        for (index, paragraph) in story.paragraphs().iter().enumerate() {
            for (position, run) in paragraph.runs().iter().enumerate() {
                if run.kind() != RunKind::Field {
                    continue;
                }
                let tag = start_tag(self.content, run.range());
                let name = local_name(tag);
                if matches!(name, "bookmark" | "bookmark-start" | "bookmark-end")
                    || name.starts_with("reference-mark")
                {
                    continue;
                }
                let location = Some(format!("/slide[{}]", page + 1));
                let entry =
                    |status| Field::new(name.to_ascii_uppercase(), location.clone(), status);
                let computed = match name {
                    "page-number" => page_number(tag, page),
                    "page-count" => Ok(self.pages.to_string()),
                    _ => odf_field(
                        name,
                        attribute(tag, "fixed") == Some("true"),
                        attribute(tag, "data-style-name").is_some(),
                        self.metadata,
                        self.options.clock_value(),
                    ),
                };
                let text = match computed {
                    Ok(text) => text,
                    Err((status, reason)) => {
                        let mut skipped = entry(status);
                        if let Some(reason) = reason {
                            skipped = skipped.because(reason);
                        }
                        self.report.push(skipped);
                        continue;
                    },
                };
                if run.text() == text {
                    self.report.push(entry(Status::Unchanged));
                    continue;
                }
                self.staged = self.staged.saturating_add(text.len());
                if self.staged > self.limits.max_staged_bytes() {
                    return Err(TransactionError::Limit {
                        kind: LimitKind::StagedBytes,
                        observed: self.staged,
                        maximum: self.limits.max_staged_bytes(),
                    }
                    .into());
                }
                if self.changes.len() >= self.limits.max_expanded_targets() {
                    return Err(TransactionError::Limit {
                        kind: LimitKind::ExpandedTargets,
                        observed: self.changes.len() + 1,
                        maximum: self.limits.max_expanded_targets(),
                    }
                    .into());
                }
                self.changes.push(Change::SetFieldText {
                    story: at.clone(),
                    paragraph: index,
                    run: position,
                    text,
                });
                self.report.push(entry(Status::Refreshed));
            }
        }
        Ok(())
    }
}

/// The page number of a `text:page-number` field on the slide at `page`.
fn page_number(tag: &str, page: usize) -> std::result::Result<String, crate::refresh::Skip> {
    let number = match attribute(tag, "select-page") {
        Some("previous") => page,
        Some("next") => page + 2,
        _ => page + 1,
    };
    if attribute(tag, "fixed") == Some("true") {
        return Err((Status::Locked, Some("the field is fixed (text:fixed)")));
    }
    let number = u64::try_from(number).unwrap_or(u64::MAX);
    let format = match attribute(tag, "num-format") {
        None | Some("1" | "") => return Ok(number.to_string()),
        Some("i") => "roman",
        Some("I") => "ROMAN",
        Some("a") => "alphabetic",
        Some("A") => "ALPHABETIC",
        Some(_) => {
            return Err((
                Status::Unsupported,
                Some("the page number format is not applied"),
            ));
        },
    };
    format_number(number, format).ok_or((
        Status::Unsupported,
        Some("the page number format is not applied"),
    ))
}

/// The start tag of the element at `range`.
fn start_tag(content: &str, range: std::ops::Range<usize>) -> &str {
    let element = content.get(range).unwrap_or("");
    element.find('>').map_or(element, |end| &element[..=end])
}

/// The local name of a start tag.
fn local_name(tag: &str) -> &str {
    let name = tag
        .trim_start_matches('<')
        .split(|character: char| character.is_whitespace() || character == '>' || character == '/')
        .next()
        .unwrap_or("");
    name.rsplit(':').next().unwrap_or(name)
}

/// The raw value of the attribute whose local name is `local`, whatever
/// its prefix.
fn attribute<'t>(tag: &'t str, local: &str) -> Option<&'t str> {
    let mut rest = tag;
    while let Some(equals) = rest.find("=\"") {
        let name = rest[..equals]
            .rsplit(|character: char| character.is_whitespace())
            .next()
            .unwrap_or("");
        let value_start = equals + 2;
        let value_end = rest[value_start..].find('"')? + value_start;
        if name.rsplit(':').next() == Some(local) {
            return Some(&rest[value_start..value_end]);
        }
        rest = &rest[value_end + 1..];
    }
    None
}
