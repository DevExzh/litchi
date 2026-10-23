//! DOCX template merge (record 0274): placeholders and `MERGEFIELD` fields
//! in every story, through the leaf's story substitution, in one native
//! commit.

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::FileFormat;
use litchi_docx::document::substitute::{Replacement, Story};

use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::format::signature::Signature;
use crate::merge::Options;
use crate::merge::{Planner, Record, Summary, normalize_breaks, validate_value};
use crate::transaction::SemanticLimits;

/// Fill `record` into every story of `file`. `None` means nothing matched:
/// no native commit was made.
pub(crate) fn merge(
    file: &Docx,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let result = merge_inner(file, record, options, limits);
    file.owner().finish(result)
}

fn merge_inner(
    file: &Docx,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let edit_options = litchi_docx::document::EditOptions::default()
        .max_snapshot_bytes(limits.max_snapshot_bytes());
    let editor = owning_editor(file, edit_options)?;
    let WorkEdit::Eager(mut edit) = WorkEdit::from(editor)? else {
        return Err(Error::new(
            Code::InternalError,
            "a DOCX merge needs the owning package editor",
        )
        .with_format(FileFormat::Docx));
    };
    let mut planner = Planner::new(record, options, limits);
    let story = edit.story()?;
    let replacements = plan(&story, &mut planner)?;
    let mut changed = edit.substitute(&story, &replacements)?.paragraphs() > 0;
    for part in edit.story_parts()? {
        let story = edit.part_story(part.name())?;
        let replacements = plan(&story, &mut planner)?;
        if !replacements.is_empty() {
            changed |= edit.substitute_part(&story, &replacements)?.paragraphs() > 0;
        }
    }
    let summary = planner.finish();
    if !changed {
        return Ok((None, summary));
    }
    file.owner().check_source()?;
    let commit = WorkEdit::Eager(edit).commit(file)?;
    Ok((Some(commit), summary))
}

/// An editor over the owning package: story parts publish only through
/// it. An unsigned source-backed owner goes straight to its materialized
/// package, whose edit applies the same admission a source edit would;
/// every other state takes [`Docx::edit`]'s policy first.
pub(super) fn owning_editor(
    file: &Docx,
    options: litchi_docx::document::EditOptions,
) -> Result<DocxEditor> {
    if let (Signature::Unsigned, DocxOwner::Source(package), true) =
        (file.signature(), file.owner(), file.0.retain_source)
    {
        let promoted = file.promoted(package)?;
        return Ok(DocxEditor::Edit(Box::new(
            promoted.edit_with_options(options).map_err(Error::from)?,
        )));
    }
    let editor = file.edit_with_options(options)?;
    file.eager_editor(editor)
}

/// Plan one story: merge fields first, then placeholders outside them.
fn plan(story: &Story, planner: &mut Planner<'_>) -> Result<Vec<Replacement>> {
    let mut replacements = Vec::new();
    for (index, paragraph) in story.paragraphs().iter().enumerate() {
        let text = paragraph.text();
        let mut planned: Vec<(Range<usize>, Cow<'_, str>)> = Vec::new();
        if planner.options().merges_fields() && !paragraph.fields().is_empty() {
            for field in paragraph.fields() {
                if let Some(value) = planner.field(field.instruction(), field.is_replaceable())? {
                    planned.push((field.range(), Cow::Owned(value)));
                }
            }
        }
        let taken: Vec<Range<usize>> = planned.iter().map(|(range, _)| range.clone()).collect();
        planner.text(
            text,
            |range| paragraph.is_text(range),
            &taken,
            |range, value| planned.push((range, Cow::Borrowed(value))),
        )?;
        if !taken.is_empty() {
            planned.sort_by_key(|(range, _)| range.start);
        }
        for (range, value) in planned {
            // Placeholder values were checked by the planner; a field's
            // value carries its switches' text too.
            if matches!(value, Cow::Owned(_)) {
                validate_value(&value)?;
            }
            let value = match normalize_breaks(&value) {
                Cow::Borrowed(_) => value.into_owned(),
                Cow::Owned(normalized) => normalized,
            };
            replacements.push(Replacement::new(index, range, value));
        }
    }
    Ok(replacements)
}
