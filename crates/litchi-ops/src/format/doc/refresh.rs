//! DOC field refresh (record 0295, record 0279 follow-up F1): the main
//! story's fields computed as DOCX computes them (`DATE`, `TIME`, the
//! document-property and statistics fields, `SEQ`), their cached results
//! rewritten through the native field-result edit
//! (`TextTarget::FieldResult`), all in one native commit. `PAGE`,
//! `NUMPAGES`, `SECTIONPAGES` and `PAGEREF` keep their results
//! (`needs_layout`); `REF` and fields the native edit cannot address are
//! `unsupported`.
//!
//! The native field-result edit addresses the balanced, non-nested fields
//! of the main story that have a result. A document with a nested field or
//! a field without a result separator is reported field by field as
//! `unsupported`, unchanged.

use std::collections::HashMap;

use litchi_core::{FileFormat, Position};
use litchi_doc::body_text::TextTarget;

use super::spec::Kind;
use super::{Doc, Nav, WriteCommit};
use crate::error::{Error, Result};
use crate::path::navigate::{KindSet, Navigate};
use crate::refresh::{Field, Instruction, Options, Status, Summary, word_field};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// The kinds that need a layout engine.
const LAYOUT: [&str; 5] = ["PAGE", "NUMPAGES", "SECTIONPAGES", "PAGEREF", "SECTION"];

/// Refresh the main story's fields. `None` means nothing changed.
pub(crate) fn refresh(
    file: &Doc,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<WriteCommit>, Summary)> {
    file.check_source()?;
    let result = (|| {
        let reader = file.reader()?;
        let fields: Vec<litchi_doc::FieldText> = reader
            .fields()
            .map_err(Error::from)?
            .into_iter()
            .filter(|field| field.field.story == litchi_doc::FieldStory::Main)
            .collect();
        if fields.is_empty() {
            return Ok((None, Summary::default()));
        }
        // The native edit addresses fields only when every main-story field
        // is balanced, top level and has a result.
        let addressable = fields
            .iter()
            .all(|field| field.field.nesting_depth == 0 && field.field.has_separator);
        let metadata = file.metadata()?.unwrap_or_default();
        let mut counts = None;
        let mut seq = HashMap::new();
        let mut summary = Vec::with_capacity(fields.len());
        let mut updates: Vec<(usize, String)> = Vec::new();
        let mut staged = 0usize;
        for (index, field) in fields.iter().enumerate() {
            let instruction = Instruction::parse(&field.instruction);
            let kind = instruction.kind.to_ascii_uppercase();
            let report = |status| Field::new(kind.clone(), Some("/body".to_owned()), status);
            if LAYOUT.contains(&kind.as_str()) {
                summary.push(report(Status::NeedsLayout));
                continue;
            }
            if kind == "REF" {
                summary.push(
                    report(Status::Unsupported).because("DOC bookmarks are not read for REF"),
                );
                continue;
            }
            let computed = word_field(
                &kind,
                &instruction,
                &metadata,
                options.clock_value(),
                &mut seq,
                |_| None,
                || *counts.get_or_insert_with(|| story_counts(file)),
            );
            let text = match computed {
                Ok(text) => text,
                Err((status, reason)) => {
                    let mut entry = report(status);
                    if let Some(reason) = reason {
                        entry = entry.because(reason);
                    }
                    summary.push(entry);
                    continue;
                },
            };
            if !addressable {
                summary.push(
                    report(Status::Unsupported)
                        .because("the main story has nested fields or fields without a result"),
                );
                continue;
            }
            if field.result.as_deref() == Some(text.as_str()) {
                summary.push(report(Status::Unchanged));
                continue;
            }
            staged = staged.saturating_add(text.len());
            if staged > limits.max_staged_bytes() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::StagedBytes,
                    observed: staged,
                    maximum: limits.max_staged_bytes(),
                }
                .into());
            }
            updates.push((index, text));
            summary.push(report(Status::Refreshed));
        }
        if updates.len() > limits.max_expanded_targets() {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: updates.len(),
                maximum: limits.max_expanded_targets(),
            }
            .into());
        }
        let summary = Summary::from_fields(summary);
        if updates.is_empty() {
            return Ok((None, summary));
        }
        let mut edit = file.native_edit(limits)?;
        let operations = updates.len();
        for (index, text) in updates {
            edit.replace_text(TextTarget::FieldResult(Position::new(index)), &text)
                .map_err(Error::from)?;
        }
        Ok((Some(file.publish(edit, operations, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(FileFormat::Doc));
    file.check_source()?;
    result
}

/// The main story's words and non-space characters, field instructions
/// left out.
fn story_counts(file: &Doc) -> (usize, usize) {
    let nav = Nav::new(file);
    let mut counts = (0, 0);
    let Ok(root) = nav.root() else {
        return counts;
    };
    let Ok(bodies) = nav.children(&root, KindSet::of(Kind::Body)) else {
        return counts;
    };
    for body in bodies.flatten() {
        let Ok(paragraphs) = nav.children(&body, KindSet::of(Kind::Paragraph)) else {
            continue;
        };
        for paragraph in paragraphs.flatten() {
            if let Ok(Some(text)) = paragraph.text() {
                // Field instructions are not text: skip from each field
                // begin mark to its separator.
                let mut visible = String::with_capacity(text.len());
                let mut instruction = 0usize;
                for character in text.chars() {
                    match character {
                        '\u{13}' => instruction += 1,
                        '\u{14}' => instruction = instruction.saturating_sub(1),
                        '\u{15}' => visible.push(' '),
                        _ if instruction > 0 => {},
                        _ if character.is_control() => visible.push(' '),
                        _ => visible.push(character),
                    }
                }
                counts.0 += visible.split_whitespace().count();
                counts.1 += visible
                    .chars()
                    .filter(|character| !character.is_whitespace())
                    .count();
            }
        }
    }
    counts
}
