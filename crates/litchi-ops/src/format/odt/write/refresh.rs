//! ODT (and OTH/ODM) field refresh (record 0295, record 0279 follow-up
//! F1): the body's text fields computed where Litchi can (`text:date` and
//! `text:time` with a clock, the document-property and stored-time fields
//! from the metadata) and their displayed text rewritten through the leaf's
//! field edit (`Edit::set_field_text`), all in one leaf commit. Page fields
//! keep their text (`needs_layout`); fixed fields are `locked`; fields with
//! an ODF data style, user and statistics fields are `unsupported`.

use litchi_odt::body::text_fields;

use super::super::Odt;
use super::{Commit, finish, staging_error};
use crate::error::{Error, Result};
use crate::refresh::{Field, Options, Status, Summary, odf_field};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};

/// Refresh the body's text fields. `None` means nothing changed.
pub(crate) fn refresh(
    file: &Odt,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    file.check_source()?;
    let format = file.format();
    let result = (|| {
        file.admit_edit()?;
        let body = file.body()?;
        let fields = text_fields(body).map_err(super::super::core_error)?;
        let metadata = file.metadata()?.unwrap_or_default();
        let mut report = Vec::with_capacity(fields.len());
        let mut updates = Vec::new();
        let mut staged = 0usize;
        for field in &fields {
            let kind = field.name().to_ascii_uppercase();
            let entry = |status| Field::new(kind.clone(), None, status);
            if matches!(
                field.name(),
                "page-number" | "page-count" | "page-continuation"
            ) {
                report.push(entry(Status::NeedsLayout));
                continue;
            }
            let fixed = field.attribute("fixed").as_deref() == Some("true");
            let styled = field.attribute("data-style-name").is_some();
            let text = match odf_field(
                field.name(),
                fixed,
                styled,
                &metadata,
                options.clock_value(),
            ) {
                Ok(text) => text,
                Err((status, reason)) => {
                    let mut skipped = entry(status);
                    if let Some(reason) = reason {
                        skipped = skipped.because(reason);
                    }
                    report.push(skipped);
                    continue;
                },
            };
            if field.text() == text {
                report.push(entry(Status::Unchanged));
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
            updates.push((field, text));
            report.push(entry(Status::Refreshed));
        }
        let summary = Summary::from_fields(report);
        if updates.is_empty() {
            return Ok((None, summary));
        }
        if updates.len() > limits.max_expanded_targets() {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: updates.len(),
                maximum: limits.max_expanded_targets(),
            }
            .into());
        }
        let mut edit = file.package()?.edit();
        let operations = updates.len();
        for (field, text) in updates {
            edit.set_field_text(field, &text).map_err(staging_error)?;
        }
        Ok((Some(finish(file, edit, operations, limits)?), summary))
    })()
    .map_err(|error: Error| error.with_format(format));
    file.check_source()?;
    result
}
