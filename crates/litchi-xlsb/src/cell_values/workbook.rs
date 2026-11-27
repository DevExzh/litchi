#![allow(
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Workbook-level publication for lossless cell-value edits.

use super::worksheet::{Commit, Limits, Snapshot, Value};
use crate::external_link::ExternalLinkLimits;
use crate::package::error::{Error, Result};
use crate::workbook::DrawingLoadPolicy;
use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI, Part};

const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Read editable cell values under an explicit [`crate::Leniency`].
pub(crate) fn read_with_limits_and_leniency(
    package: &OpcPackage,
    worksheet: &PackURI,
    limits: Limits,
    leniency: crate::Leniency,
) -> Result<Snapshot> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    super::worksheet::read_with_limits_and_leniency(part.blob(), limits, leniency)
}

/// The outcome of one validated cell-value publication.
///
/// Validating a candidate requires parsing it as a complete XLSB workbook.
/// This type hands that parse back instead of dropping it, so a caller that
/// owns a [`crate::Workbook`] publishes the parse it already paid for rather
/// than reparsing the same candidate bytes a second time. A caller that owns
/// only an [`OpcPackage`] takes the package back out of it.
pub(crate) enum Applied {
    /// The patch reproduced the stored worksheet bytes exactly, so there is no
    /// candidate: nothing was cloned, nothing was parsed, and the caller's
    /// package is already the published package.
    Unchanged,
    /// The patch changed the worksheet and the candidate passed every check.
    Published {
        /// The validated candidate workbook. Its package carries the updated
        /// worksheet part; a signed source is refused before any candidate is
        /// built. Boxed because a parsed workbook is far larger than a
        /// snapshot.
        workbook: Box<crate::Workbook>,
    },
}

/// Apply one source-checked cell-value commit and hand the validated candidate
/// parse back to the caller without publishing it.
///
/// Nothing the caller owns is touched: every refusal below leaves the caller's
/// package, and any workbook parsed from it, exactly as they were. An exact
/// no-op is answered from the stored bytes, without cloning the package or
/// parsing a candidate at all. The candidate is parsed under the caller's
/// drawing policy, so a skipped drawing projection stays opaque.
///
/// # Errors
///
/// Returns an error for a stale patch, a non-worksheet target, or a candidate
/// that fails complete XLSB workbook validation.
pub(crate) fn apply_retaining_parse(
    package: &OpcPackage,
    worksheet: &PackURI,
    commit: &Commit,
    external_link_limits: ExternalLinkLimits,
    drawing_load_policy: DrawingLoadPolicy,
    leniency: crate::Leniency,
    raw_limits: crate::raw::Limits,
) -> Result<Applied> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    let updated = commit.patch().apply(part.blob())?;
    if updated.as_slice() == part.blob() {
        return Ok(Applied::Unchanged);
    }
    crate::signature_policy::ensure_unsigned(package)?;

    let mut candidate = package.clone();
    crate::worksheet_index::maintain(&mut candidate, worksheet, part.blob(), &updated)?;
    candidate.get_part_mut(worksheet)?.set_blob(updated);
    let parsed = crate::Workbook::reparse_candidate_with_policy(
        candidate,
        external_link_limits,
        drawing_load_policy,
        leniency,
        raw_limits,
    )?;
    let worksheet_index = (0..parsed.worksheet_count())
        .find_map(|index| {
            parsed
                .worksheet_uri(index)
                .ok()
                .filter(|candidate_uri| candidate_uri == worksheet)
                .map(|_| index)
        })
        .ok_or(Error::InvalidFormat(
            "the candidate worksheet part is missing",
        ))?;
    let _ = parsed.worksheet(worksheet_index)?;
    validate_dependencies(commit.snapshot(), &parsed)?;
    Ok(Applied::Published {
        workbook: Box::new(parsed),
    })
}

fn validate_dependencies(snapshot: &Snapshot, workbook: &crate::Workbook) -> Result<()> {
    for cell in snapshot.cells() {
        let reference = cell.reference();
        let style = usize::try_from(cell.style().get()).map_err(|_| {
            Error::InvalidCellReference("cell (, ) style index does not fit this platform")
        })?;
        let format =
            workbook
                .styles()
                .get_cell_format(style)
                .ok_or(Error::InvalidCellReference(
                    "cell (, ) references missing style index",
                ))?;
        for (_resource, _index, present) in [
            (
                "font",
                format.font_id,
                workbook
                    .styles()
                    .get_font(usize::try_from(format.font_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "fill",
                format.fill_id,
                workbook
                    .styles()
                    .get_fill(usize::try_from(format.fill_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "border",
                format.border_id,
                workbook
                    .styles()
                    .get_border(usize::try_from(format.border_id).unwrap_or(usize::MAX))
                    .is_some(),
            ),
            (
                "number format",
                format.num_fmt_id,
                format.num_fmt_id < 164
                    || workbook.styles().get_num_fmt(format.num_fmt_id).is_some(),
            ),
        ] {
            if !present {
                return Err(Error::InvalidCellReference(
                    "cell (, ) style references missing",
                ));
            }
        }
        match cell.value() {
            Value::SharedStringIndex(index) => {
                let index = usize::try_from(*index).map_err(|_| {
                    Error::InvalidCellReference(
                        "cell (, ) shared-string index does not fit this platform",
                    )
                })?;
                let string =
                    workbook
                        .shared_strings()
                        .get(index)
                        .ok_or(Error::InvalidCellReference(
                            "cell (, ) references missing shared-string index",
                        ))?;
                validate_string_fonts(string, workbook, reference)?;
            },
            Value::RichString(string) => validate_string_fonts(string, workbook, reference)?,
            _ => {},
        }
    }
    Ok(())
}

fn validate_string_fonts(
    string: &crate::package::shared_strings::SharedString,
    workbook: &crate::Workbook,
    _reference: super::Reference,
) -> Result<()> {
    for font_id in string
        .runs
        .iter()
        .map(|run| run.font_id)
        .chain(string.phonetic.iter().map(|phonetic| phonetic.font_id))
    {
        if workbook.styles().get_font(usize::from(font_id)).is_none() {
            return Err(Error::InvalidCellReference(
                "cell (, ) rich string references missing font",
            ));
        }
    }
    Ok(())
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == WORKSHEET_CONTENT_TYPE {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: WORKSHEET_CONTENT_TYPE,
        })
    }
}
