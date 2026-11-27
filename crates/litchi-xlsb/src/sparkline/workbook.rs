#![allow(
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! XLSB package traversal for worksheet sparkline transactions.

use super::worksheet::{self, Commit, Snapshot};
use super::{Formula, FormulaKind, Groups, Limits};
use crate::package::error::{Error, Result};
use crate::package::formula::{Context, SupportingLink};
use litchi_opc::{OpcPackage, PackURI, Part};

const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Read a worksheet's optional sparkline block with explicit limits.
pub(crate) fn read_with_limits(
    package: &OpcPackage,
    worksheet: &PackURI,
    limits: Limits,
) -> Result<Snapshot> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    worksheet::read(part.blob(), limits).map_err(map_error)
}

/// Prove every structurally valid formula against its workbook-owned name and
/// supporting-link context without evaluating or resolving formula values.
pub(crate) fn validate_context(snapshot: &Snapshot, context: &Context) -> Result<()> {
    validate_groups_context(snapshot.groups(), context)
}

pub(crate) fn validate_commit_context(commit: &Commit, context: &Context) -> Result<()> {
    validate_groups_context(commit.groups(), context)
}

/// Prove a detached collection against finalized workbook metadata without
/// evaluating formulas or resolving external targets.
pub(crate) fn validate_groups_context(groups: Option<&Groups>, context: &Context) -> Result<()> {
    let Some(groups) = groups else {
        return Ok(());
    };
    for group in groups.iter() {
        if let Some(formula) = group.date_formula() {
            validate_formula_context(formula, context)?;
        }
        for sparkline in group.sparklines() {
            if let Some(formula) = sparkline.formula() {
                validate_formula_context(formula, context)?;
            }
        }
    }
    Ok(())
}

fn validate_formula_context(formula: &Formula, context: &Context) -> Result<()> {
    match formula.kind() {
        FormulaKind::Name => {
            let index = formula
                .name_index()
                .expect("structurally validated PtgName has an index");
            let resolved = usize::try_from(index)
                .ok()
                .and_then(|value| value.checked_sub(1));
            if !resolved.is_some_and(|value| value < context.defined_names.len()) {
                return Err(Error::InvalidFormula(
                    "sparkline PtgName index exceeds workbook names",
                ));
            }
        },
        FormulaKind::ExternalName => validate_external_name(formula, context)?,
        FormulaKind::Reference3d | FormulaKind::Area3d => {
            let ixti = formula
                .ixti()
                .expect("structurally validated 3-D token has an XTI");
            if !context.is_internal_single_sheet_xti(ixti) {
                return Err(Error::InvalidFormula(
                    "sparkline 3-D source XTI must select exactly one valid internal worksheet",
                ));
            }
        },
    }
    Ok(())
}

fn validate_external_name(formula: &Formula, context: &Context) -> Result<()> {
    let ixti = formula
        .ixti()
        .expect("structurally validated PtgNameX has an XTI");
    let name_index = formula
        .name_index()
        .expect("structurally validated PtgNameX has a name index");
    let xti = context
        .external_sheets
        .get(usize::from(ixti))
        .ok_or(Error::InvalidFormula(
            "sparkline PtgNameX XTI exceeds entries",
        ))?;
    let link_index = usize::try_from(xti.external_link)
        .map_err(|_| Error::InvalidFormula("sparkline external-link index overflow"))?;
    let link = context
        .supporting_links
        .get(link_index)
        .ok_or(Error::InvalidFormula(
            "sparkline PtgNameX refers to missing supporting link",
        ))?;
    let SupportingLink::ExternalWorkbook(book_index) = link else {
        return Err(Error::InvalidFormula(
            "sparkline PtgNameX must refer to an external BrtSupBookSrc",
        ));
    };
    let book_index = usize::try_from(*book_index)
        .map_err(|_| Error::InvalidFormula("sparkline external-book index overflow"))?;
    let book = context
        .external_books
        .get(book_index)
        .ok_or(Error::InvalidFormula(
            "sparkline PtgNameX refers to missing external book",
        ))?;
    let metadata = book.metadata_ref();
    if !metadata.is_workbook() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_sparklines",
            "the sparkline formula refers to inert external data",
        ));
    }
    let resolved = usize::try_from(name_index)
        .ok()
        .and_then(|value| value.checked_sub(1));
    if !resolved.is_some_and(|value| value < metadata.defined_names().len()) {
        return Err(Error::InvalidFormula(
            "sparkline external name index exceeds names",
        ));
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

fn map_error(error: super::Error) -> Error {
    Error::Sparkline(error)
}
