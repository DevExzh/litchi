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
    worksheet::read_with_limits(part.blob(), limits).map_err(map_error)
}

/// Apply an exact-source commit atomically to one worksheet part.
pub(crate) fn apply(
    package: &mut OpcPackage,
    worksheet: &PackURI,
    commit: Commit,
) -> Result<Snapshot> {
    let part = package.get_part(worksheet)?;
    require_worksheet(part)?;
    let current = part.blob();
    let (_, limits, patch) = commit.into_publication();
    let (changed, updated) = patch.apply_owned(current).map_err(map_error)?;
    if !changed {
        return worksheet::read_owned(updated, limits).map_err(map_error);
    }

    let snapshot = worksheet::read_with_limits(&updated, limits).map_err(map_error)?;
    let mut candidate = package.clone();
    candidate.get_part_mut(worksheet)?.set_blob(updated);
    candidate.unsign();
    crate::Workbook::from_opc_package(candidate.clone())?;
    *package = candidate;
    Ok(snapshot)
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
                return Err(Error::InvalidFormula(format!(
                    "sparkline PtgName index {index} exceeds {} workbook names",
                    context.defined_names.len()
                )));
            }
        },
        FormulaKind::ExternalName => validate_external_name(formula, context)?,
        FormulaKind::Reference3d | FormulaKind::Area3d => {
            let ixti = formula
                .ixti()
                .expect("structurally validated 3-D token has an XTI");
            if !context.is_internal_single_sheet_xti(ixti) {
                return Err(Error::InvalidFormula(format!(
                    "sparkline 3-D source XTI {ixti} must select exactly one valid internal worksheet"
                )));
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
        .ok_or_else(|| {
            Error::InvalidFormula(format!(
                "sparkline PtgNameX XTI {ixti} exceeds {} entries",
                context.external_sheets.len()
            ))
        })?;
    let link_index = usize::try_from(xti.external_link)
        .map_err(|_| Error::InvalidFormula("sparkline external-link index overflow".to_string()))?;
    let link = context.supporting_links.get(link_index).ok_or_else(|| {
        Error::InvalidFormula(format!(
            "sparkline PtgNameX refers to missing supporting link {}",
            xti.external_link
        ))
    })?;
    let SupportingLink::ExternalWorkbook(book_index) = link else {
        return Err(Error::InvalidFormula(
            "sparkline PtgNameX must refer to an external BrtSupBookSrc".to_string(),
        ));
    };
    let book_index = usize::try_from(*book_index)
        .map_err(|_| Error::InvalidFormula("sparkline external-book index overflow".to_string()))?;
    let book = context.external_books.get(book_index).ok_or_else(|| {
        Error::InvalidFormula(format!(
            "sparkline PtgNameX refers to missing external book {book_index}"
        ))
    })?;
    let metadata = book.metadata_ref();
    if !metadata.is_workbook() {
        return Err(Error::UnsupportedFeature(
            "sparkline PtgNameX refers to inert DDE or OLE data rather than an external workbook"
                .to_string(),
        ));
    }
    let resolved = usize::try_from(name_index)
        .ok()
        .and_then(|value| value.checked_sub(1));
    if !resolved.is_some_and(|value| value < metadata.defined_names().len()) {
        return Err(Error::InvalidFormula(format!(
            "sparkline external name index {name_index} exceeds {} names",
            metadata.defined_names().len()
        )));
    }
    Ok(())
}

fn require_worksheet(part: &dyn Part) -> Result<()> {
    if part.content_type() == WORKSHEET_CONTENT_TYPE {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: WORKSHEET_CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        })
    }
}

fn map_error(error: super::Error) -> Error {
    Error::InvalidFormat(error.to_string())
}
