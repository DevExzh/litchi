#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! Semantic remapping of workbook-local formula dependencies.

#![deny(
    clippy::cast_lossless,
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    clippy::checked_conversions,
    clippy::expect_used,
    clippy::float_cmp,
    clippy::let_underscore_must_use,
    clippy::unnecessary_unwrap,
    reason = "formula transfer rewrites only parser-proven fixed-width dependency indexes"
)]

use super::CellFormula;
use crate::Workbook;
use crate::formula::{Parser, Resolution, TableReference, Token};
use crate::package::error::{Error, Result};
use crate::package::formula::{Context, SupportingLink};
use crate::raw::{Records, kind};

#[derive(Debug, Clone, PartialEq, Eq)]
struct NameIdentity {
    name: String,
    scope: Option<String>,
    formula: Option<Vec<u8>>,
    function: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum XtiIdentity {
    SameSheet,
    Internal { first: String, last: String },
    External(String),
}

pub(super) fn remap(
    source: &Workbook,
    target: &Workbook,
    source_sheet: usize,
    target_sheet: usize,
    formula: &CellFormula,
) -> Result<CellFormula> {
    let source_context = source
        .formula_context
        .for_sheet(source.catalog_position_for_worksheet(source_sheet)?);
    let target_context = target
        .formula_context
        .for_sheet(target.catalog_position_for_worksheet(target_sheet)?);
    let mut tokens = formula.tokens().to_vec();
    let parsed = Parser::with_extra(formula.tokens(), formula.ancillary())
        .parse_spanned()
        .map_err(formula_error)?;

    for (token, span) in parsed {
        match token {
            Token::Name(index) => {
                require_span(&span, 5, "PtgName")?;
                write_u32(
                    &mut tokens,
                    span.start + 1,
                    map_name(source, target, index)?,
                )?;
            },
            Token::ExternalName {
                sheet_index,
                name_index,
            } => {
                require_span(&span, 7, "PtgNameX")?;
                let (target_xti, target_name) =
                    map_external_name(&source_context, &target_context, sheet_index, name_index)?;
                write_u16(&mut tokens, span.start + 1, target_xti)?;
                write_u32(&mut tokens, span.start + 3, target_name)?;
            },
            Token::CellRef3d { sheet_index, .. } => {
                require_span(&span, 9, "PtgRef3d")?;
                write_u16(
                    &mut tokens,
                    span.start + 1,
                    map_xti(&source_context, &target_context, sheet_index)?,
                )?;
            },
            Token::AreaRef3d { sheet_index, .. } => {
                require_span(&span, 15, "PtgArea3d")?;
                write_u16(
                    &mut tokens,
                    span.start + 1,
                    map_xti(&source_context, &target_context, sheet_index)?,
                )?;
            },
            Token::ReferenceError {
                is_area,
                sheet_index: Some(sheet_index),
            } => {
                require_span(&span, if is_area { 15 } else { 9 }, "3D reference error")?;
                write_u16(
                    &mut tokens,
                    span.start + 1,
                    map_xti(&source_context, &target_context, sheet_index)?,
                )?;
            },
            Token::TableReference(reference) if !reference.invalid => {
                require_span(&span, 14, "PtgList")?;
                let mapped = map_table(&source_context, &target_context, &reference)?;
                write_u16(&mut tokens, span.start + 2, mapped.sheet_index)?;
                if let Some(table_id) = mapped.list_index {
                    write_u32(&mut tokens, span.start + 6, table_id)?;
                }
            },
            Token::PivotName(index) => {
                require_span(&span, 6, "PtgSxName")?;
                let source_name =
                    Resolution::pivot_name(&source_context, index).map_err(formula_error)?;
                let target_name =
                    Resolution::pivot_name(&target_context, index).map_err(formula_error)?;
                if source_name != target_name {
                    return Err(dependency_error("PivotTable name", &source_name));
                }
            },
            Token::Unknown(kind) => {
                return Err(Error::UnsupportedFeature(format!(
                    "formula transfer cannot prove dependencies for unknown Ptg 0x{kind:02X}"
                )));
            },
            Token::Number(_)
            | Token::String(_)
            | Token::Bool(_)
            | Token::Error(_)
            | Token::Int(_)
            | Token::MissingArg
            | Token::Parenthesis
            | Token::Attribute(_)
            | Token::Array { .. }
            | Token::Memory { .. }
            | Token::CellRef { .. }
            | Token::AreaRef { .. }
            | Token::ReferenceError {
                sheet_index: None, ..
            }
            | Token::BinaryOp(_)
            | Token::UnaryOp(_)
            | Token::Function { .. }
            | Token::TableReference(_) => {},
        }
    }

    CellFormula::new(formula.flags(), tokens, formula.ancillary().to_vec())
}

fn map_name(source: &Workbook, target: &Workbook, index: u32) -> Result<u32> {
    let source_index = usize::try_from(index)
        .ok()
        .and_then(|value| value.checked_sub(1))
        .ok_or_else(|| Error::InvalidFormula("PtgName index is zero".to_string()))?;
    let source_names = defined_names(source)?;
    let identity = source_names
        .get(source_index)
        .ok_or_else(|| Error::InvalidFormula(format!("source PtgName index {index} is absent")))?;
    let candidates = defined_names(target)?
        .iter()
        .enumerate()
        .filter(|(_, candidate)| name_identity_eq(candidate, identity))
        .map(|(candidate, _)| {
            u32::try_from(candidate)
                .ok()
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| Error::InvalidFormula("target name index overflow".to_string()))
        })
        .collect::<Result<Vec<_>>>()?;
    choose(candidates, None, "defined name", &identity.name)
}

fn defined_names(workbook: &Workbook) -> Result<Vec<NameIdentity>> {
    let workbook_uri = workbook.package.main_document_part()?.partname().clone();
    let source = workbook.package.get_part(&workbook_uri)?.blob();
    let mut names = Vec::new();
    for item in Records::new(source) {
        let record = item?;
        if record.kind() != kind::NAME {
            continue;
        }
        let definition = crate::named_ranges::parse(record.payload())
            .map_err(|error| Error::InvalidFormula(error.to_string()))?;
        let scope = definition
            .sheet_id
            .map(|sheet| {
                usize::try_from(sheet)
                    .ok()
                    .and_then(|sheet| workbook.formula_context.worksheet_names.get(sheet))
                    .cloned()
                    .ok_or_else(|| {
                        Error::InvalidFormula(format!(
                            "defined name {:?} has an invalid sheet scope {sheet}",
                            definition.name
                        ))
                    })
            })
            .transpose()?;
        names.push(NameIdentity {
            name: definition.name,
            scope,
            formula: definition.formula,
            function: definition.function,
        });
    }
    Ok(names)
}

fn name_identity_eq(left: &NameIdentity, right: &NameIdentity) -> bool {
    left.name.eq_ignore_ascii_case(&right.name)
        && match (&left.scope, &right.scope) {
            (Some(left), Some(right)) => left.eq_ignore_ascii_case(right),
            (None, None) => true,
            (Some(_), None) | (None, Some(_)) => false,
        }
        && left.formula == right.formula
        && left.function == right.function
}

fn map_xti(source: &Context, target: &Context, index: u16) -> Result<u16> {
    let identity = xti_identity(source, index)?;
    let mut candidates = Vec::new();
    for candidate in 0..target.external_sheets.len() {
        let candidate = u16::try_from(candidate)
            .map_err(|_| Error::InvalidFormula("target XTI index overflow".to_string()))?;
        if xti_identity(target, candidate).is_ok_and(|value| xti_identity_eq(&value, &identity)) {
            candidates.push(candidate);
        }
    }
    choose(
        candidates,
        Some(index),
        "sheet/XTI",
        &format!("{identity:?}"),
    )
}

fn xti_identity(context: &Context, index: u16) -> Result<XtiIdentity> {
    let xti = context
        .external_sheets
        .get(usize::from(index))
        .ok_or_else(|| Error::InvalidFormula(format!("XTI index {index} is absent")))?;
    let link = usize::try_from(xti.external_link)
        .ok()
        .and_then(|index| context.supporting_links.get(index))
        .ok_or_else(|| {
            Error::InvalidFormula(format!(
                "XTI index {index} has no supporting link {}",
                xti.external_link
            ))
        })?;
    match link {
        SupportingLink::SameSheet => {
            if xti.first_sheet != -2 || xti.last_sheet != -2 {
                return Err(Error::InvalidFormula(format!(
                    "same-sheet XTI {index} does not use -2/-2"
                )));
            }
            Ok(XtiIdentity::SameSheet)
        },
        SupportingLink::SelfWorkbook => {
            let first = usize::try_from(xti.first_sheet)
                .ok()
                .and_then(|sheet| context.worksheet_names.get(sheet))
                .cloned()
                .ok_or_else(|| {
                    Error::InvalidFormula(format!("XTI {index} first worksheet is invalid"))
                })?;
            let last = usize::try_from(xti.last_sheet)
                .ok()
                .and_then(|sheet| context.worksheet_names.get(sheet))
                .cloned()
                .ok_or_else(|| {
                    Error::InvalidFormula(format!("XTI {index} last worksheet is invalid"))
                })?;
            Ok(XtiIdentity::Internal { first, last })
        },
        SupportingLink::ExternalWorkbook(_) => Ok(XtiIdentity::External(
            Resolution::sheet_prefix(context, index).map_err(formula_error)?,
        )),
        SupportingLink::AddIn => Err(Error::UnsupportedFeature(format!(
            "formula XTI {index} refers to an add-in"
        ))),
    }
}

fn xti_identity_eq(left: &XtiIdentity, right: &XtiIdentity) -> bool {
    match (left, right) {
        (XtiIdentity::SameSheet, XtiIdentity::SameSheet) => true,
        (
            XtiIdentity::Internal {
                first: left_first,
                last: left_last,
            },
            XtiIdentity::Internal {
                first: right_first,
                last: right_last,
            },
        ) => {
            left_first.eq_ignore_ascii_case(right_first)
                && left_last.eq_ignore_ascii_case(right_last)
        },
        (XtiIdentity::External(left), XtiIdentity::External(right)) => {
            left.eq_ignore_ascii_case(right)
        },
        (XtiIdentity::SameSheet, XtiIdentity::Internal { .. } | XtiIdentity::External(_))
        | (XtiIdentity::Internal { .. } | XtiIdentity::External(_), XtiIdentity::SameSheet)
        | (XtiIdentity::Internal { .. }, XtiIdentity::External(_))
        | (XtiIdentity::External(_), XtiIdentity::Internal { .. }) => false,
    }
}

fn map_external_name(
    source: &Context,
    target: &Context,
    xti: u16,
    name: u32,
) -> Result<(u16, u32)> {
    let identity = Resolution::external_name(source, xti, name).map_err(formula_error)?;
    let maximum_names = target
        .external_books
        .iter()
        .map(|book| book.metadata_ref().defined_names().len())
        .max()
        .unwrap_or(0);
    let mut candidates = Vec::new();
    for candidate_xti in 0..target.external_sheets.len() {
        let candidate_xti = u16::try_from(candidate_xti)
            .map_err(|_| Error::InvalidFormula("target external XTI overflow".to_string()))?;
        for candidate_name in 1..=maximum_names {
            let candidate_name = u32::try_from(candidate_name).map_err(|_| {
                Error::InvalidFormula("target external-name index overflow".to_string())
            })?;
            if Resolution::external_name(target, candidate_xti, candidate_name)
                .is_ok_and(|value| value.eq_ignore_ascii_case(&identity))
            {
                candidates.push((candidate_xti, candidate_name));
            }
        }
    }
    choose(candidates, Some((xti, name)), "external name", &identity)
}

fn map_table(
    source: &Context,
    target: &Context,
    reference: &TableReference,
) -> Result<TableReference> {
    let identity = Resolution::table_reference(source, reference).map_err(formula_error)?;
    let mut candidates = Vec::new();
    if reference.external.is_some() {
        for candidate_xti in 0..target.external_sheets.len() {
            let candidate_xti = u16::try_from(candidate_xti)
                .map_err(|_| Error::InvalidFormula("target table XTI overflow".to_string()))?;
            let mut candidate = reference.clone();
            candidate.sheet_index = candidate_xti;
            if Resolution::table_reference(target, &candidate)
                .is_ok_and(|value| value.eq_ignore_ascii_case(&identity))
            {
                candidates.push(candidate);
            }
        }
    } else {
        for table in target.tables.iter() {
            for candidate_xti in 0..target.external_sheets.len() {
                let candidate_xti = u16::try_from(candidate_xti)
                    .map_err(|_| Error::InvalidFormula("target table XTI overflow".to_string()))?;
                let mut candidate = reference.clone();
                candidate.sheet_index = candidate_xti;
                candidate.list_index = Some(table.table_id());
                if Resolution::table_reference(target, &candidate)
                    .is_ok_and(|value| value.eq_ignore_ascii_case(&identity))
                {
                    candidates.push(candidate);
                }
            }
        }
    }
    choose(
        candidates,
        Some(reference.clone()),
        "structured table",
        &identity,
    )
}

fn choose<T: Clone + PartialEq>(
    candidates: Vec<T>,
    preferred: Option<T>,
    kind: &str,
    identity: &str,
) -> Result<T> {
    if let Some(preferred) = preferred
        && candidates.iter().any(|candidate| candidate == &preferred)
    {
        return Ok(preferred);
    }
    let mut candidates = candidates.into_iter();
    let first = candidates
        .next()
        .ok_or_else(|| dependency_error(kind, identity))?;
    if candidates.next().is_some() {
        return Err(Error::UnsupportedFeature(format!(
            "formula {kind} dependency {identity:?} is ambiguous in the target workbook"
        )));
    }
    Ok(first)
}

fn require_span(span: &std::ops::Range<usize>, expected: usize, kind: &str) -> Result<()> {
    if span.end.saturating_sub(span.start) != expected {
        return Err(Error::InvalidFormula(format!(
            "{kind} dependency span has {} bytes, expected {expected}",
            span.end.saturating_sub(span.start)
        )));
    }
    Ok(())
}

fn write_u16(bytes: &mut [u8], offset: usize, value: u16) -> Result<()> {
    let end = offset.checked_add(2).ok_or(Error::CapacityOverflow {
        resource: "formula dependency u16 range",
    })?;
    bytes
        .get_mut(offset..end)
        .ok_or_else(|| {
            Error::InvalidFormula("formula dependency u16 is out of bounds".to_string())
        })?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) -> Result<()> {
    let end = offset.checked_add(4).ok_or(Error::CapacityOverflow {
        resource: "formula dependency u32 range",
    })?;
    bytes
        .get_mut(offset..end)
        .ok_or_else(|| {
            Error::InvalidFormula("formula dependency u32 is out of bounds".to_string())
        })?
        .copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn dependency_error(kind: &str, identity: &str) -> Error {
    Error::UnsupportedFeature(format!(
        "formula {kind} dependency {identity:?} has no equivalent target resource"
    ))
}

fn formula_error(error: crate::formula::Error) -> Error {
    Error::InvalidFormula(error.to_string())
}
