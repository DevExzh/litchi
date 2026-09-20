//! Resolver-side adapter for `INDIRECT`.
//!
//! The lexical `INDIRECT` parser lives in the scalar lookup module because it
//! is also used by the resolver-free evaluator.  Its result owns the parsed
//! sheet names, however, and those names must not become part of a retained
//! value descriptor.  This module consumes that result while it is borrowed,
//! resolves every sheet name through the immutable resolver, and retains only
//! canonical resolver borrows in the runtime areas.  No cell is read here.

use super::super::{
    EvaluationFailure, EvaluationResult, Rect, Resolver, RuntimeArea, RuntimeAreaSet, RuntimeValue,
    ScalarError, SheetExtent, SheetRef, ValueEvaluator, WorkingValue,
};
use crate::codec::formula::reference::{
    Address, Endpoint, EndpointValue, Reference, SheetSelector,
};

/// Turn one parser-owned reference into a value-side descriptor.
///
/// The `Reference` borrow is intentionally shorter than `'expr`: all retained
/// names come from `Resolver::sheet_name_at`, and the returned runtime value
/// contains no pointer into parser storage.  Callers may therefore drop the
/// parser result immediately after this function returns.
pub(super) fn adapt<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    reference: &Reference,
) -> EvaluationResult<RuntimeValue<'expr>>
where
    R: Resolver + ?Sized,
{
    // Descriptor construction is metadata work, but it still belongs to the
    // caller's cancellation/work budget before any provider operation.
    evaluator.scalar.charge_work(0)?;
    match reference {
        Reference::Source { .. }
            if evaluator.reference_kind_only || evaluator.reject_source_reference =>
        {
            // ISREF and the SHEET/SHEETS metadata boundary can inspect the
            // source-qualified kind without asking an external provider for
            // geometry or values.
            Ok(RuntimeValue::SourceReference)
        },
        Reference::Source { .. } => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        )),
        Reference::Error => Ok(formula_error(ScalarError::Reference)),
        Reference::Local(address) => match local_reference(evaluator, address)? {
            Some(areas) => Ok(RuntimeValue::Areas(areas.mark_scalar_result())),
            None => Ok(formula_error(ScalarError::Reference)),
        },
    }
}

fn formula_error<'a>(error: ScalarError) -> RuntimeValue<'a> {
    RuntimeValue::Scalar(WorkingValue::Error(error))
}

#[derive(Clone, Copy)]
struct ResolvedSheet<'a> {
    sheet: SheetRef<'a>,
    index: usize,
}

fn local_reference<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    address: &Address,
) -> EvaluationResult<Option<RuntimeAreaSet<'expr>>>
where
    R: Resolver + ?Sized,
{
    match address {
        Address::Cell(endpoint) => {
            let Some(area) = resolve_endpoint(evaluator, endpoint, None)? else {
                return Ok(None);
            };
            evaluator.check_reference_cells(area.rect.count()?)?;
            let mut result = RuntimeAreaSet::derived(evaluator)?;
            result.push(area, evaluator)?;
            Ok(Some(result))
        },
        Address::Cells(first, second)
        | Address::Columns(first, second)
        | Address::Rows(first, second) => {
            let Some(first_sheet) = resolve_sheet(evaluator, &first.sheet, None)? else {
                return Ok(None);
            };
            let Some(second_sheet) = resolve_sheet(evaluator, &second.sheet, Some(first_sheet))?
            else {
                return Ok(None);
            };
            let Some(first_rect) = endpoint_rect(evaluator, first, first_sheet.sheet)? else {
                return Ok(None);
            };
            let Some(second_rect) = endpoint_rect(evaluator, second, second_sheet.sheet)? else {
                return Ok(None);
            };
            let Some(rect) = first_rect.bounding(second_rect).ok() else {
                return Ok(None);
            };

            let (start, end) = if first_sheet.index <= second_sheet.index {
                (first_sheet.index, second_sheet.index)
            } else {
                (second_sheet.index, first_sheet.index)
            };
            let sheet_count = resolver_sheet_count(evaluator)?;
            if end >= sheet_count {
                return Ok(None);
            }
            let per_sheet = rect.count()?;
            let sheet_span = end
                .checked_sub(start)
                .and_then(|span| span.checked_add(1))
                .ok_or(EvaluationFailure::InvalidExpression(
                    "INDIRECT sheet span overflows",
                ))?;
            let total = per_sheet
                .checked_mul(sheet_span)
                .ok_or_else(|| evaluator.reference_limit_error(usize::MAX))?;
            evaluator.check_reference_cells(total)?;
            evaluator.check_reference_areas(sheet_span)?;

            // Validate and retain each plane in one pass.  The descriptor is
            // still local until the function returns, so a provider refusal,
            // cancellation, or failed reservation drops all partial state.
            let mut result = RuntimeAreaSet::derived(evaluator)?;
            for index in start..=end {
                evaluator.scalar.charge_work(1)?;
                let sheet = plane_sheet(evaluator, index, first_sheet, second_sheet)?;
                if !rect_within_extent(evaluator, sheet.sheet, rect)? {
                    return Ok(None);
                }
                result.push(
                    RuntimeArea {
                        sheet: sheet.sheet,
                        sheet_index: index,
                        rect,
                    },
                    evaluator,
                )?;
            }
            Ok(Some(result))
        },
    }
}

fn resolve_sheet<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    selector: &SheetSelector,
    inherited: Option<ResolvedSheet<'expr>>,
) -> EvaluationResult<Option<ResolvedSheet<'expr>>>
where
    R: Resolver + ?Sized,
{
    match selector {
        SheetSelector::Current => {
            let position = evaluator.position;
            let name = position.sheet();
            let Some(index) = resolver_sheet_index(evaluator, name)? else {
                return Ok(None);
            };
            Ok(Some(ResolvedSheet {
                sheet: SheetRef::Current,
                index,
            }))
        },
        SheetSelector::Inherited => inherited
            .ok_or(EvaluationFailure::InvalidExpression(
                "INDIRECT inherited endpoint has no first sheet",
            ))
            .map(Some),
        SheetSelector::Explicit(locator) if locator.subtables.is_empty() => {
            let Some(index) = resolver_sheet_index(evaluator, locator.sheet.name.as_str())? else {
                return Ok(None);
            };
            let name = resolver_sheet_name_at(evaluator, index)?.ok_or(
                EvaluationFailure::InvalidExpression("INDIRECT sheet index has no canonical name"),
            )?;
            Ok(Some(ResolvedSheet {
                sheet: SheetRef::Named(name),
                index,
            }))
        },
        // Subtable locators are syntactically valid in the inert reference
        // grammar but have no geometry provider in this evaluator profile.
        SheetSelector::Explicit(_) => Err(EvaluationFailure::Unsupported(
            super::super::super::UnsupportedKind::Reference,
        )),
    }
}

fn plane_sheet<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    index: usize,
    first: ResolvedSheet<'expr>,
    second: ResolvedSheet<'expr>,
) -> EvaluationResult<ResolvedSheet<'expr>>
where
    R: Resolver + ?Sized,
{
    if index == first.index {
        return Ok(first);
    }
    if index == second.index {
        return Ok(second);
    }
    let name = resolver_sheet_name_at(evaluator, index)?.ok_or(
        EvaluationFailure::InvalidExpression("INDIRECT sheet order index has no name"),
    )?;
    Ok(ResolvedSheet {
        sheet: SheetRef::Named(name),
        index,
    })
}

fn resolve_endpoint<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    endpoint: &Endpoint,
    inherited: Option<ResolvedSheet<'expr>>,
) -> EvaluationResult<Option<RuntimeArea<'expr>>>
where
    R: Resolver + ?Sized,
{
    let Some(resolved) = resolve_sheet(evaluator, &endpoint.sheet, inherited)? else {
        return Ok(None);
    };
    let Some(rect) = endpoint_rect(evaluator, endpoint, resolved.sheet)? else {
        return Ok(None);
    };
    if !rect_within_extent(evaluator, resolved.sheet, rect)? {
        return Ok(None);
    }
    Ok(Some(RuntimeArea {
        sheet: resolved.sheet,
        sheet_index: resolved.index,
        rect,
    }))
}

fn rect_within_extent<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    sheet: SheetRef<'expr>,
    rect: Rect,
) -> EvaluationResult<bool>
where
    R: Resolver + ?Sized,
{
    let position = evaluator.position;
    let name = match sheet {
        SheetRef::Current => position.sheet(),
        SheetRef::Named(name) => name,
    };
    let Some(extent) = resolver_sheet_extent(evaluator, name)? else {
        return Ok(false);
    };
    Ok(rect.row_end <= extent.rows() && rect.column_end <= extent.columns())
}

/// Convert an endpoint into a finite rectangle.  Out-of-range coordinates are
/// ordinary formula `#REF!` results, represented by `None`; typed provider and
/// cancellation failures still leave this function through `EvaluationResult`.
fn endpoint_rect<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    endpoint: &Endpoint,
    sheet: SheetRef<'expr>,
) -> EvaluationResult<Option<Rect>>
where
    R: Resolver + ?Sized,
{
    match &endpoint.value {
        EndpointValue::Cell(cell) => {
            let Some(column) = column_number(&cell.column.label) else {
                return Ok(None);
            };
            let Some(row_number) = cell.row.number.checked_sub(1) else {
                return Ok(None);
            };
            let Some(row) = usize::try_from(row_number).ok() else {
                return Ok(None);
            };
            Ok(Rect::cell(row, column).ok())
        },
        EndpointValue::Column(column) => {
            let Some(start) = column_number(&column.label) else {
                return Ok(None);
            };
            let position = evaluator.position;
            let name = match sheet {
                SheetRef::Current => position.sheet(),
                SheetRef::Named(name) => name,
            };
            let Some(extent) = resolver_sheet_extent(evaluator, name)? else {
                return Ok(None);
            };
            let Some(end) = start.checked_add(1) else {
                return Ok(None);
            };
            if end > extent.columns() {
                return Ok(None);
            }
            Ok(Rect::new(0, extent.rows(), start, end).ok())
        },
        EndpointValue::Row(row) => {
            let Some(row_number) = row.number.checked_sub(1) else {
                return Ok(None);
            };
            let Some(start) = usize::try_from(row_number).ok() else {
                return Ok(None);
            };
            let position = evaluator.position;
            let name = match sheet {
                SheetRef::Current => position.sheet(),
                SheetRef::Named(name) => name,
            };
            let Some(extent) = resolver_sheet_extent(evaluator, name)? else {
                return Ok(None);
            };
            let Some(end) = start.checked_add(1) else {
                return Ok(None);
            };
            if end > extent.rows() {
                return Ok(None);
            }
            Ok(Rect::new(start, end, 0, extent.columns()).ok())
        },
    }
}

fn column_number(label: &str) -> Option<usize> {
    if label.is_empty() {
        return None;
    }
    let mut value = 0usize;
    for byte in label.bytes() {
        if !byte.is_ascii_uppercase() {
            return None;
        }
        value = value
            .checked_mul(26)?
            .checked_add(usize::from(byte - b'A' + 1))?;
    }
    value.checked_sub(1)
}

// Resolver metadata is part of the same bounded operation as cell reads.  A
// successful provider call gets a post-call cancellation fence; a provider
// failure is returned unchanged so a typed provider/resource error is not
// reclassified as cancellation.  The wrappers also make every repeated
// extent/name query visible in the work accounting.
fn resolver_sheet_index<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
) -> EvaluationResult<Option<usize>>
where
    R: Resolver + ?Sized,
{
    evaluator.scalar.charge_work(1)?;
    let result = evaluator.resolver.sheet_index(name, evaluator.execution);
    match result {
        Ok(value) => {
            evaluator.scalar.charge_work(0)?;
            Ok(value)
        },
        Err(error) => Err(error),
    }
}

fn resolver_sheet_name_at<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    index: usize,
) -> EvaluationResult<Option<&'expr str>>
where
    R: Resolver + ?Sized,
{
    evaluator.scalar.charge_work(1)?;
    let result = evaluator.resolver.sheet_name_at(index, evaluator.execution);
    match result {
        Ok(value) => {
            evaluator.scalar.charge_work(0)?;
            Ok(value)
        },
        Err(error) => Err(error),
    }
}

fn resolver_sheet_count<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
) -> EvaluationResult<usize>
where
    R: Resolver + ?Sized,
{
    evaluator.scalar.charge_work(1)?;
    let result = evaluator.resolver.sheet_count(evaluator.execution);
    match result {
        Ok(value) => {
            evaluator.scalar.charge_work(0)?;
            Ok(value)
        },
        Err(error) => Err(error),
    }
}

fn resolver_sheet_extent<'expr, 'scalar, 'exec, 'position, R>(
    evaluator: &mut ValueEvaluator<'expr, 'scalar, 'exec, 'position, R>,
    name: &str,
) -> EvaluationResult<Option<SheetExtent>>
where
    R: Resolver + ?Sized,
{
    evaluator.scalar.charge_work(1)?;
    let result = evaluator.resolver.sheet_extent(name, evaluator.execution);
    match result {
        Ok(value) => {
            evaluator.scalar.charge_work(0)?;
            Ok(value)
        },
        Err(error) => Err(error),
    }
}
