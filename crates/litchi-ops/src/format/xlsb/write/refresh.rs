//! XLSB refresh (record 0295, record 0279 follow-up F8): recalculate every
//! formula through the `eval` engine and, with `write_cache`, write each
//! stale or missing result back as the formula record's cached value: one
//! native cell-values edit per changed worksheet, chained, in one unified
//! commit. The formula tokens are kept; a result whose type needs another
//! record family than the stored one (a number cache becoming text) is
//! reported `unsupported` and left as it is.

use std::collections::BTreeMap;

use litchi_xlsb::cell_values::CellError;
use litchi_xlsb::sheet::Selector as NativeSelector;

use super::super::Xlsb;
use super::{
    Commit, NativeCellOptions, Reference, cell_patch_container_bytes, cell_snapshot_budget,
    cell_transfer_limits, native_error, package_from_editor,
};
use crate::calc::{ErrorValue, Status as Calc, Value};
use crate::refresh::{Field, Options, Status, Summary, formula_report, workbook_summary};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{File, Result};

/// The stored form of a computed result.
enum Cache {
    Number(f64),
    Boolean(bool),
    Error(CellError),
    Text(String),
}

/// Refresh `file`, whose XLSB handle is `handle`. `None` means nothing was
/// written.
pub(crate) fn refresh(
    file: &File,
    handle: &Xlsb,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<Commit>, Summary)> {
    let report = file.recalculate(options.calc_options())?;
    if !options.writes_cache() {
        return Ok((None, workbook_summary(file, &report, false)?));
    }
    let names = crate::calc::sheet_names(file)?;
    handle.check_source()?;
    let result = (|| {
        let mut fields: Vec<Field> = Vec::with_capacity(report.formulas().len());
        // Planned writes per worksheet: (reference, cache, field index).
        let mut planned: BTreeMap<usize, Vec<(Reference, Cache, usize)>> = BTreeMap::new();
        let mut writes = 0usize;
        for formula in report.formulas() {
            let stale = matches!(formula.status(), Calc::Mismatch | Calc::Uncached);
            let cache = formula.computed().and_then(cache_value);
            let (true, Some(cache)) = (stale, cache) else {
                let field = formula_report(&names, formula, false);
                fields.push(if stale {
                    field.with_status(Status::Unsupported, "the result cannot be stored")
                } else {
                    field
                });
                continue;
            };
            writes += 1;
            if writes > limits.max_expanded_targets() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::ExpandedTargets,
                    observed: writes,
                    maximum: limits.max_expanded_targets(),
                }
                .into());
            }
            let reference =
                Reference::new(formula.row(), formula.column()).map_err(native_error)?;
            planned
                .entry(formula.sheet())
                .or_default()
                .push((reference, cache, fields.len()));
            fields.push(formula_report(&names, formula, true));
        }
        if planned.is_empty() {
            return Ok((None, Summary::from_fields(fields)));
        }
        let mut remaining_snapshot = cell_snapshot_budget(limits.max_snapshot_bytes())?;
        let canonical_limits = cell_transfer_limits(remaining_snapshot, limits)?;
        let canonical_sheet = planned.keys().next().copied();
        let mut package = package_from_editor(handle.editor_for_write()?);
        let mut patches = Vec::new();
        let mut retained_snapshot = cell_patch_container_bytes();
        for (sheet, caches) in planned {
            let native_limits = cell_transfer_limits(remaining_snapshot, limits)?;
            let Some(mut edit) = package
                .edit_cell_values_with_options(
                    NativeSelector::from(sheet),
                    NativeCellOptions::new().transfer_limits(native_limits),
                )
                .map_err(native_error)?
                .resolve()
                .map_err(native_error)?
            else {
                for (_, _, index) in caches {
                    refuse(&mut fields[index]);
                }
                continue;
            };
            let mut staged = 0usize;
            for (reference, cache, index) in caches {
                let written = match cache {
                    Cache::Number(value) => edit.set_formula_number_cache(reference, value),
                    Cache::Boolean(value) => edit.set_formula_boolean_cache(reference, value),
                    Cache::Error(value) => edit.set_formula_error_cache(reference, value),
                    Cache::Text(value) => edit.set_formula_string_cache(reference, value),
                };
                if written.is_err() {
                    refuse(&mut fields[index]);
                } else {
                    staged += 1;
                }
            }
            if staged == 0 {
                continue;
            }
            let (next, patch) = edit.commit().map_err(native_error)?.into_parts();
            let retained = patch.retained_bytes();
            if retained > remaining_snapshot {
                let consumed = limits
                    .max_snapshot_bytes()
                    .saturating_sub(remaining_snapshot);
                return Err(super::snapshot_limit(
                    consumed.saturating_add(retained),
                    limits.max_snapshot_bytes(),
                ));
            }
            remaining_snapshot -= retained;
            retained_snapshot = retained_snapshot
                .checked_add(retained)
                .ok_or_else(|| super::snapshot_limit(usize::MAX, limits.max_snapshot_bytes()))?;
            package = next;
            patches.push(patch);
        }
        let summary = Summary::from_fields(fields);
        if patches.is_empty() {
            return Ok((None, summary));
        }
        let canonical_sheet = canonical_sheet.ok_or_else(|| {
            crate::Error::new(
                crate::Code::InternalError,
                "XLSB refresh lost its worksheet owner",
            )
            .with_format(litchi_core::FileFormat::Xlsb)
        })?;
        let commit = Commit::from_cell_parts(
            handle,
            package,
            patches,
            (canonical_sheet, canonical_limits),
            retained_snapshot,
            limits.max_snapshot_bytes(),
            writes,
        )?;
        Ok((Some(commit), summary))
    })();
    let freshness = handle.check_source();
    freshness.and(result)
}

/// A planned write the native editor refused: the stored record family
/// cannot hold the result's type.
fn refuse(field: &mut Field) {
    *field = field.clone().with_status(
        Status::Unsupported,
        "the formula's stored result has another type than the computed one",
    );
}

/// The stored form of a computed value; `None` for a value an XLSB
/// formula record cannot cache.
fn cache_value(value: &Value) -> Option<Cache> {
    Some(match value {
        Value::Empty => Cache::Text(String::new()),
        Value::Bool(value) => Cache::Boolean(*value),
        Value::Number(number) if number.is_finite() => Cache::Number(*number),
        Value::Number(_) => return None,
        Value::Text(text) => Cache::Text(text.clone()),
        Value::Error(error) => Cache::Error(match error {
            ErrorValue::Null => CellError::Null,
            ErrorValue::Div0 => CellError::DivisionByZero,
            ErrorValue::Value => CellError::Value,
            ErrorValue::Ref => CellError::Reference,
            ErrorValue::Name => CellError::Name,
            ErrorValue::Num => CellError::Number,
            ErrorValue::NotAvailable => CellError::NotAvailable,
            ErrorValue::GettingData => CellError::GettingData,
            _ => return None,
        }),
    })
}
