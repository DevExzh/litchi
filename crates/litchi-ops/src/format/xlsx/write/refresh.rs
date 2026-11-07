//! XLSX refresh (record 0279): recalculate every formula through the
//! `eval` engine and, with `write_cache`, write each stale or missing
//! result back as the cell's cached value in one native workbook edit. The
//! formula markup is kept byte for byte.

use litchi_xlsx::{self as xlsx};
use std::mem::size_of;

use super::super::Xlsx;
use super::{Commit, native_error};
use crate::calc::{ErrorValue, Status as Calc, Value};
use crate::refresh::{
    Field, Options, Status, Summary, column_letters, formula_report, workbook_summary,
};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{File, Result};

/// Refresh `file`, whose XLSX handle is `handle`. `None` means nothing was
/// written.
pub(crate) fn refresh(
    file: &File,
    handle: &Xlsx,
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
        let source = handle.editor_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
        let mut edit = super::edit_with_snapshot_limit(&source, limits.max_snapshot_bytes())?;
        let mut fields: Vec<Field> = Vec::with_capacity(report.formulas().len());
        let mut writes = 0usize;
        let mut staged = 0usize;
        let mut index = 0;
        let formulas = report.formulas();
        while index < formulas.len() {
            let sheet = formulas[index].sheet();
            let mut worksheet = match names.get(sheet) {
                Some(name) => edit.sheet(name.as_str()).map_err(native_error)?,
                None => None,
            };
            while index < formulas.len() && formulas[index].sheet() == sheet {
                let formula = &formulas[index];
                index += 1;
                let stale = matches!(formula.status(), Calc::Mismatch | Calc::Uncached);
                let (Some(worksheet), true, Some(computed)) =
                    (worksheet.as_mut(), stale, formula.computed())
                else {
                    fields.push(formula_report(&names, formula, false));
                    continue;
                };
                let Some(value) = cache_value(computed) else {
                    fields.push(
                        formula_report(&names, formula, false)
                            .with_status(Status::Unsupported, "the result cannot be stored"),
                    );
                    continue;
                };
                let address = format!("{}{}", column_letters(formula.column()), formula.row() + 1);
                let next_writes = writes.saturating_add(1);
                if next_writes > limits.max_expanded_targets() {
                    return Err(TransactionError::Limit {
                        kind: LimitKind::ExpandedTargets,
                        observed: next_writes,
                        maximum: limits.max_expanded_targets(),
                    }
                    .into());
                }
                let next_staged = staged.saturating_add(cache_value_bytes(&value));
                if next_staged > limits.max_staged_bytes() {
                    return Err(TransactionError::Limit {
                        kind: LimitKind::StagedBytes,
                        observed: next_staged,
                        maximum: limits.max_staged_bytes(),
                    }
                    .into());
                }
                match worksheet.set_cached(address.as_str(), value) {
                    Ok(_) => {},
                    Err(error) if cache_target_is_unsupported(&error) => {
                        fields.push(formula_report(&names, formula, false).with_status(
                            Status::Unsupported,
                            "the cell's formula is not a scalar formula",
                        ));
                        continue;
                    },
                    Err(error) => return Err(native_error(error)),
                }
                writes = next_writes;
                staged = next_staged;
                fields.push(formula_report(&names, formula, true));
            }
        }
        let summary = Summary {
            fields,
            bookmarks: 0,
            missing_page_numbers: 0,
        };
        if writes == 0 {
            return Ok((None, summary));
        }
        let native = edit.commit().map_err(native_error)?;
        Ok((
            Some(Commit::from_native_with_limit(
                handle,
                source,
                native,
                limits.max_snapshot_bytes(),
            )?),
            summary,
        ))
    })();
    let freshness = handle.check_source();
    freshness.and(result)
}

fn cache_target_is_unsupported(error: &xlsx::Error) -> bool {
    matches!(
        error,
        xlsx::Error::Invalid(message)
            if message == "a formula cache needs a stored scalar formula"
    )
}

fn cache_value_bytes(value: &xlsx::Value) -> usize {
    size_of::<xlsx::Value>().saturating_add(match value {
        xlsx::Value::Text(text) => text.as_str().len(),
        xlsx::Value::Number(number) => number.as_str().len(),
        xlsx::Value::Date(date) => date.as_str().len(),
        xlsx::Value::Bool(_) | xlsx::Value::Error(_) => 0,
        _ => 0,
    })
}

/// The stored form of a computed value; `None` for a number a workbook
/// cannot hold.
fn cache_value(value: &Value) -> Option<xlsx::Value> {
    Some(match value {
        Value::Empty => xlsx::Value::text(""),
        Value::Bool(value) => xlsx::Value::from(*value),
        Value::Number(number) => {
            if !number.is_finite() {
                return None;
            }
            xlsx::Value::from(xlsx::Number::new(format!("{number}")).ok()?)
        },
        Value::Text(text) => xlsx::Value::text(text.as_str()),
        Value::Error(error) => xlsx::Value::Error(match error {
            ErrorValue::Null => xlsx::ErrorValue::Null,
            ErrorValue::Div0 => xlsx::ErrorValue::DivZero,
            ErrorValue::Value => xlsx::ErrorValue::Value,
            ErrorValue::Ref => xlsx::ErrorValue::Ref,
            ErrorValue::Name => xlsx::ErrorValue::Name,
            ErrorValue::Num => xlsx::ErrorValue::Num,
            ErrorValue::NotAvailable => xlsx::ErrorValue::NotAvailable,
            ErrorValue::GettingData => xlsx::ErrorValue::GettingData,
            ErrorValue::Spill => xlsx::ErrorValue::Spill,
            ErrorValue::Calc => xlsx::ErrorValue::Calc,
            ErrorValue::Connect => xlsx::ErrorValue::Connect,
            ErrorValue::Blocked => xlsx::ErrorValue::Blocked,
            ErrorValue::Field => xlsx::ErrorValue::Field,
            _ => return None,
        }),
    })
}
