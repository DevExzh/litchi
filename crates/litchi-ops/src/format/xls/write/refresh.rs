//! XLS refresh (record 0295, record 0279 follow-up F8): recalculate every
//! formula through the `eval` engine and, with `write_cache`, write each
//! stale or missing result back as the `Formula` record's cached result
//! (`FormulaCache`), keeping the formula tokens. Writes are committed in
//! native edits of [`CHUNK`] cells, each on the previous one's result,
//! inside one unified mutation (as CSV import does). A result the record
//! cannot cache is reported `unsupported` and left as it is.

use litchi_core::Position;
use litchi_xls::cell_values::{CellError, FormulaCache, Reference, Selector, Value as NativeValue};
use std::mem::size_of;

use super::super::Xls;
use super::{WriteCommit, finish_commit, native_edit_error, native_value_error};
use crate::calc::{ErrorValue, Status as Calc, Value};
use crate::refresh::{Field, Options, Status, Summary, formula_report, workbook_summary};
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{File, Result};

/// The cells one native BIFF8 edit stages before it commits.
const CHUNK: usize = 1_024;

/// Refresh `file`, whose XLS handle is `handle`: the chained commits
/// (empty when nothing was written) and the report.
pub(crate) fn refresh(
    file: &File,
    handle: &Xls,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Vec<WriteCommit>, Summary)> {
    let report = file.recalculate(options.calc_options())?;
    if !options.writes_cache() {
        return Ok((Vec::new(), workbook_summary(file, &report, false)?));
    }
    let names = file.workbook()?.worksheet_names()?;
    handle.check_source()?;
    let result = (|| {
        handle.admit_edit_for_write()?;
        let mut fields: Vec<Field> = Vec::with_capacity(report.formulas().len());
        let mut planned: Vec<(usize, Reference, FormulaCache, usize)> = Vec::new();
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
            if planned.len() >= limits.max_expanded_targets() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::ExpandedTargets,
                    observed: planned.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: limits.max_expanded_targets(),
                }
                .into());
            }
            let reference =
                Reference::new(formula.row(), formula.column()).map_err(native_value_error)?;
            planned.push((formula.sheet(), reference, cache, fields.len()));
            fields.push(formula_report(&names, formula, true));
        }
        let mut commits: Vec<WriteCommit> = Vec::new();
        for chunk in planned.chunks(CHUNK) {
            let base = commits
                .last()
                .map_or_else(|| handle.clone(), |commit| commit.after.clone());
            let used = commits.iter().fold(0usize, |total, commit| {
                total.saturating_add(commit.snapshot_bytes)
            });
            let remaining = limits.max_snapshot_bytes().saturating_sub(used);
            let chunk_limits = limits.with_snapshot_budget(remaining);
            let additional = chunk.iter().fold(0usize, |total, (_, _, cache, _)| {
                let payload = match cache {
                    FormulaCache::String(text) => text.len(),
                    _ => 0,
                };
                total
                    .saturating_add(payload)
                    .saturating_add(size_of::<NativeValue>())
            });
            base.preflight_snapshot_bytes(remaining, additional)?;
            let mut edit = base
                .editable_package_for_write_with_snapshot_limit(remaining)?
                .edit()
                .map_err(native_edit_error)?;
            let mut staged = 0usize;
            for (sheet, reference, cache, index) in chunk {
                let written = edit.set_value(
                    Selector::Position(Position::new(*sheet)),
                    *reference,
                    NativeValue::FormulaCache(cache.clone()),
                );
                if written.is_err() {
                    fields[*index] = fields[*index].clone().with_status(
                        Status::Unsupported,
                        "the formula record cannot cache this result",
                    );
                } else {
                    staged += 1;
                }
            }
            if staged == 0 {
                continue;
            }
            let native = edit.commit().map_err(native_edit_error)?;
            commits.push(finish_commit(&base, native, staged, &chunk_limits)?);
        }
        Ok((commits, Summary::from_fields(fields)))
    })();
    let freshness = handle.check_source();
    freshness.and(result)
}

/// The cached form of a computed value; `None` for one a `Formula` record
/// cannot hold.
fn cache_value(value: &Value) -> Option<FormulaCache> {
    Some(match value {
        Value::Empty => FormulaCache::Empty,
        Value::Bool(value) => FormulaCache::Boolean(*value),
        Value::Number(number) if number.is_finite() => FormulaCache::Number(*number),
        Value::Number(_) => return None,
        Value::Text(text) => FormulaCache::String(text.clone()),
        Value::Error(error) => FormulaCache::Error(
            CellError::new(match error {
                ErrorValue::Null => 0x00,
                ErrorValue::Div0 => 0x07,
                ErrorValue::Value => 0x0f,
                ErrorValue::Ref => 0x17,
                ErrorValue::Name => 0x1d,
                ErrorValue::Num => 0x24,
                ErrorValue::NotAvailable => 0x2a,
                ErrorValue::GettingData => 0x2b,
                _ => return None,
            })
            .ok()?,
        ),
    })
}
