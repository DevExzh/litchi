//! Transaction-local semantic guards.

use super::super::{Cache, Chart, Value, XlValue};
use super::model::{CacheValue, Identity};
use crate::{Error, Result};

pub(super) fn ensure_editable(chart: &Chart) -> Result<()> {
    if chart.is_pristine() {
        Ok(())
    } else {
        Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "only a pristine parsed chart has a replayable source stream",
        })
    }
}

pub(super) fn ensure_value(cache: &Cache, value: &CacheValue) -> Result<()> {
    match (cache, value) {
        (Cache::Graph { value: current, .. }, CacheValue::Graph(replacement)) => {
            match (current, replacement) {
                (Value::Number(_), Value::Number(value)) => ensure_xnum(*value),
                (Value::Text(_), Value::Text(_)) | (Value::Blank, Value::Blank) => Ok(()),
                _ => same_wire_class(),
            }
        },
        (Cache::Excel { value: current, .. }, CacheValue::Excel(replacement)) => {
            match (current, replacement) {
                (XlValue::Number(_), XlValue::Number(value)) => ensure_xnum(*value),
                (XlValue::Text(_), XlValue::Text(_)) | (XlValue::Blank, XlValue::Blank) => Ok(()),
                (XlValue::Bool(_), XlValue::Bool(_))
                | (XlValue::Bool(_), XlValue::Error(_))
                | (XlValue::Error(_), XlValue::Bool(_))
                | (XlValue::Error(_), XlValue::Error(_)) => Ok(()),
                _ => same_wire_class(),
            }
        },
        (Cache::Graph { .. }, CacheValue::Excel(_))
        | (Cache::Excel { .. }, CacheValue::Graph(_)) => Err(Error::InvalidModel {
            field: "cache",
            reason: "replacement producer does not match the chart cache",
        }),
    }
}

pub(super) fn ensure_identity(cache: &Cache, identity: Identity) -> Result<()> {
    if Identity::from_cache(cache) == identity {
        Ok(())
    } else {
        Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "patch cache identity does not match the target snapshot",
        })
    }
}

fn same_wire_class<T>() -> Result<T> {
    Err(Error::UnsupportedMutation {
        operation: "cache-value-patch",
        reason: "replacement would change the physical cache record class",
    })
}

fn ensure_xnum(value: f64) -> Result<()> {
    if !value.is_finite() || value.is_subnormal() || (value == 0.0 && value.is_sign_negative()) {
        return Err(Error::InvalidModel {
            field: "cache value",
            reason: "Xnum must be finite, normalized, non-NaN, and not negative zero",
        });
    }
    Ok(())
}
