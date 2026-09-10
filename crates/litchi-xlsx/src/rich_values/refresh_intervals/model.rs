//! Typed, inert rich-value refresh interval metadata.

use super::{MAX_INTERVALS, valid_xml10};
use crate::error::{Result, invalid};
use crate::rich_values::MAX_STRING_BYTES;

/// The `richvaluerefresh` namespace from [MS-XLSX] section 2.4.92.
pub const REFRESH_INTERVALS_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2020/richvaluerefresh";

/// One `CT_RichValueRefreshInterval` record.
///
/// The record only describes inert metadata.  It never performs a refresh or
/// dereferences a resource identifier.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RefreshInterval {
    resource_id_int: Option<i32>,
    resource_id_str: Option<String>,
    interval: i32,
}

impl RefreshInterval {
    /// Construct one interval.  At least one resource identifier is required
    /// by [MS-XLSX] section 2.6.221; an explicitly present empty string is
    /// retained because `xsd:string` permits it.
    pub fn new(
        resource_id_int: Option<i32>,
        resource_id_str: Option<String>,
        interval: i32,
    ) -> Result<Self> {
        if resource_id_int.is_none() && resource_id_str.is_none() {
            return Err(invalid(
                "rich-value refresh interval requires a resourceIdInt or resourceIdStr",
            ));
        }
        if let Some(value) = resource_id_str.as_deref() {
            validate_string(value, "resourceIdStr")?;
        }
        Ok(Self {
            resource_id_int,
            resource_id_str,
            interval,
        })
    }

    /// Numeric resource identifier, if authored.
    #[must_use]
    pub const fn resource_id_int(&self) -> Option<i32> {
        self.resource_id_int
    }

    /// String resource identifier, if authored.
    #[must_use]
    pub fn resource_id_str(&self) -> Option<&str> {
        self.resource_id_str.as_deref()
    }

    /// Refresh interval in seconds.  The values `-1` and `0` retain their
    /// protocol meanings; other values are left inert for a service to apply.
    #[must_use]
    pub const fn interval(&self) -> i32 {
        self.interval
    }
}

/// A non-empty `CT_RichValueRefreshIntervals` collection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RefreshIntervals {
    intervals: Vec<RefreshInterval>,
}

impl RefreshIntervals {
    /// Construct a bounded, non-empty collection.
    pub fn new(intervals: Vec<RefreshInterval>) -> Result<Self> {
        validate_intervals(&intervals)?;
        Ok(Self { intervals })
    }

    /// The intervals in schema order.
    #[must_use]
    pub fn intervals(&self) -> &[RefreshInterval] {
        &self.intervals
    }

    /// Number of interval records.
    #[must_use]
    pub fn len(&self) -> usize {
        self.intervals.len()
    }

    /// Whether this collection is empty.  This is always false for a valid
    /// value, but is useful when consuming a staged vector.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.intervals.is_empty()
    }

    pub(crate) fn into_intervals(self) -> Vec<RefreshInterval> {
        self.intervals
    }
}

/// The refresh metadata associated with one `CT_RichValueType` occurrence.
///
/// The position of this value in [`Snapshot::types`](super::Snapshot::types)
/// is the stable selector used by the transaction.  Type names are exposed
/// for inspection only because the wire schema does not require them to be
/// unique.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TypeRefreshIntervals {
    pub(crate) name: String,
    pub(crate) intervals: Option<RefreshIntervals>,
}

impl TypeRefreshIntervals {
    /// The `CT_RichValueType/@name` value.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The optional `refreshIntervals` payload for this type.
    #[must_use]
    pub fn intervals(&self) -> Option<&RefreshIntervals> {
        self.intervals.as_ref()
    }
}

pub(crate) fn validate_intervals(value: &[RefreshInterval]) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(
            "rich-value refreshIntervals requires at least one refreshInterval",
        ));
    }
    if value.len() > MAX_INTERVALS {
        return Err(invalid(
            "rich-value refresh interval count exceeds the limit",
        ));
    }
    for interval in value {
        if interval.resource_id_int.is_none() && interval.resource_id_str.is_none() {
            return Err(invalid(
                "rich-value refresh interval requires a resourceIdInt or resourceIdStr",
            ));
        }
        if let Some(value) = interval.resource_id_str.as_deref() {
            validate_string(value, "resourceIdStr")?;
        }
    }
    Ok(())
}

fn validate_string(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_STRING_BYTES {
        return Err(invalid(format!("{field} exceeds the string limit")));
    }
    if !value.chars().all(valid_xml10) {
        return Err(invalid(format!(
            "{field} contains an XML 1.0 invalid character"
        )));
    }
    Ok(())
}
