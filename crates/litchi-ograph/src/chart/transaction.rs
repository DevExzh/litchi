//! Snapshot-scoped, lossless chart transactions.
//!
//! The transaction deliberately owns a parsed [`Chart`] snapshot. It can
//! replace an existing cache value or the fixed-size chart-area rectangle only
//! when the physical BIFF record grammar remains unchanged. It can also patch
//! the existing fixed-size `ShtProps` metadata record without changing the
//! structural `PlotArea` record. Commit therefore patches the retained stream
//! in place and never reconstructs the surrounding chart grammar.

pub mod chart_area;
mod model;
pub mod sheet_props;
mod validation;

use super::{Chart, Error, Rect, Result, codec};
pub(crate) use model::Request;
pub use model::{CacheValue, Change, Commit, Identity, Patch};

/// A bounded semantic editor for one parsed chart snapshot.
#[derive(Debug)]
pub struct Editor {
    chart: Chart,
    requests: Vec<Request>,
    chart_area: Option<chart_area::Request>,
    sheet_props: Option<sheet_props::Request>,
}

impl Editor {
    pub(crate) fn new(chart: Chart) -> Result<Self> {
        validation::ensure_editable(&chart)?;
        Ok(Self {
            chart,
            requests: Vec::new(),
            chart_area: None,
            sheet_props: None,
        })
    }

    /// Number of distinct semantic operations currently staged.
    #[must_use]
    pub fn len(&self) -> usize {
        self.requests
            .len()
            .saturating_add(usize::from(self.chart_area.is_some()))
            .saturating_add(usize::from(self.sheet_props.is_some()))
    }

    /// Whether this transaction has no staged semantic operation.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.requests.is_empty() && self.chart_area.is_none() && self.sheet_props.is_none()
    }

    /// Stage a replacement for one existing cache cell.
    ///
    /// The cell's coordinates, section, and number-format identity are not
    /// arguments: they are carried by the source snapshot and cannot be
    /// changed through this capability. A replacement must use the same
    /// producer-specific physical cache record class. Text replacements also
    /// have to fit the original record byte length; this keeps opaque record
    /// placement and every following record offset stable.
    pub fn set_cache_value<V>(&mut self, index: usize, value: V) -> Result<&mut Self>
    where
        V: Into<CacheValue>,
    {
        let value = value.into();
        let cache = self.chart.caches.get(index).ok_or(Error::InvalidModel {
            field: "cache",
            reason: "cache index is outside the parsed chart",
        })?;
        validation::ensure_value(cache, &value)?;

        if let Some(request) = self
            .requests
            .iter_mut()
            .find(|request| request.index == index)
        {
            request.value = value;
            return Ok(self);
        }

        if self.requests.len() >= self.chart.limits.max_cached_values {
            return Err(Error::LimitExceeded {
                resource: "cache edit count",
                observed: u64::try_from(self.requests.len().saturating_add(1)).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.chart.limits.max_cached_values).unwrap_or(u64::MAX),
            });
        }
        self.requests
            .try_reserve(1)
            .map_err(|_| Error::Allocation {
                resource: "cache edits",
            })?;
        self.requests.push(Request { index, value });
        Ok(self)
    }

    /// Stage a replacement for the fixed-size `[MS-OGRAPH]` `Chart` area.
    ///
    /// The wire record stores four signed 16.16 fixed-point values. The
    /// specification requires the chart-area origin to be zero and its width
    /// and height to be nonnegative. The edit keeps the record kind, payload
    /// length, location, and every surrounding unknown record unchanged.
    pub fn set_rect(&mut self, value: Rect) -> Result<&mut Self> {
        chart_area::validation::ensure_pair(self.chart.rect, value)?;
        if let Some(request) = self.chart_area.as_mut() {
            request.value = value;
        } else {
            self.chart_area = Some(chart_area::Request { value });
        }
        Ok(self)
    }

    /// Stage a replacement for the existing four-byte `[MS-OGRAPH]`
    /// `ShtProps` record.
    ///
    /// The typed chart properties include the defined flags, blank-cell mode,
    /// and the parsed `PlotArea` presence. The latter is source topology and
    /// must remain unchanged: this operation never inserts, removes, or
    /// reorders the empty `PlotArea` record.
    pub fn set_props(&mut self, value: super::Props) -> Result<&mut Self> {
        sheet_props::validation::ensure_pair(self.chart.props(), value)?;
        if let Some(request) = self.sheet_props.as_mut() {
            request.value = value;
        } else {
            self.sheet_props = Some(sheet_props::Request {
                value,
                expected_offset: None,
            });
        }
        Ok(self)
    }

    /// Validate and publish the staged chart edits.
    ///
    /// The returned [`Commit`] owns the post-edit chart snapshot and a
    /// reversible semantic patch. The source chart passed to [`Chart::edit`]
    /// was consumed and is never mutated by reference.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            mut chart,
            requests,
            chart_area,
            sheet_props,
        } = self;
        let mut effective = Vec::new();
        effective
            .try_reserve_exact(requests.len())
            .map_err(|_| Error::Allocation {
                resource: "effective cache edits",
            })?;

        let mut changes = Vec::new();
        changes
            .try_reserve_exact(requests.len())
            .map_err(|_| Error::Allocation {
                resource: "chart patch changes",
            })?;
        for request in requests {
            let cache = chart.caches.get(request.index).ok_or(Error::InvalidModel {
                field: "cache",
                reason: "cache index disappeared during transaction",
            })?;
            let before = CacheValue::from_cache(cache);
            if before == request.value {
                continue;
            }
            let identity = Identity::from_cache(cache);
            changes.push(Change::new(
                request.index,
                identity,
                before,
                request.value.clone(),
            ));
            effective.push(request);
        }

        let chart_area = chart_area.and_then(|request| {
            (chart.rect != request.value)
                .then(|| chart_area::Change::new(chart.rect, request.value))
        });
        let sheet_props = match sheet_props {
            Some(request) if chart.props != request.value => {
                let before = chart.props;
                let offset = sheet_props::codec::locate(&chart, before)?;
                if request.expected_offset.is_some_and(|value| value != offset) {
                    return Err(Error::UnsupportedMutation {
                        operation: "sheet-props-patch",
                        reason: "source ShtProps record identity does not match the patch",
                    });
                }
                Some(sheet_props::Change::new(before, request.value, offset))
            },
            _ => None,
        };

        // Validate every source seam before the first byte is changed. The
        // physical cache patcher performs the same complete preflight before
        // applying its prepared payloads.
        if let Some(change) = &chart_area {
            chart_area::codec::locate(&chart, change.before())?;
        }
        if let Some(change) = &sheet_props {
            sheet_props::codec::locate(&chart, change.before())?;
        }

        if !effective.is_empty() {
            codec::patch(&mut chart, &effective)?;
            for request in effective {
                let cache = chart
                    .caches
                    .get_mut(request.index)
                    .ok_or(Error::InvalidModel {
                        field: "cache",
                        reason: "cache index disappeared while applying the patch",
                    })?;
                CacheValue::replace_cache(cache, request.value)?;
            }
        }

        if let Some(change) = &chart_area {
            chart_area::codec::patch(&mut chart, change.before(), change.after())?;
            chart.rect = change.after();
        }

        if let Some(change) = &sheet_props {
            sheet_props::codec::patch(
                &mut chart,
                change.before(),
                change.after(),
                change.offset(),
            )?;
            chart.props = change.after();
        }

        Ok(Commit::new(
            chart,
            Patch::new(changes, chart_area, sheet_props),
        ))
    }

    fn set_expected(&mut self, change: &Change) -> Result<&mut Self> {
        let cache = self
            .chart
            .caches
            .get(change.index())
            .ok_or(Error::InvalidModel {
                field: "cache",
                reason: "patch cache index is outside the target chart",
            })?;
        validation::ensure_identity(cache, *change.identity())?;
        if CacheValue::from_cache(cache) != *change.before() {
            return Err(Error::UnsupportedMutation {
                operation: "cache-value-patch",
                reason: "patch source value does not match the target snapshot",
            });
        }
        self.set_cache_value(change.index(), change.after().clone())
    }

    fn set_expected_chart_area(&mut self, change: &chart_area::Change) -> Result<&mut Self> {
        if self.chart.rect != change.before() {
            return Err(Error::UnsupportedMutation {
                operation: "chart-area-patch",
                reason: "patch source rectangle does not match the target snapshot",
            });
        }
        self.set_rect(change.after())
    }

    fn set_expected_sheet_props(&mut self, change: &sheet_props::Change) -> Result<&mut Self> {
        if self.chart.props() != change.before() {
            return Err(Error::UnsupportedMutation {
                operation: "sheet-props-patch",
                reason: "patch source ShtProps value does not match the target snapshot",
            });
        }
        sheet_props::validation::ensure_pair(self.chart.props(), change.after())?;
        self.sheet_props = Some(sheet_props::Request {
            value: change.after(),
            expected_offset: Some(change.offset()),
        });
        Ok(self)
    }
}

impl Patch {
    /// Applies this patch only to a matching parsed chart snapshot.
    pub fn apply(&self, chart: Chart) -> Result<Commit> {
        let mut editor = chart.edit()?;
        for change in self.changes() {
            editor.set_expected(change)?;
        }
        if let Some(change) = self.chart_area() {
            editor.set_expected_chart_area(change)?;
        }
        if let Some(change) = self.sheet_props() {
            editor.set_expected_sheet_props(change)?;
        }
        editor.commit()
    }
}

#[cfg(test)]
mod tests;
