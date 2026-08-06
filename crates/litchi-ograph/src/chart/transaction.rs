//! Snapshot-scoped, lossless chart cache transactions.
//!
//! The transaction deliberately owns a parsed [`Chart`] snapshot. It can
//! replace the value of an existing cache cell only when the physical BIFF
//! record kind and payload length remain unchanged. Commit therefore patches
//! the retained stream in place and never reconstructs the surrounding chart
//! grammar.

mod model;
mod validation;

use super::{Chart, Error, Result, codec};
pub(crate) use model::Request;
pub use model::{CacheValue, Change, Commit, Identity, Patch};

/// A bounded semantic editor for one parsed chart snapshot.
#[derive(Debug)]
pub struct Editor {
    chart: Chart,
    requests: Vec<Request>,
}

impl Editor {
    pub(crate) fn new(chart: Chart) -> Result<Self> {
        validation::ensure_editable(&chart)?;
        Ok(Self {
            chart,
            requests: Vec::new(),
        })
    }

    /// Number of distinct cache cells currently staged.
    #[must_use]
    pub fn len(&self) -> usize {
        self.requests.len()
    }

    /// Whether this transaction has no staged cache replacement.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.requests.is_empty()
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

    /// Validate and publish the staged cache replacements.
    ///
    /// The returned [`Commit`] owns the post-edit chart snapshot and a
    /// reversible semantic patch. The source chart passed to [`Chart::edit`]
    /// was consumed and is never mutated by reference.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            mut chart,
            requests,
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

        Ok(Commit::new(chart, Patch::new(changes)))
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
}

impl Patch {
    /// Applies this patch only to a matching parsed chart snapshot.
    pub fn apply(&self, chart: Chart) -> Result<Commit> {
        let mut editor = chart.edit()?;
        for change in self.changes() {
            editor.set_expected(change)?;
        }
        editor.commit()
    }
}

#[cfg(test)]
mod tests;
