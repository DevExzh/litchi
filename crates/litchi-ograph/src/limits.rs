use crate::{Error, LimitKind, Result};
use litchi_biff::{Limits as BiffLimits, MAX_RECORD_BYTES};

/// Ingress options for a standalone Graph package and its projections.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct OpenOptions {
    limits: Limits,
}

impl OpenOptions {
    /// Conservative default ingress bounds.
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the bounds used to validate the package.
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// The selected ingress bounds.
    pub const fn limits_value(&self) -> Limits {
        self.limits
    }
}

/// Resource bounds applied before allocation or traversal.
///
/// Fields are public so callers can concisely override a few values with
/// struct-update syntax. Every public parser and encoder validates the whole
/// value before using it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum compound-file size.
    pub max_package_bytes: usize,
    /// Maximum direct children allowed below the CFB root.
    pub max_streams: usize,
    /// Maximum size of any allowed root stream.
    pub max_stream_bytes: usize,
    /// Maximum size of the required `Workbook` stream.
    pub max_workbook_bytes: usize,
    /// Physical BIFF frame bounds shared with every BIFF consumer.
    pub biff: BiffLimits,
    /// Maximum chart substreams discovered in one Workbook stream.
    pub max_charts: usize,
    /// Maximum records contained in one chart substream.
    pub max_chart_records: usize,
    /// Maximum semantic series in one chart.
    pub max_series: usize,
    /// Maximum chart groups in one chart, never greater than ten.
    pub max_groups: usize,
    /// Maximum axes in one chart.
    pub max_axes: usize,
    /// Maximum bytes retained for one inert formula token array.
    pub max_formula_bytes: usize,
    /// Maximum cached values in one chart.
    pub max_cached_values: usize,
    /// Maximum aggregate bytes retained for unknown records.
    pub max_unknown_bytes: usize,
    /// Maximum nesting depth of Begin/End record collections.
    pub max_nesting: usize,
}

impl Limits {
    /// Checks that all configured limits are usable and mutually consistent.
    // The record codecs repeatedly validate retained policies. Keep the
    // successful scalar checks inline even though error conversion is cold.
    #[inline(always)]
    pub fn validate(self) -> Result<Self> {
        nonzero(LimitKind::PackageBytes, self.max_package_bytes)?;
        nonzero(LimitKind::RootEntries, self.max_streams)?;
        nonzero(LimitKind::StreamBytes, self.max_stream_bytes)?;
        nonzero(LimitKind::WorkbookBytes, self.max_workbook_bytes)?;
        nonzero(LimitKind::Charts, self.max_charts)?;
        nonzero(LimitKind::ChartRecords, self.max_chart_records)?;
        nonzero(LimitKind::Series, self.max_series)?;
        nonzero(LimitKind::Groups, self.max_groups)?;
        nonzero(LimitKind::Axes, self.max_axes)?;
        nonzero(LimitKind::FormulaBytes, self.max_formula_bytes)?;
        nonzero(LimitKind::CachedValues, self.max_cached_values)?;
        nonzero(LimitKind::UnknownBytes, self.max_unknown_bytes)?;
        nonzero(LimitKind::Nesting, self.max_nesting)?;
        self.biff.validate()?;
        if self.max_groups > 10 {
            return Err(Error::InvalidLimit {
                kind: LimitKind::Groups,
                value: as_u64(self.max_groups),
                maximum: 10,
            });
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_package_bytes: 256 * 1024 * 1024,
            max_streams: 3,
            max_stream_bytes: 128 * 1024 * 1024,
            max_workbook_bytes: 128 * 1024 * 1024,
            biff: BiffLimits {
                max_records: 1_000_000,
                max_record_bytes: MAX_RECORD_BYTES,
                max_input_bytes: 128 * 1024 * 1024,
                max_output_bytes: 128 * 1024 * 1024,
            },
            max_charts: 512,
            max_chart_records: 65_536,
            max_series: 255,
            max_groups: 10,
            max_axes: 6,
            max_formula_bytes: MAX_RECORD_BYTES - 8,
            max_cached_values: 32_000,
            max_unknown_bytes: 16 * 1024 * 1024,
            max_nesting: 128,
        }
    }
}

fn nonzero(kind: LimitKind, value: usize) -> Result<()> {
    if value == 0 {
        return Err(Error::InvalidLimit {
            kind,
            value: 0,
            maximum: if kind == LimitKind::Groups {
                10
            } else {
                as_u64(usize::MAX)
            },
        });
    }
    Ok(())
}

pub(crate) fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn invalid_configuration_reports_the_dimension_and_supported_ceiling() {
        let cases = [
            (
                LimitKind::PackageBytes,
                Limits {
                    max_package_bytes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::RootEntries,
                Limits {
                    max_streams: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::StreamBytes,
                Limits {
                    max_stream_bytes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::WorkbookBytes,
                Limits {
                    max_workbook_bytes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::Charts,
                Limits {
                    max_charts: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::ChartRecords,
                Limits {
                    max_chart_records: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::Series,
                Limits {
                    max_series: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::Groups,
                Limits {
                    max_groups: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::Axes,
                Limits {
                    max_axes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::FormulaBytes,
                Limits {
                    max_formula_bytes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::CachedValues,
                Limits {
                    max_cached_values: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::UnknownBytes,
                Limits {
                    max_unknown_bytes: 0,
                    ..Limits::default()
                },
            ),
            (
                LimitKind::Nesting,
                Limits {
                    max_nesting: 0,
                    ..Limits::default()
                },
            ),
        ];
        for (expected, limits) in cases {
            let ceiling = if expected == LimitKind::Groups {
                10
            } else {
                usize::MAX as u64
            };
            assert!(matches!(limits.validate(), Err(Error::InvalidLimit {
                kind, value: 0, maximum,
            }) if kind == expected && maximum == ceiling));
        }
        assert!(matches!(
            Limits {
                max_groups: 11,
                ..Limits::default()
            }
            .validate(),
            Err(Error::InvalidLimit {
                kind: LimitKind::Groups,
                value: 11,
                maximum: 10
            })
        ));
    }

    #[test]
    fn physical_configuration_keeps_its_biff_dimension_and_numeric_evidence() {
        let limits = Limits {
            biff: BiffLimits {
                max_record_bytes: MAX_RECORD_BYTES + 1,
                ..BiffLimits::default()
            },
            ..Limits::default()
        };
        assert!(matches!(limits.validate(), Err(Error::InvalidLimit {
            kind: LimitKind::Biff(litchi_biff::Resource::RecordBytes), value, maximum,
        }) if value == (MAX_RECORD_BYTES + 1) as u64 && maximum == MAX_RECORD_BYTES as u64));
    }
}
