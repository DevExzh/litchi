use crate::{Error, Result};

/// Maximum payload size of a BIFF8 record, in bytes.
pub const MAX_BIFF_RECORD_BYTES: usize = 8_224;

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
    /// Maximum number of BIFF records traversed or encoded.
    pub max_records: usize,
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
    /// Maximum BIFF payload size, never greater than 8,224.
    pub max_record_bytes: usize,
    /// Maximum size of an encoded BIFF stream.
    pub max_output_bytes: usize,
}

impl Limits {
    /// Checks that all configured limits are usable and mutually consistent.
    pub fn validate(self) -> Result<Self> {
        nonzero("package bytes", self.max_package_bytes)?;
        nonzero("root entries", self.max_streams)?;
        nonzero("stream bytes", self.max_stream_bytes)?;
        nonzero("Workbook bytes", self.max_workbook_bytes)?;
        nonzero("record count", self.max_records)?;
        nonzero("chart count", self.max_charts)?;
        nonzero("chart record count", self.max_chart_records)?;
        nonzero("series count", self.max_series)?;
        nonzero("group count", self.max_groups)?;
        nonzero("axis count", self.max_axes)?;
        nonzero("formula bytes", self.max_formula_bytes)?;
        nonzero("cached value count", self.max_cached_values)?;
        nonzero("unknown bytes", self.max_unknown_bytes)?;
        nonzero("chart nesting", self.max_nesting)?;
        nonzero("record bytes", self.max_record_bytes)?;
        nonzero("output bytes", self.max_output_bytes)?;

        if self.max_record_bytes > MAX_BIFF_RECORD_BYTES {
            return Err(Error::InvalidLimit {
                resource: "record bytes",
                value: as_u64(self.max_record_bytes),
                reason: "BIFF8 payloads cannot exceed 8,224 bytes",
            });
        }
        if self.max_groups > 10 {
            return Err(Error::InvalidLimit {
                resource: "group count",
                value: as_u64(self.max_groups),
                reason: "BIFF charts cannot contain more than ten groups",
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
            max_records: 1_000_000,
            max_charts: 512,
            max_chart_records: 65_536,
            max_series: 255,
            max_groups: 10,
            max_axes: 6,
            max_formula_bytes: MAX_BIFF_RECORD_BYTES - 8,
            max_cached_values: 32_000,
            max_unknown_bytes: 16 * 1024 * 1024,
            max_nesting: 128,
            max_record_bytes: MAX_BIFF_RECORD_BYTES,
            max_output_bytes: 128 * 1024 * 1024,
        }
    }
}

fn nonzero(resource: &'static str, value: usize) -> Result<()> {
    if value == 0 {
        return Err(Error::InvalidLimit {
            resource,
            value: 0,
            reason: "must be non-zero",
        });
    }
    Ok(())
}

pub(crate) fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
