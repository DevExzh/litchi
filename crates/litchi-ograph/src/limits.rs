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
        nonzero("record bytes", self.max_record_bytes)?;
        nonzero("output bytes", self.max_output_bytes)?;

        if self.max_record_bytes > MAX_BIFF_RECORD_BYTES {
            return Err(Error::InvalidLimit {
                resource: "record bytes",
                value: as_u64(self.max_record_bytes),
                reason: "BIFF8 payloads cannot exceed 8,224 bytes",
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
