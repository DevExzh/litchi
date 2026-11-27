use crate::{Error, Resource, Result};

/// Number of bytes in a BIFF record header.
pub(crate) const HEADER_BYTES: usize = 4;

/// Maximum payload in one BIFF record frame.
pub const MAX_RECORD_BYTES: usize = 8_224;

const DEFAULT_MAX_RECORDS: usize = 1_000_000;
const DEFAULT_MAX_STREAM_BYTES: usize = 128 * 1024 * 1024;
const DEFAULT_MAX_OUTPUT_BYTES: usize = 128 * 1024 * 1024;

/// Resource ceilings for borrowed traversal and owned frame encoding.
///
/// The limits are checked before traversal or allocation. `max_record_bytes`
/// cannot exceed the 8,224-byte payload bound in `[MS-XLS]` section 2.1.4 and
/// `[MS-OGRAPH]` section 2.1.4. Zero is a valid restrictive limit: it permits
/// an empty stream and rejects every non-empty record or output as applicable.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of frames traversed or emitted.
    pub max_records: usize,
    /// Maximum payload bytes in one frame.
    pub max_record_bytes: usize,
    /// Maximum bytes accepted from one borrowed input stream.
    pub max_input_bytes: usize,
    /// Maximum bytes emitted by one encoder or owned frame.
    pub max_output_bytes: usize,
}

impl Limits {
    /// Validates the physical constraints of these ceilings.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] when the per-record payload ceiling is
    /// larger than the maximum representable BIFF payload.
    pub fn validate(self) -> Result<Self> {
        if self.max_record_bytes > MAX_RECORD_BYTES {
            return Err(Error::InvalidLimit {
                resource: Resource::RecordBytes,
                value: as_u64(self.max_record_bytes),
                maximum: as_u64(MAX_RECORD_BYTES),
            });
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_records: DEFAULT_MAX_RECORDS,
            max_record_bytes: MAX_RECORD_BYTES,
            max_input_bytes: DEFAULT_MAX_STREAM_BYTES,
            max_output_bytes: DEFAULT_MAX_OUTPUT_BYTES,
        }
    }
}

pub(crate) fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_limits_are_valid_and_spec_bounded() {
        let limits = Limits::default();
        assert_eq!(limits.max_record_bytes, MAX_RECORD_BYTES);
        assert_eq!(limits.validate(), Ok(limits));
    }

    #[test]
    fn zero_limits_are_valid_restrictive_bounds() {
        let limits = Limits {
            max_records: 0,
            max_record_bytes: 0,
            max_input_bytes: 0,
            max_output_bytes: 0,
        };
        assert_eq!(limits.validate(), Ok(limits));
    }

    #[test]
    fn record_limit_above_spec_maximum_is_rejected() {
        let limits = Limits {
            max_record_bytes: MAX_RECORD_BYTES + 1,
            ..Limits::default()
        };
        assert!(matches!(
            limits.validate(),
            Err(Error::InvalidLimit {
                resource: Resource::RecordBytes,
                value,
                maximum
            }) if value == as_u64(MAX_RECORD_BYTES + 1) && maximum == as_u64(MAX_RECORD_BYTES)
        ));
    }
}
