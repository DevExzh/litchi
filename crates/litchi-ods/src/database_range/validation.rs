//! Contextual database-range validation and preservation gates.

use crate::model::database_range::{self as vocabulary, Range, database_ranges_serialized_size};
use litchi_core::{Error, Result};

use super::codec::Location;

const MAX_SOURCE_BYTES: usize = 64 * 1024 * 1024;
const MAX_SERIALIZED_BYTES: usize = 32 * 1024 * 1024;

pub(crate) fn validate_snapshot(
    source_xml: &str,
    location: &Location,
    ranges: &[Range],
) -> Result<()> {
    if source_xml.len() > MAX_SOURCE_BYTES {
        return Err(Error::InvalidFormat(
            "ODS database-range source exceeds the snapshot limit".to_string(),
        ));
    }
    vocabulary::validate_database_range_collection(ranges)?;
    if location.container.is_none() && !ranges.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS database-range declarations have no physical owner".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_candidate(
    location: &Location,
    original: &Option<Vec<Range>>,
    candidate: &Option<Vec<Range>>,
) -> Result<()> {
    if let Some(ranges) = candidate {
        let serialized = database_ranges_serialized_size(ranges)?;
        if serialized > MAX_SERIALIZED_BYTES {
            return Err(Error::InvalidFormat(
                "database-range serialized output exceeds the safety limit".to_string(),
            ));
        }
    }
    if original != candidate && location.opaque {
        return Err(Error::InvalidFormat(
            "ODS database-ranges contain opaque or unsupported XML; refusing a lossy edit"
                .to_string(),
        ));
    }
    Ok(())
}
