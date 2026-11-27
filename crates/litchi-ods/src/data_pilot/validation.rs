//! Contextual `DataPilot` validation and preservation gates.

use crate::model::data_pilot::{self, Table};
use litchi_core::{Error, Result};

use super::codec::Location;

const MAX_SOURCE_BYTES: usize = 64 * 1024 * 1024;

pub(crate) fn validate_snapshot(
    source_xml: &str,
    location: &Location,
    tables: &[Table],
) -> Result<()> {
    if source_xml.len() > MAX_SOURCE_BYTES {
        return Err(Error::InvalidFormat(
            "ODS DataPilot source exceeds the snapshot limit".to_string(),
        ));
    }
    data_pilot::validate_data_pilot_tables(tables)?;
    if location.container.is_none() && !tables.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS DataPilot declarations have no physical owner".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_candidate(
    location: &Location,
    original: &Option<Vec<Table>>,
    candidate: &Option<Vec<Table>>,
) -> Result<()> {
    if let Some(tables) = candidate {
        data_pilot::validate_data_pilot_tables(tables)?;
    }
    if original != candidate && location.opaque {
        return Err(Error::InvalidFormat(
            "ODS DataPilot contains unknown XML; refusing a lossy edit".to_string(),
        ));
    }
    Ok(())
}
