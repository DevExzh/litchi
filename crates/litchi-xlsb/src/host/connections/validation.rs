//! Semantic and source-owned bounds for inert external connections.

use super::{Connections, UnknownRecord};
use crate::package::error::Result;

/// Validate the typed connection/query/parameter model.
pub fn connections(value: &Connections) -> Result<()> {
    super::write::validate_connections(value)
}

/// Validate opaque records before they are merged into an authored stream.
pub(crate) fn unknown_records(records: &[UnknownRecord]) -> Result<()> {
    if records.len() > super::write::MAX_UNKNOWN_RECORDS {
        return Err(super::invalid("opaque record count limit exceeded"));
    }
    let mut total = 0usize;
    for record in records {
        let (header, header_len) = crate::raw::Header::parse(
            record.bytes(),
            crate::raw::Limits::new(super::write::MAX_CONNECTIONS_PART_BYTES, 1_048_576),
        )?;
        if header_len.checked_add(header.len()) != Some(record.bytes().len())
            || u16::from(header.kind()) != record.kind()
            || super::codec::is_modeled_record(header.kind())
        {
            return Err(super::invalid("invalid opaque record envelope"));
        }
        total = total
            .checked_add(record.bytes().len())
            .ok_or_else(|| super::invalid("opaque byte count overflow"))?;
        if total > super::write::MAX_UNKNOWN_BYTES {
            return Err(super::invalid("opaque byte limit exceeded"));
        }
    }
    Ok(())
}

/// Validate the complete workbook-level connections graph.
pub fn graph(package: &litchi_opc::OpcPackage) -> Result<Option<Connections>> {
    super::package::validate_graph(package)
}
