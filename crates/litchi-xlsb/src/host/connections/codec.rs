//! Source-preserving BIFF12 envelopes for external connections.

use std::sync::Arc;

use super::model::UnknownRecord;
use super::{parse_connections_part, write};
use crate::package::error::Result;

/// Parsed connection data plus the exact source and opaque records skipped by
/// the semantic parser.
#[derive(Debug, Clone)]
pub(crate) struct Source {
    pub(crate) unknown_records: Vec<UnknownRecord>,
}

/// Parse a complete connections stream while retaining future records.
pub(crate) fn parse_source(data: &[u8]) -> Result<Source> {
    parse_connections_part(data)?;
    let mut unknown_records = Vec::new();
    let mut owner = None;
    let mut root_known = 0usize;
    let mut local_known = 0usize;
    let mut opaque_bytes = 0usize;

    for record in crate::raw::Records::with_limits(
        data,
        crate::raw::Limits::new(write::MAX_CONNECTIONS_PART_BYTES, 1_048_576),
    ) {
        let record = record?;
        let kind = record.kind();
        if kind == crate::raw::kind::BEGIN_EXT_CONNECTION {
            root_known = root_known.saturating_add(1);
            owner = connection_id(record.payload());
            local_known = 0;
            continue;
        }
        if is_modeled_record(kind) {
            if owner.is_some() {
                local_known = local_known.saturating_add(1);
                if kind == crate::raw::kind::END_EXT_CONNECTION {
                    owner = None;
                    local_known = 0;
                }
            } else {
                root_known = root_known.saturating_add(1);
            }
            continue;
        }

        let offset = record.offset();
        let (_, header_len) = crate::raw::Header::parse(
            &data[offset..],
            crate::raw::Limits::new(write::MAX_CONNECTIONS_PART_BYTES, 1_048_576),
        )?;
        let end = offset
            .checked_add(header_len)
            .and_then(|value| value.checked_add(record.len()))
            .ok_or_else(|| crate::package::error::Error::Unrecognized {
                typ: "ExternalDataConnections".to_string(),
                val: "opaque record range overflow".to_string(),
            })?;
        let bytes: Arc<[u8]> = Arc::from(data[offset..end].to_vec());
        opaque_bytes = opaque_bytes.saturating_add(bytes.len());
        if unknown_records.len() >= write::MAX_UNKNOWN_RECORDS
            || opaque_bytes > write::MAX_UNKNOWN_BYTES
        {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "ExternalDataConnections".to_string(),
                val: "opaque record bound exceeded".to_string(),
            });
        }
        unknown_records.push(UnknownRecord::new(
            u16::from(kind),
            owner,
            if owner.is_some() {
                local_known
            } else {
                root_known
            },
            bytes,
            header_len,
        ));
    }

    Ok(Source { unknown_records })
}

pub(crate) fn is_modeled_record(kind: crate::raw::Kind) -> bool {
    matches!(
        kind,
        crate::raw::kind::BEGIN_EXT_CONNECTIONS
            | crate::raw::kind::END_EXT_CONNECTIONS
            | crate::raw::kind::BEGIN_EXT_CONNECTION
            | crate::raw::kind::END_EXT_CONNECTION
            | crate::raw::kind::BEGIN_EC_DB_PROPS
            | crate::raw::kind::END_EC_DB_PROPS
            | crate::raw::kind::BEGIN_EC_OLAP_PROPS
            | crate::raw::kind::END_EC_OLAP_PROPS
            | crate::raw::kind::BEGIN_EC_WEB_PROPS
            | crate::raw::kind::END_EC_WEB_PROPS
            | crate::raw::kind::BEGIN_EC_PARAMS
            | crate::raw::kind::END_EC_PARAMS
            | crate::raw::kind::BEGIN_EC_PARAM
            | crate::raw::kind::END_EC_PARAM
            | crate::raw::kind::BEGIN_EC_WP_TABLES
            | crate::raw::kind::END_EC_WP_TABLES
            | crate::raw::kind::PCDI_MISSING
            | crate::raw::kind::PCDI_STRING
            | crate::raw::kind::PCDI_INDEX
    )
}

fn connection_id(payload: &[u8]) -> Option<u32> {
    payload
        .get(18..22)
        .and_then(|bytes| Some(u32::from_le_bytes(bytes.try_into().ok()?)))
}
