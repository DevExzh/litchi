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
pub(crate) fn parse_source(data: &[u8], raw_limits: crate::raw::Limits) -> Result<Source> {
    parse_connections_part(data, raw_limits)?;
    let mut unknown_records = Vec::new();
    let mut owner = None;
    let mut root_known = 0usize;
    let mut local_known = 0usize;
    let mut opaque_bytes = 0usize;
    let mut opaque_ends = Vec::new();

    for record in crate::raw::Records::new(
        data,
        crate::raw::Limits::new(
            write::MAX_CONNECTIONS_PART_BYTES.min(raw_limits.payload()),
            raw_limits.string_units().min(1_048_576),
        ),
    ) {
        let record = record?;
        let kind = record.kind();
        let was_opaque = !opaque_ends.is_empty();
        let starts_opaque = crate::package::walker::paired_end(kind)
            .filter(|_| was_opaque || !is_modeled_record(kind));
        let opaque_context = was_opaque || starts_opaque.is_some();
        if opaque_ends.last() == Some(&kind) {
            opaque_ends.pop();
        } else if let Some(end) = starts_opaque {
            if opaque_ends.len() >= 256 {
                return Err(super::invalid(
                    "opaque connection collection depth exceeded",
                ));
            }
            opaque_ends.push(end);
        }
        if !opaque_context && kind == crate::raw::kind::BEGIN_EXT_CONNECTION {
            root_known = root_known.saturating_add(1);
            owner = connection_id(record.payload());
            local_known = 0;
            continue;
        }
        if !opaque_context && is_modeled_record(kind) {
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
            crate::raw::Limits::new(
                write::MAX_CONNECTIONS_PART_BYTES.min(raw_limits.payload()),
                raw_limits.string_units().min(1_048_576),
            ),
        )?;
        let end = offset
            .checked_add(header_len)
            .and_then(|value| value.checked_add(record.len()))
            .ok_or(crate::package::error::Error::Unrecognized {
                typ: "ExternalDataConnections",
                val: "opaque record range overflow",
            })?;
        let bytes: Arc<[u8]> = Arc::from(data[offset..end].to_vec());
        opaque_bytes = opaque_bytes.saturating_add(bytes.len());
        if unknown_records.len() >= write::MAX_UNKNOWN_RECORDS
            || opaque_bytes > write::MAX_UNKNOWN_BYTES
        {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "ExternalDataConnections",
                val: "opaque record bound exceeded",
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
            opaque_context,
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
            | crate::raw::kind::BEGIN_EC_TXT_WIZ
            | crate::raw::kind::END_EC_TXT_WIZ
            | crate::raw::kind::BEGIN_EC_TW_FLD_INFO_LIST
            | crate::raw::kind::END_EC_TW_FLD_INFO_LIST
            | crate::raw::kind::BEGIN_EC_TW_FLD_INFO
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

/// Reinsert each opaque record at its retained modeled-record ordinal.
/// Root positions are published only between complete connection collections.
pub(crate) fn restore_opaque_positions(
    data: &[u8],
    unknown: &[UnknownRecord],
    limits: crate::raw::Limits,
) -> Result<Vec<u8>> {
    use std::collections::{BTreeMap, BTreeSet};
    let mut ids = BTreeSet::new();
    for record in crate::raw::Records::new(data, limits) {
        let record = record?;
        if record.kind() == crate::raw::kind::BEGIN_EXT_CONNECTION {
            ids.insert(
                connection_id(record.payload())
                    .ok_or_else(|| super::invalid("connection ID missing"))?,
            );
        }
    }
    let mut positions = BTreeMap::<(Option<u32>, usize), Vec<&UnknownRecord>>::new();
    for record in unknown {
        if record.connection_id().is_none_or(|id| ids.contains(&id)) {
            positions
                .entry((record.connection_id(), record.after_known()))
                .or_default()
                .push(record);
        }
    }
    let mut output = Vec::new();
    let emit = |output: &mut Vec<u8>, entries: Option<Vec<&UnknownRecord>>| -> Result<()> {
        for entry in entries.into_iter().flatten() {
            if output.len().saturating_add(entry.bytes().len()) > write::MAX_CONNECTIONS_PART_BYTES
            {
                return Err(super::invalid(
                    "serialized connections part size limit exceeded",
                ));
            }
            output.extend_from_slice(entry.bytes());
        }
        Ok(())
    };
    emit(&mut output, positions.remove(&(None, 0)))?;
    let mut owner = None;
    let mut root = 0;
    let mut local = 0;
    let mut opaque_ends = Vec::new();
    for record in crate::raw::Records::new(data, limits) {
        let record = record?;
        let (_, header) = crate::raw::Header::parse(&data[record.offset()..], limits)?;
        let end = record.offset() + header + record.len();
        if output.len().saturating_add(end - record.offset()) > write::MAX_CONNECTIONS_PART_BYTES {
            return Err(super::invalid(
                "serialized connections part size limit exceeded",
            ));
        }
        output.extend_from_slice(&data[record.offset()..end]);
        let kind = record.kind();
        if !opaque_ends.is_empty() || !is_modeled_record(kind) {
            if opaque_ends.last() == Some(&kind) {
                opaque_ends.pop();
            } else if let Some(end) = crate::package::walker::paired_end(kind) {
                opaque_ends.push(end);
            }
            continue;
        }
        if record.kind() == crate::raw::kind::BEGIN_EXT_CONNECTION {
            root += 1;
            owner = connection_id(record.payload());
            local = 0;
            emit(&mut output, positions.remove(&(owner, 0)))?;
        } else if owner.is_some() {
            local += 1;
            emit(&mut output, positions.remove(&(owner, local)))?;
            if record.kind() == crate::raw::kind::END_EXT_CONNECTION {
                owner = None;
                local = 0;
                emit(&mut output, positions.remove(&(None, root)))?;
            }
        } else {
            root += 1;
            emit(&mut output, positions.remove(&(None, root)))?;
        }
    }
    if !positions.is_empty() {
        return Err(super::invalid(
            "opaque connection record has no retained position",
        ));
    }
    Ok(output)
}

/// Preserve structural context when unknown records constrain an owner's layout.
/// Scalar changes are allowed; changing that scope's known record sequence is not.
pub(crate) fn check_opaque_layout(
    before: &[u8],
    after: &[u8],
    unknown: &[UnknownRecord],
    limits: crate::raw::Limits,
) -> Result<()> {
    use std::collections::{BTreeMap, BTreeSet};
    let scopes = unknown
        .iter()
        .map(UnknownRecord::connection_id)
        .collect::<BTreeSet<_>>();
    if scopes.is_empty() {
        return Ok(());
    }
    let layout = |data: &[u8]| -> Result<BTreeMap<Option<u32>, Vec<(u16, Option<u32>)>>> {
        let opaque = parse_source(data, limits)?.unknown_records;
        let mut retained = opaque.iter().peekable();
        let mut owner = None;
        let mut result = BTreeMap::<Option<u32>, Vec<(u16, Option<u32>)>>::new();
        // Match exact opaque record images in source order; modeled records in
        // an ignored collection are opaque too and must not affect the layout.
        for record in crate::raw::Records::new(data, limits) {
            let record = record?;
            let (_, header) = crate::raw::Header::parse(&data[record.offset()..], limits)?;
            let end = record.offset() + header + record.len();
            if retained
                .peek()
                .is_some_and(|r| r.bytes() == &data[record.offset()..end])
            {
                retained.next();
                continue;
            }
            if record.kind() == crate::raw::kind::BEGIN_EXT_CONNECTION {
                let id = connection_id(record.payload());
                if scopes.contains(&None) {
                    result
                        .entry(None)
                        .or_default()
                        .push((record.kind().get(), id));
                }
                owner = id;
            } else if is_modeled_record(record.kind()) {
                if scopes.contains(&owner) {
                    result
                        .entry(owner)
                        .or_default()
                        .push((record.kind().get(), None));
                }
                if record.kind() == crate::raw::kind::END_EXT_CONNECTION {
                    owner = None;
                }
            }
        }
        Ok(result)
    };
    let old = layout(before)?;
    let new = layout(after)?;
    for scope in scopes {
        if scope.is_some() && !new.contains_key(&scope) {
            continue;
        }
        if old.get(&scope) != new.get(&scope) {
            return Err(crate::package::error::Error::UnsupportedFeature(
                "connection layout changes around opaque records",
            ));
        }
    }
    Ok(())
}

/// Keep complete unchanged connection collections, and keep a changed
/// collection's original common header when its typed header is unchanged.
pub(crate) fn retain_unchanged_connections(
    source: &[u8],
    old: &super::Connections,
    new: &super::Connections,
    canonical: &[u8],
    limits: crate::raw::Limits,
) -> Result<(Vec<u8>, std::collections::BTreeSet<u32>)> {
    use std::collections::{BTreeMap, BTreeSet};
    fn ranges(
        data: &[u8],
        limits: crate::raw::Limits,
    ) -> Result<BTreeMap<u32, (usize, usize, usize)>> {
        let mut result = BTreeMap::new();
        let mut current = None;
        let mut opaque = Vec::new();
        for r in crate::raw::Records::new(data, limits) {
            let r = r?;
            let kind = r.kind();
            if !opaque.is_empty() || !is_modeled_record(kind) {
                if opaque.last() == Some(&kind) {
                    opaque.pop();
                } else if let Some(end) = crate::package::walker::paired_end(kind) {
                    opaque.push(end);
                }
                continue;
            }
            let (_, header) = crate::raw::Header::parse(&data[r.offset()..], limits)?;
            let end = r.offset() + header + r.len();
            if kind == crate::raw::kind::BEGIN_EXT_CONNECTION {
                current = Some((
                    connection_id(r.payload())
                        .ok_or_else(|| super::invalid("connection ID missing"))?,
                    r.offset(),
                    end,
                ));
            } else if kind == crate::raw::kind::END_EXT_CONNECTION {
                let (id, start, header_end) = current
                    .take()
                    .ok_or_else(|| super::invalid("connection end has no owner"))?;
                result.insert(id, (start, header_end, end));
            }
        }
        Ok(result)
    }
    let original = ranges(source, limits)?;
    let old_bytes = write::write_connections_part(old, limits)?;
    let old_ranges = ranges(&old_bytes, limits)?;
    let new_ranges = ranges(canonical, limits)?;
    let mut ordered = new_ranges.iter().collect::<Vec<_>>();
    ordered.sort_by_key(|(_, r)| r.0);
    let mut output = Vec::new();
    let mut position = 0;
    let mut retained = BTreeSet::new();
    for (&id, &(start, header_end, end)) in ordered {
        output.extend_from_slice(&canonical[position..start]);
        if let (
            Some(before),
            Some(after),
            Some(&(old_start, old_header_end, _)),
            Some(&(src_start, src_header_end, src_end)),
        ) = (
            old.by_id(id),
            new.by_id(id),
            old_ranges.get(&id),
            original.get(&id),
        ) {
            if before == after {
                output.extend_from_slice(&source[src_start..src_end]);
                retained.insert(id);
            } else if old_bytes[old_start..old_header_end] == canonical[start..header_end] {
                output.extend_from_slice(&source[src_start..src_header_end]);
                output.extend_from_slice(&canonical[header_end..end]);
            } else {
                output.extend_from_slice(&canonical[start..end]);
            }
        } else {
            output.extend_from_slice(&canonical[start..end]);
        }
        position = end;
    }
    output.extend_from_slice(&canonical[position..]);
    if output.len() > write::MAX_CONNECTIONS_PART_BYTES {
        return Err(super::invalid(
            "serialized connections part size limit exceeded",
        ));
    }
    Ok((output, retained))
}
