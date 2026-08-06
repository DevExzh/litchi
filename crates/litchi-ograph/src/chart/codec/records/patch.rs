//! In-place cache-value patching for an already parsed chart stream.

use super::cache;
use super::wire::{
    CELL_LABEL, EXCEL_BLANK, EXCEL_BOOL_ERR, EXCEL_NUMBER, GRAPH_BLANK, GRAPH_NUMBER, SI_INDEX,
};
use crate::chart::transaction::{CacheValue, Request};
use crate::chart::{Cache, Chart, Kind, Limits};
use crate::{Error, Result};
use litchi_biff::{Encoder, RecordRef, Records};

struct PhysicalEdit {
    offset: usize,
    payload: Vec<u8>,
}

pub(crate) fn patch(chart: &mut Chart, requests: &[Request]) -> Result<()> {
    if requests.is_empty() {
        return Ok(());
    }
    let kind = chart.context.kind();
    let limits = chart.limits;
    let caches = &chart.caches;
    let stream = match &mut chart.origin {
        crate::chart::model::Origin::Parsed(stream) => stream,
        crate::chart::model::Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "cache-value-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    let source = stream.as_bytes();
    let mut found = Vec::new();
    found
        .try_reserve_exact(requests.len())
        .map_err(|_| Error::Allocation {
            resource: "cache patch markers",
        })?;
    found.resize(requests.len(), false);

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(requests.len())
        .map_err(|_| Error::Allocation {
            resource: "physical cache patches",
        })?;

    let mut section = None;
    let mut cache_index = 0usize;
    for item in Records::new(source) {
        let record = item?;
        if record.kind() == SI_INDEX && kind == Kind::Excel {
            section = Some(read_section(record)?);
            continue;
        }
        if !is_cache_record(record.kind(), kind) {
            continue;
        }

        let index = cache_index;
        cache_index = cache_index.checked_add(1).ok_or(Error::SizeOverflow {
            resource: "cache index",
        })?;
        let Some((request_index, request)) = requests
            .iter()
            .enumerate()
            .find(|(_, request)| request.index == index)
        else {
            continue;
        };
        let current = caches.get(index).ok_or(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "source cache inventory no longer matches the semantic snapshot",
        })?;
        check_identity(current, section, record, kind)?;
        let payload = encode_replacement(current, &request.value, limits, record)?;
        edits.push(PhysicalEdit {
            offset: record.offset(),
            payload,
        });
        found[request_index] = true;
    }

    if cache_index != caches.len() {
        return Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "source cache inventory no longer matches the semantic snapshot",
        });
    }
    if found.iter().any(|found| !found) {
        return Err(Error::InvalidModel {
            field: "cache",
            reason: "cache patch index is outside the parsed chart",
        });
    }

    let bytes = stream.as_bytes_mut();
    for edit in edits {
        let payload_start = edit.offset.checked_add(4).ok_or(Error::SizeOverflow {
            resource: "cache record offset",
        })?;
        let payload_end =
            payload_start
                .checked_add(edit.payload.len())
                .ok_or(Error::SizeOverflow {
                    resource: "cache record payload",
                })?;
        let target =
            bytes
                .get_mut(payload_start..payload_end)
                .ok_or(Error::UnsupportedMutation {
                    operation: "cache-value-patch",
                    reason: "cache record payload falls outside the source stream",
                })?;
        target.copy_from_slice(&edit.payload);
    }
    Ok(())
}

fn is_cache_record(kind: litchi_biff::Kind, chart_kind: Kind) -> bool {
    match chart_kind {
        Kind::Graph => matches!(kind, GRAPH_BLANK | GRAPH_NUMBER | CELL_LABEL),
        Kind::Excel => matches!(
            kind,
            EXCEL_BLANK | EXCEL_BOOL_ERR | EXCEL_NUMBER | CELL_LABEL
        ),
    }
}

fn read_section(record: RecordRef<'_>) -> Result<crate::chart::cache::Index> {
    let value = u16_at(record.payload(), 0).ok_or(Error::InvalidChart {
        offset: record.offset(),
        reason: "SIIndex is truncated",
    })?;
    match value {
        1 => Ok(crate::chart::cache::Index::Values),
        2 => Ok(crate::chart::cache::Index::Categories),
        3 => Ok(crate::chart::cache::Index::Bubbles),
        _ => Err(Error::InvalidChart {
            offset: record.offset(),
            reason: "SIIndex is outside the defined cache sections",
        }),
    }
}

fn check_identity(
    cache: &Cache,
    section: Option<crate::chart::cache::Index>,
    record: RecordRef<'_>,
    kind: Kind,
) -> Result<()> {
    match cache {
        Cache::Graph {
            row,
            col,
            ifmt,
            value,
        } => {
            if kind != Kind::Graph
                || record.kind()
                    != match value {
                        crate::chart::Value::Number(_) => GRAPH_NUMBER,
                        crate::chart::Value::Text(_) => CELL_LABEL,
                        crate::chart::Value::Blank => GRAPH_BLANK,
                    }
                || u16_at(record.payload(), 0) != Some(row.get())
                || u16_at(record.payload(), 2) != Some(col.get())
                || u16_at(record.payload(), 5) != Some(ifmt.get())
            {
                return identity_error();
            }
        },
        Cache::Excel {
            section: expected_section,
            row,
            col,
            xf,
            value,
        } => {
            if kind != Kind::Excel
                || section != Some(*expected_section)
                || record.kind()
                    != match value {
                        crate::chart::XlValue::Number(_) => EXCEL_NUMBER,
                        crate::chart::XlValue::Text(_) => CELL_LABEL,
                        crate::chart::XlValue::Bool(_) | crate::chart::XlValue::Error(_) => {
                            EXCEL_BOOL_ERR
                        },
                        crate::chart::XlValue::Blank => EXCEL_BLANK,
                    }
                || u16_at(record.payload(), 0) != Some(*row)
                || u16_at(record.payload(), 2) != Some(u16::from(*col))
                || u16_at(record.payload(), 4) != Some(xf.get())
            {
                return identity_error();
            }
        },
    }
    Ok(())
}

fn encode_replacement(
    current: &Cache,
    value: &CacheValue,
    limits: Limits,
    source: RecordRef<'_>,
) -> Result<Vec<u8>> {
    let replacement = match (current, value) {
        (Cache::Graph { row, col, ifmt, .. }, CacheValue::Graph(value)) => {
            Cache::graph(*row, *col, *ifmt, value.clone())
        },
        (
            Cache::Excel {
                section,
                row,
                col,
                xf,
                ..
            },
            CacheValue::Excel(value),
        ) => Cache::excel(*section, *row, *col, *xf, value.clone()),
        _ => return identity_error(),
    };

    let mut encoder = Encoder::with_limits(limits.biff)?;
    cache::encode_cache(&mut encoder, &replacement)?;
    let encoded = encoder.finish();
    let mut records = Records::new(&encoded);
    let record = records.next().ok_or(Error::UnsupportedMutation {
        operation: "cache-value-patch",
        reason: "replacement did not produce one cache record",
    })??;
    if records.next().is_some()
        || record.kind() != source.kind()
        || record.payload().len() != source.payload().len()
    {
        return Err(Error::UnsupportedMutation {
            operation: "cache-value-patch",
            reason: "replacement changes the physical cache record length or class",
        });
    }
    let mut payload = Vec::new();
    payload
        .try_reserve_exact(record.payload().len())
        .map_err(|_| Error::Allocation {
            resource: "replacement cache payload",
        })?;
    payload.extend_from_slice(record.payload());
    Ok(payload)
}

fn u16_at(payload: &[u8], offset: usize) -> Option<u16> {
    let bytes = payload.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([*bytes.first()?, *bytes.get(1)?]))
}

fn identity_error<T>() -> Result<T> {
    Err(Error::UnsupportedMutation {
        operation: "cache-value-patch",
        reason: "source cache identity no longer matches the semantic snapshot",
    })
}
