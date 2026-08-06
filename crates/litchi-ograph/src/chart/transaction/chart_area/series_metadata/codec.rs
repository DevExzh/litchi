//! Lossless fixed-width `Series` record scanning and patching.

use super::super::super::super::{Chart, Count, DataKind};
use super::model::{Change, Entry, Metadata};
use super::validation;
use crate::chart::model::Origin;
use crate::{Error, Result};
use litchi_biff::{Kind, RecordRef, Records};

const SERIES: Kind = Kind::from_wire(0x1003);
const PAYLOAD_BYTES: usize = 12;
const MAX_COUNT: u16 = 0x0F9F;

pub(super) struct Scan {
    pub(super) source: u64,
    pub(super) entries: Vec<Entry>,
}

pub(super) fn scan(chart: &Chart) -> Result<Scan> {
    let source = match &chart.origin {
        Origin::Parsed(stream) => stream.as_bytes(),
        Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    let mut entries = Vec::new();
    for item in Records::new(source) {
        let record = item?;
        if record.kind() != SERIES {
            continue;
        }
        let metadata = decode(record)?;
        let index = entries.len();
        entries.try_reserve(1).map_err(|_| Error::Allocation {
            resource: "series metadata inventory",
        })?;
        entries.push(Entry {
            index,
            offset: record.offset(),
            metadata,
        });
    }
    if entries.len() != chart.series.len()
        || entries
            .iter()
            .zip(chart.series.iter())
            .any(|(entry, series)| entry.metadata != Metadata::from_series(series))
    {
        return Err(Error::UnsupportedMutation {
            operation: "series-metadata-patch",
            reason: "source Series inventory does not match the semantic snapshot",
        });
    }
    Ok(Scan {
        source: fingerprint(source),
        entries,
    })
}

pub(super) fn patch(chart: &mut Chart, expected_source: u64, changes: &[Change]) -> Result<u64> {
    if changes.is_empty() {
        return Ok(expected_source);
    }
    let current = scan(chart)?;
    if current.source != expected_source {
        return Err(Error::UnsupportedMutation {
            operation: "series-metadata-patch",
            reason: "patch source fingerprint does not match the chart stream",
        });
    }

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(changes.len())
        .map_err(|_| Error::Allocation {
            resource: "physical Series patches",
        })?;
    let mut seen = Vec::new();
    seen.try_reserve_exact(changes.len())
        .map_err(|_| Error::Allocation {
            resource: "Series patch markers",
        })?;
    for change in changes {
        if seen.contains(&change.index()) {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch contains duplicate Series indexes",
            });
        }
        seen.push(change.index());
        let entry = current
            .entries
            .get(change.index())
            .ok_or(Error::InvalidModel {
                field: "series metadata",
                reason: "patch Series index is outside the source inventory",
            })?;
        if entry.index != change.index() || entry.offset != change.offset() {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch Series record identity does not match the source",
            });
        }
        if entry.metadata != change.before() {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "patch Series metadata does not match the source",
            });
        }
        validation::ensure(change.after())?;
        let payload_start = change.offset().checked_add(4).ok_or(Error::SizeOverflow {
            resource: "Series record offset",
        })?;
        let payload_end = payload_start
            .checked_add(PAYLOAD_BYTES)
            .ok_or(Error::SizeOverflow {
                resource: "Series record payload",
            })?;
        edits.push(Physical {
            payload_start,
            payload_end,
            payload: encode(change.after()),
        });
    }

    // Complete all bounds checks before the first byte or semantic field is
    // changed.  The mutation phase below contains only infallible fixed-slice
    // copies and direct field assignments.
    let source_len = match &chart.origin {
        Origin::Parsed(stream) => stream.as_bytes().len(),
        Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    if edits
        .iter()
        .any(|edit| edit.payload_end > source_len || edit.payload_start > edit.payload_end)
    {
        return Err(Error::UnsupportedMutation {
            operation: "series-metadata-patch",
            reason: "Series payload falls outside the source stream",
        });
    }
    if changes
        .iter()
        .any(|change| change.index() >= chart.series.len())
    {
        return Err(Error::InvalidModel {
            field: "series metadata",
            reason: "Series index disappeared during publication",
        });
    }

    let bytes = match &mut chart.origin {
        Origin::Parsed(stream) => stream.as_bytes_mut(),
        Origin::Fresh => {
            return Err(Error::UnsupportedMutation {
                operation: "series-metadata-patch",
                reason: "the source chart is not a parsed replayable stream",
            });
        },
    };
    for edit in &edits {
        let target = &mut bytes[edit.payload_start..edit.payload_end];
        target.copy_from_slice(&edit.payload);
    }
    for change in changes {
        // The complete preflight above proves every index exists.  This direct
        // semantic update cannot fail and keeps Chart::encode pristine.
        change.after().apply(&mut chart.series[change.index()]);
    }
    Ok(fingerprint(bytes))
}

struct Physical {
    payload_start: usize,
    payload_end: usize,
    payload: [u8; PAYLOAD_BYTES],
}

fn decode(record: RecordRef<'_>) -> Result<Metadata> {
    let data = record.payload();
    if data.len() != PAYLOAD_BYTES {
        return Err(Error::InvalidRecordLength {
            kind: u16::from(SERIES),
            expected: PAYLOAD_BYTES,
            actual: data.len(),
        });
    }
    let category_kind = match u16_at(data, 0) {
        1 => DataKind::Numeric,
        3 => DataKind::Text,
        value => {
            return Err(Error::InvalidRecordValue {
                kind: u16::from(SERIES),
                field: "sdtX",
                value: u64::from(value),
            });
        },
    };
    for (field, offset) in [("sdtY", 2usize), ("sdtBSize", 8usize)] {
        let value = u16_at(data, offset);
        if value != 1 {
            return Err(Error::InvalidRecordValue {
                kind: u16::from(SERIES),
                field,
                value: u64::from(value),
            });
        }
    }
    let category_count = count(u16_at(data, 4), "cValx")?;
    let value_count = count(u16_at(data, 6), "cValy")?;
    let bubble_count = count(u16_at(data, 10), "cValBSize")?;
    let metadata = Metadata::new(category_kind, category_count, value_count, bubble_count);
    validation::ensure(metadata)?;
    Ok(metadata)
}

fn count(value: u16, field: &'static str) -> Result<Count> {
    if value > MAX_COUNT {
        return Err(Error::InvalidRecordValue {
            kind: u16::from(SERIES),
            field,
            value: u64::from(value),
        });
    }
    Count::new(value).ok_or(Error::InvalidRecordValue {
        kind: u16::from(SERIES),
        field,
        value: u64::from(value),
    })
}

fn u16_at(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn encode(metadata: Metadata) -> [u8; PAYLOAD_BYTES] {
    let mut data = [0u8; PAYLOAD_BYTES];
    data[0..2].copy_from_slice(
        &match metadata.category_kind() {
            DataKind::Numeric => 1u16,
            DataKind::Text => 3u16,
        }
        .to_le_bytes(),
    );
    data[2..4].copy_from_slice(&1u16.to_le_bytes());
    data[4..6].copy_from_slice(&metadata.category_count().get().to_le_bytes());
    data[6..8].copy_from_slice(&metadata.value_count().get().to_le_bytes());
    data[8..10].copy_from_slice(&1u16.to_le_bytes());
    data[10..12].copy_from_slice(&metadata.bubble_count().get().to_le_bytes());
    data
}

pub(super) fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01B3);
    }
    hash
}
