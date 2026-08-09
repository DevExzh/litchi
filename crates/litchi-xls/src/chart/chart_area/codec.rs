//! Fixed-record codec and exact-offset patcher for `Chart`.

use litchi_ograph::chart::Rect;

use super::super::Limits;
use super::super::package::{is_chart_bof, ranges_with};
use super::super::wire::{BOF, CHART, EOF, exact, i32_at};
use super::model::{Change, Snapshot};
use super::validation;
use crate::{Error, Result};

const PAYLOAD_BYTES: usize = 16;

/// Decodes one complete 16-byte `Chart` record payload.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn decode(payload: &[u8]) -> Result<Snapshot> {
    exact(payload, PAYLOAD_BYTES, CHART)?;
    Ok(Snapshot::from_wire(Rect {
        x: i32_at(payload, 0)?,
        y: i32_at(payload, 4)?,
        width: i32_at(payload, 8)?,
        height: i32_at(payload, 12)?,
    }))
}

/// Encodes a semantically valid 16-byte `Chart` record payload.
/// # Errors
///
/// Returns an error if validation, decoding, encoding, or the requested operation fails.
pub fn encode(snapshot: &Snapshot) -> Result<[u8; PAYLOAD_BYTES]> {
    validation::ensure(snapshot.rect())?;
    let rect = snapshot.rect();
    let mut payload = [0; PAYLOAD_BYTES];
    payload[0..4].copy_from_slice(&rect.x.to_le_bytes());
    payload[4..8].copy_from_slice(&rect.y.to_le_bytes());
    payload[8..12].copy_from_slice(&rect.width.to_le_bytes());
    payload[12..16].copy_from_slice(&rect.height.to_le_bytes());
    Ok(payload)
}

/// Patches only the existing `Chart` payload, retaining all record offsets.
pub(crate) fn patch(input: &[u8], change: Change, limits: Limits) -> Result<Vec<u8>> {
    validation::ensure_pair(change.before(), change.after())?;
    let records = ranges_with(input, limits.max_records_per_chart)?;
    if records.first().is_none_or(|value| {
        value.kind != BOF || !is_chart_bof(&input[value.body_start..value.body_end])
    }) || records.last().is_none_or(|value| value.kind != EOF)
    {
        return Err(invalid(
            "chart-area patch requires a complete chart substream",
        ));
    }
    if records.get(1).is_none_or(|value| value.kind != CHART) {
        return Err(unsafe_edit(
            "source Chart record is not at the chart-area owner position",
        ));
    }

    let mut chart_range = None;
    for value in records.iter().filter(|value| value.kind == CHART) {
        if chart_range.is_some() {
            return Err(unsafe_edit("source Chart record is duplicated"));
        }
        let payload = input
            .get(value.body_start..value.body_end)
            .ok_or_else(|| invalid("source Chart payload falls outside the stream"))?;
        let actual = decode(payload)?;
        if actual != Snapshot::from_wire(change.before()) {
            return Err(unsafe_edit(
                "source Chart record does not match the chart-area snapshot",
            ));
        }
        chart_range = Some((value.body_start, value.body_end));
    }
    let (body_start, body_end) =
        chart_range.ok_or_else(|| unsafe_edit("source Chart record is missing"))?;
    let replacement = encode(&Snapshot::from_wire(change.after()))?;
    if body_end.saturating_sub(body_start) != replacement.len() {
        return Err(unsafe_edit("Chart payload length changed during patching"));
    }
    let mut output = input.to_vec();
    output[body_start..body_end].copy_from_slice(&replacement);
    Ok(output)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: CHART,
        message: message.into(),
    }
}

fn unsafe_edit(message: impl Into<String>) -> Error {
    Error::UnsafeEdit(message.into())
}
