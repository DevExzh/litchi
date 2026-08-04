//! Wire-preserving Arrange-panel state for native chart drawables.

use crate::archive::RawMessage;
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::wire::{parse_wire_fields, patch_varint_field, transform_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;

/// Editable state exposed by the chart Arrange panel.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct ChartArrangement {
    locked: bool,
    constrain_proportions: bool,
}

impl ChartArrangement {
    /// Construct chart Arrange-panel state.
    pub const fn new(locked: bool, constrain_proportions: bool) -> Self {
        Self {
            locked,
            constrain_proportions,
        }
    }

    /// Return whether the chart is locked against interactive editing.
    pub const fn locked(self) -> bool {
        self.locked
    }

    /// Return whether interactive resizing preserves the chart's aspect ratio.
    pub const fn constrain_proportions(self) -> bool {
        self.constrain_proportions
    }

    /// Return this state with the requested interactive lock.
    pub const fn with_locked(mut self, locked: bool) -> Self {
        self.locked = locked;
        self
    }

    /// Return this state with the requested aspect-ratio constraint.
    pub const fn with_constrain_proportions(mut self, constrain_proportions: bool) -> Self {
        self.constrain_proportions = constrain_proportions;
        self
    }
}

/// Read one chart's effective Arrange-panel state.
pub(crate) fn chart_arrangement(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartArrangement> {
    let (_, message) = chart_message(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    read_arrangement(&message)
}

/// Set one chart's Arrange-panel state without normalizing unrelated bytes.
pub(crate) fn set_chart_arrangement(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    arrangement: ChartArrangement,
) -> Result<()> {
    let (message_index, message) = chart_message(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = raw_arrangement(&message)?;
    if current.effective() == arrangement {
        return Ok(());
    }
    let data =
        transform_length_delimited_field(&message, CHART_DRAWABLE_SUPER_FIELD, |drawable| {
            patch_arrangement(drawable, current, arrangement)
        })?;
    if read_arrangement(&data)? != arrangement {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} arrangement update failed validation"
        )));
    }
    package.update_archive(chart_archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        Ok(object
            .replace_message(
                message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
            )
            .map(|_| ())?)
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RawChartArrangement {
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
}

impl RawChartArrangement {
    fn effective(self) -> ChartArrangement {
        ChartArrangement::new(
            self.locked.unwrap_or(false),
            self.constrain_proportions.unwrap_or(false),
        )
    }
}

fn read_arrangement(data: &[u8]) -> Result<ChartArrangement> {
    Ok(raw_arrangement(data)?.effective())
}

fn raw_arrangement(data: &[u8]) -> Result<RawChartArrangement> {
    let fields = parse_wire_fields(data)?;
    let drawable = singular_field(&fields, CHART_DRAWABLE_SUPER_FIELD, "chart drawable super")?;
    require_wire_type(drawable, 2, "chart drawable super")?;
    let drawable = &data[drawable.payload_start..drawable.end];
    Ok(RawChartArrangement {
        locked: strict_optional_bool(drawable, DRAWABLE_LOCKED_FIELD, "chart lock")?,
        constrain_proportions: strict_optional_bool(
            drawable,
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
            "chart aspect-ratio lock",
        )?,
    })
}

fn patch_arrangement(
    drawable: &[u8],
    current: RawChartArrangement,
    replacement: ChartArrangement,
) -> Result<Vec<u8>> {
    let mut patched = patch_varint_field(
        drawable,
        DRAWABLE_LOCKED_FIELD,
        current.locked.is_some(),
        replacement_presence(current.locked, replacement.locked()).map(u64::from),
    )?;
    patched = patch_varint_field(
        &patched,
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
        current.constrain_proportions.is_some(),
        replacement_presence(
            current.constrain_proportions,
            replacement.constrain_proportions(),
        )
        .map(u64::from),
    )?;
    Ok(patched)
}

const fn replacement_presence(current: Option<bool>, replacement: bool) -> Option<bool> {
    if current.is_some() || replacement {
        Some(replacement)
    } else {
        None
    }
}

fn chart_message(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<(usize, Vec<u8>)> {
    let archive = package.archive(chart_archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE);
    let Some((message_index, message)) = messages.next() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} drawable {drawable_object_id} has no chart payload"
        )));
    };
    if messages.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} drawable {drawable_object_id} has multiple chart payloads"
        )));
    }
    Ok((message_index, message.data.clone()))
}

fn strict_optional_bool(data: &[u8], field_number: u32, label: &str) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "{label} field occurs {} times",
            matches.len()
        )));
    }
    let Some(field) = matches.first().copied() else {
        return Ok(None);
    };
    require_wire_type(field, 0, label)?;
    let (value, length) =
        crate::varint::decode_varint_from_bytes(&data[field.payload_start..field.end])
            .map_err(|error| Error::InvalidFormat(format!("invalid {label}: {error}")))?;
    if field.payload_start + length != field.end {
        return Err(Error::InvalidFormat(format!(
            "{label} contains trailing bytes"
        )));
    }
    match value {
        0 => Ok(Some(false)),
        1 => Ok(Some(true)),
        _ => Err(Error::InvalidFormat(format!(
            "{label} must be encoded as zero or one, found {value}"
        ))),
    }
}

fn singular_field<'a>(
    fields: &'a [crate::wire::WireField],
    field_number: u32,
    label: &str,
) -> Result<&'a crate::wire::WireField> {
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "{label} must occur exactly once, found {}",
            matches.len()
        )));
    }
    Ok(matches[0])
}

fn require_wire_type(field: &crate::wire::WireField, expected: u8, label: &str) -> Result<()> {
    if field.wire_type != expected {
        return Err(Error::InvalidFormat(format!(
            "{label} has wire type {}, expected {expected}",
            field.wire_type
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use prost::Message;

    use super::*;
    use crate::protobuf::{tsch, tsd};
    use crate::wire::append_varint_field;

    #[test]
    fn arrangement_patch_preserves_unknowns_and_restores_explicit_defaults() {
        let drawable = tsd::DrawableArchive {
            locked: Some(false),
            aspect_ratio_locked: Some(false),
            ..Default::default()
        };
        let chart = tsch::ChartDrawableArchive {
            super_: Some(drawable),
        };
        let mut original = chart.encode_to_vec();
        append_varint_field(&mut original, 99, 990).unwrap();
        let changed =
            transform_length_delimited_field(&original, CHART_DRAWABLE_SUPER_FIELD, |drawable| {
                patch_arrangement(
                    drawable,
                    RawChartArrangement {
                        locked: Some(false),
                        constrain_proportions: Some(false),
                    },
                    ChartArrangement::new(true, true),
                )
            })
            .unwrap();
        assert_eq!(
            read_arrangement(&changed).unwrap(),
            ChartArrangement::new(true, true)
        );
        assert!(
            changed
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0xde])
        );

        let restored =
            transform_length_delimited_field(&changed, CHART_DRAWABLE_SUPER_FIELD, |drawable| {
                patch_arrangement(
                    drawable,
                    RawChartArrangement {
                        locked: Some(true),
                        constrain_proportions: Some(true),
                    },
                    ChartArrangement::default(),
                )
            })
            .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn arrangement_reader_rejects_duplicate_and_non_boolean_fields() {
        let drawable = tsd::DrawableArchive {
            locked: Some(false),
            ..Default::default()
        };
        let chart = tsch::ChartDrawableArchive {
            super_: Some(drawable.clone()),
        };
        let duplicate = transform_length_delimited_field(
            &chart.encode_to_vec(),
            CHART_DRAWABLE_SUPER_FIELD,
            |drawable| {
                let mut duplicate = drawable.to_vec();
                append_varint_field(&mut duplicate, DRAWABLE_LOCKED_FIELD, 1)?;
                Ok(duplicate)
            },
        )
        .unwrap();
        assert!(read_arrangement(&duplicate).is_err());

        let non_boolean = transform_length_delimited_field(
            &tsch::ChartDrawableArchive {
                super_: Some(drawable),
            }
            .encode_to_vec(),
            CHART_DRAWABLE_SUPER_FIELD,
            |drawable| patch_varint_field(drawable, DRAWABLE_LOCKED_FIELD, true, Some(2)),
        )
        .unwrap();
        assert!(read_arrangement(&non_boolean).is_err());
    }
}
