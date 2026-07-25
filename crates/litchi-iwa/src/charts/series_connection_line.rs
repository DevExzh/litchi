//! Strict per-series connection-line geometry CRUD for native 2D charts.
//!
//! Connection geometry is behavioral rather than visual, so iWork stores it
//! in sparse series non-style objects. Straight is the native default; curved
//! is an explicit override.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_non_style::{
    GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD, chart_series_non_style_values,
    generated_chart_series_non_style_extension, set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const BUBBLE_LINE_TYPE_FIELD: u32 = 17;
const LINE_LINE_TYPE_FIELD: u32 = 18;
const SCATTER_LINE_TYPE_FIELD: u32 = 20;
const RADAR_LINE_TYPE_FIELD: u32 = 189;

const STRAIGHT_NATIVE_VALUE: i32 = 1;
const CURVED_NATIVE_VALUE: i32 = 2;

/// Geometry used to connect adjacent data points in one chart series.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesConnectionLine {
    /// Connect data points with straight segments.
    #[default]
    Straight,
    /// Interpolate a smooth curve through the data points.
    Curved,
}

impl ChartSeriesConnectionLine {
    const fn native(self) -> i32 {
        match self {
            Self::Straight => STRAIGHT_NATIVE_VALUE,
            Self::Curved => CURVED_NATIVE_VALUE,
        }
    }

    fn from_native(value: i32) -> Result<Self> {
        match value {
            STRAIGHT_NATIVE_VALUE => Ok(Self::Straight),
            CURVED_NATIVE_VALUE => Ok(Self::Curved),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native chart series connection-line type {value}"
            ))),
        }
    }
}

/// Native connection-line field family for an unambiguous 2D chart kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesConnectionLineKind {
    Bubble2d,
    Line2d,
    Radar2d,
    Scatter2d,
}

impl ChartSeriesConnectionLineKind {
    /// Resolve the native field family used by a chart kind.
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Bubble2d => Ok(Self::Bubble2d),
            ChartKind::Line2d => Ok(Self::Line2d),
            ChartKind::Radar2d => Ok(Self::Radar2d),
            ChartKind::Scatter2d => Ok(Self::Scatter2d),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no unambiguous series connection-line family"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Bubble2d => BUBBLE_LINE_TYPE_FIELD,
            Self::Line2d => LINE_LINE_TYPE_FIELD,
            Self::Radar2d => RADAR_LINE_TYPE_FIELD,
            Self::Scatter2d => SCATTER_LINE_TYPE_FIELD,
        }
    }
}

/// Read connection geometry in native series order.
pub(crate) fn chart_series_connection_lines(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesConnectionLine>> {
    let storage = ChartSeriesConnectionLineKind::for_chart_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesConnectionLine::Straight,
        |data| read_connection_line(data, storage),
    )
}

/// Replace every series' connection geometry in native series order.
#[allow(clippy::too_many_arguments)]
pub(crate) fn set_chart_series_connection_lines(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[ChartSeriesConnectionLine],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} connection-line settings",
            expected.len()
        )));
    }
    let storage = ChartSeriesConnectionLineKind::for_chart_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series connection-line",
        expected,
        ChartSeriesConnectionLine::Straight,
        |data| read_connection_line(data, storage),
        |data, replacement| patch_connection_line(data, storage, *replacement),
    )
}

fn read_connection_line(
    data: &[u8],
    storage: ChartSeriesConnectionLineKind,
) -> Result<ChartSeriesConnectionLine> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartSeriesConnectionLine::Straight);
    };
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    line_type_field(&generated, storage.field_number())
        .map_or(Ok(ChartSeriesConnectionLine::Straight), |value| {
            ChartSeriesConnectionLine::from_native(value)
        })
}

fn line_type_field(
    generated: &tsch::generated::ChartSeriesNonStyleArchive,
    field_number: u32,
) -> Option<i32> {
    match field_number {
        BUBBLE_LINE_TYPE_FIELD => generated.tschchartseriesbubblelinetype,
        LINE_LINE_TYPE_FIELD => generated.tschchartserieslinelinetype,
        RADAR_LINE_TYPE_FIELD => generated.tschchartseriesradararealinetype,
        SCATTER_LINE_TYPE_FIELD => generated.tschchartseriesscatterlinetype,
        _ => None,
    }
}

fn patch_connection_line(
    data: &[u8],
    storage: ChartSeriesConnectionLineKind,
    replacement: ChartSeriesConnectionLine,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let present = line_type_field(&generated, field_number).is_some();
    let native = (replacement != ChartSeriesConnectionLine::Straight).then_some(
        u64::try_from(replacement.native()).map_err(|_| {
            Error::InvalidFormat("negative chart series connection-line value".to_owned())
        })?,
    );
    let extension = patch_varint_field(extension, field_number, present, native)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_connection_line(&patched, storage)? != replacement {
        return Err(Error::InvalidFormat(
            "chart series connection-line wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn native_connection_line_values_are_strict_and_reversible() {
        assert_eq!(
            ChartSeriesConnectionLine::from_native(STRAIGHT_NATIVE_VALUE).unwrap(),
            ChartSeriesConnectionLine::Straight
        );
        assert_eq!(
            ChartSeriesConnectionLine::from_native(CURVED_NATIVE_VALUE).unwrap(),
            ChartSeriesConnectionLine::Curved
        );
        for invalid in [i32::MIN, 0, 3, i32::MAX] {
            assert!(ChartSeriesConnectionLine::from_native(invalid).is_err());
        }
    }

    #[test]
    fn every_unambiguous_connection_line_kind_has_a_typed_storage_family() {
        for kind in [
            ChartKind::Bubble2d,
            ChartKind::Line2d,
            ChartKind::Radar2d,
            ChartKind::Scatter2d,
        ] {
            assert!(ChartSeriesConnectionLineKind::for_chart_kind(kind).is_ok());
        }
        for kind in [
            ChartKind::Area2d,
            ChartKind::Mixed2d,
            ChartKind::TwoAxis2d,
            ChartKind::Line3d,
        ] {
            assert!(ChartSeriesConnectionLineKind::for_chart_kind(kind).is_err());
        }
    }

    #[test]
    fn curved_override_round_trips_and_straight_restores_exact_bytes() {
        let original = style_with_unknown_fields();
        let curved = patch_connection_line(
            &original,
            ChartSeriesConnectionLineKind::Line2d,
            ChartSeriesConnectionLine::Curved,
        )
        .unwrap();
        assert_eq!(
            read_connection_line(&curved, ChartSeriesConnectionLineKind::Line2d).unwrap(),
            ChartSeriesConnectionLine::Curved
        );
        assert_unknown_fields_retained(&original, &curved);
        let restored = patch_connection_line(
            &curved,
            ChartSeriesConnectionLineKind::Line2d,
            ChartSeriesConnectionLine::Straight,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    fn style_with_unknown_fields() -> Vec<u8> {
        let mut generated = tsch::generated::ChartSeriesNonStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 77).unwrap();
        let mut outer = tsch::ChartSeriesNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        outer = patch_length_delimited_field(
            &outer,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.as_slice()),
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 91).unwrap();
        outer
    }

    fn assert_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        let field = |data: &[u8], number| {
            parse_wire_fields(data)
                .unwrap()
                .into_iter()
                .find(|field| field.number == number)
                .map(|field| data[field.start..field.end].to_vec())
        };
        assert_eq!(
            field(patched, UNKNOWN_OUTER_FIELD),
            field(original, UNKNOWN_OUTER_FIELD)
        );
        let original_generated = generated_chart_series_non_style_extension(original)
            .unwrap()
            .unwrap();
        let patched_generated = generated_chart_series_non_style_extension(patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            field(patched_generated, UNKNOWN_GENERATED_FIELD),
            field(original_generated, UNKNOWN_GENERATED_FIELD)
        );
    }
}
