//! Strict per-series connection-line geometry CRUD for native 2D charts.
//!
//! Connection geometry is behavioral rather than visual, so iWork stores it
//! in sparse series non-style objects. Line and radar charts default to
//! straight segments; scatter charts default to no connecting lines.

use crate::charts::ChartKind;
use crate::charts::series_non_style::{
    GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD, NewChartSeriesNonStyleBase,
    chart_series_non_style_values, generated_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const LINE_LINE_TYPE_FIELD: u32 = 18;
const SCATTER_SHOW_LINE_FIELD: u32 = 29;
const SCATTER_LINE_TYPE_FIELD: u32 = 20;
const RADAR_LINE_TYPE_FIELD: u32 = 189;

const HIDDEN_NATIVE_VALUE: i32 = 0;
const STRAIGHT_NATIVE_VALUE: i32 = 1;
const CURVED_NATIVE_VALUE: i32 = 2;

/// Geometry used to connect adjacent data points in one chart series.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartSeriesConnectionLine {
    /// Do not connect data points. Native scatter charts support this mode.
    Hidden,
    /// Connect data points with straight segments.
    Straight,
    /// Interpolate a smooth curve through the data points.
    Curved,
}

impl ChartSeriesConnectionLine {
    const fn native(self) -> i32 {
        match self {
            Self::Hidden => HIDDEN_NATIVE_VALUE,
            Self::Straight => STRAIGHT_NATIVE_VALUE,
            Self::Curved => CURVED_NATIVE_VALUE,
        }
    }

    fn from_native(value: i32) -> Result<Self> {
        match value {
            HIDDEN_NATIVE_VALUE => Ok(Self::Hidden),
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
    Line2d,
    Radar2d,
    Scatter2d,
}

impl ChartSeriesConnectionLineKind {
    /// Resolve the native field family used by a chart kind.
    pub fn for_chart_kind(kind: ChartKind) -> Result<Self> {
        match kind {
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
            Self::Line2d => LINE_LINE_TYPE_FIELD,
            Self::Radar2d => RADAR_LINE_TYPE_FIELD,
            Self::Scatter2d => SCATTER_LINE_TYPE_FIELD,
        }
    }

    const fn default_value(self) -> ChartSeriesConnectionLine {
        match self {
            Self::Line2d | Self::Radar2d => ChartSeriesConnectionLine::Straight,
            Self::Scatter2d => ChartSeriesConnectionLine::Hidden,
        }
    }

    fn validate(self, value: ChartSeriesConnectionLine) -> Result<()> {
        if value == ChartSeriesConnectionLine::Hidden && self != Self::Scatter2d {
            return Err(Error::InvalidFormat(format!(
                "{self:?} chart series cannot hide connection lines"
            )));
        }
        Ok(())
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
        storage.default_value(),
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
    for value in expected {
        storage.validate(*value)?;
    }
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series connection-line",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        storage.default_value(),
        |data| read_connection_line(data, storage),
        |data, replacement| patch_connection_line(data, storage, *replacement),
    )
}

fn read_connection_line(
    data: &[u8],
    storage: ChartSeriesConnectionLineKind,
) -> Result<ChartSeriesConnectionLine> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(storage.default_value());
    };
    let line_type = strict_optional_enum(extension, storage.field_number())?;
    if storage == ChartSeriesConnectionLineKind::Scatter2d {
        let visible = strict_optional_bool(extension, SCATTER_SHOW_LINE_FIELD)?.unwrap_or(false);
        if !visible {
            if let Some(value) = line_type {
                ChartSeriesConnectionLine::from_native(value)?;
            }
            return Ok(ChartSeriesConnectionLine::Hidden);
        }
        let line_type = line_type.unwrap_or(STRAIGHT_NATIVE_VALUE);
        if line_type == HIDDEN_NATIVE_VALUE {
            return Err(Error::InvalidFormat(
                "visible scatter connection lines use the hidden line type".to_owned(),
            ));
        }
        return ChartSeriesConnectionLine::from_native(line_type);
    }
    line_type.map_or(Ok(storage.default_value()), |value| {
        let value = ChartSeriesConnectionLine::from_native(value)?;
        storage.validate(value)?;
        Ok(value)
    })
}

fn patch_connection_line(
    data: &[u8],
    storage: ChartSeriesConnectionLineKind,
    replacement: ChartSeriesConnectionLine,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    storage.validate(replacement)?;
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    let line_type_present = strict_optional_enum(extension, field_number)?.is_some();
    let (show_line, native) = match (storage, replacement) {
        (ChartSeriesConnectionLineKind::Scatter2d, ChartSeriesConnectionLine::Hidden) => {
            (None, None)
        },
        (ChartSeriesConnectionLineKind::Scatter2d, replacement) => (
            Some(1),
            Some(u64::try_from(replacement.native()).map_err(|_| {
                Error::InvalidFormat("negative chart series connection-line value".to_owned())
            })?),
        ),
        (_, ChartSeriesConnectionLine::Straight) => (None, None),
        (_, ChartSeriesConnectionLine::Curved) => (None, Some(CURVED_NATIVE_VALUE as u64)),
        (_, ChartSeriesConnectionLine::Hidden) => unreachable!("validated above"),
    };
    let show_line_present = strict_optional_bool(extension, SCATTER_SHOW_LINE_FIELD)?.is_some();
    let extension = if storage == ChartSeriesConnectionLineKind::Scatter2d {
        patch_varint_field(
            extension,
            SCATTER_SHOW_LINE_FIELD,
            show_line_present,
            show_line,
        )?
    } else {
        extension.to_vec()
    };
    let extension = patch_varint_field(&extension, field_number, line_type_present, native)?;
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

fn strict_optional_bool(data: &[u8], field_number: u32) -> Result<Option<bool>> {
    strict_optional_varint(data, field_number, "boolean")?.map_or(Ok(None), |value| {
        if value > 1 {
            return Err(Error::InvalidFormat(format!(
                "chart series connection-line field {field_number} is not boolean"
            )));
        }
        Ok(Some(value == 1))
    })
}

fn strict_optional_enum(data: &[u8], field_number: u32) -> Result<Option<i32>> {
    strict_optional_varint(data, field_number, "enum")?
        .map(|value| {
            i32::try_from(value).map_err(|_| {
                Error::InvalidFormat(format!(
                    "chart series connection-line field {field_number} exceeds i32"
                ))
            })
        })
        .transpose()
}

fn strict_optional_varint(data: &[u8], field_number: u32, label: &str) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series connection-line {label} field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series connection-line {label} field {field_number} is not a varint"
        )));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart series connection-line {label} field {field_number} is invalid: {error}"
        ))
    })?;
    if consumed != field.end - field.key_end {
        return Err(Error::InvalidFormat(format!(
            "chart series connection-line {label} field {field_number} has trailing bytes"
        )));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::{tsch, tss};
    use crate::wire::{append_varint_field, parse_wire_fields};
    use prost::Message;

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn native_connection_line_values_are_strict_and_reversible() {
        assert_eq!(
            ChartSeriesConnectionLine::from_native(HIDDEN_NATIVE_VALUE).unwrap(),
            ChartSeriesConnectionLine::Hidden
        );
        assert_eq!(
            ChartSeriesConnectionLine::from_native(STRAIGHT_NATIVE_VALUE).unwrap(),
            ChartSeriesConnectionLine::Straight
        );
        assert_eq!(
            ChartSeriesConnectionLine::from_native(CURVED_NATIVE_VALUE).unwrap(),
            ChartSeriesConnectionLine::Curved
        );
        for invalid in [i32::MIN, 3, i32::MAX] {
            assert!(ChartSeriesConnectionLine::from_native(invalid).is_err());
        }
    }

    #[test]
    fn every_unambiguous_connection_line_kind_has_a_typed_storage_family() {
        for kind in [ChartKind::Line2d, ChartKind::Radar2d, ChartKind::Scatter2d] {
            assert!(ChartSeriesConnectionLineKind::for_chart_kind(kind).is_ok());
        }
        for kind in [
            ChartKind::Area2d,
            ChartKind::Bubble2d,
            ChartKind::MultiDataScatter2d,
            ChartKind::Mixed2d,
            ChartKind::TwoAxis2d,
            ChartKind::Line3d,
        ] {
            assert!(ChartSeriesConnectionLineKind::for_chart_kind(kind).is_err());
        }
    }

    #[test]
    fn scatter_hidden_straight_and_curved_states_round_trip_canonically() {
        let original = style_with_unknown_fields();
        assert_eq!(
            read_connection_line(&original, ChartSeriesConnectionLineKind::Scatter2d).unwrap(),
            ChartSeriesConnectionLine::Hidden
        );
        let existing_extension = generated_chart_series_non_style_extension(&original)
            .unwrap()
            .unwrap();
        let app_hidden_extension =
            patch_varint_field(existing_extension, SCATTER_SHOW_LINE_FIELD, false, Some(0))
                .and_then(|extension| {
                    patch_varint_field(
                        &extension,
                        SCATTER_LINE_TYPE_FIELD,
                        false,
                        Some(HIDDEN_NATIVE_VALUE as u64),
                    )
                })
                .unwrap();
        let app_hidden = patch_length_delimited_field(
            &original,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(&app_hidden_extension),
        )
        .unwrap();
        assert_eq!(
            read_connection_line(&app_hidden, ChartSeriesConnectionLineKind::Scatter2d).unwrap(),
            ChartSeriesConnectionLine::Hidden
        );
        let straight = patch_connection_line(
            &original,
            ChartSeriesConnectionLineKind::Scatter2d,
            ChartSeriesConnectionLine::Straight,
        )
        .unwrap();
        assert_eq!(
            read_connection_line(&straight, ChartSeriesConnectionLineKind::Scatter2d).unwrap(),
            ChartSeriesConnectionLine::Straight
        );
        let curved = patch_connection_line(
            &straight,
            ChartSeriesConnectionLineKind::Scatter2d,
            ChartSeriesConnectionLine::Curved,
        )
        .unwrap();
        assert_eq!(
            read_connection_line(&curved, ChartSeriesConnectionLineKind::Scatter2d).unwrap(),
            ChartSeriesConnectionLine::Curved
        );
        assert_unknown_fields_retained(&original, &curved);
        let hidden = patch_connection_line(
            &curved,
            ChartSeriesConnectionLineKind::Scatter2d,
            ChartSeriesConnectionLine::Hidden,
        )
        .unwrap();
        assert_eq!(hidden, original);
    }

    #[test]
    fn hidden_connections_are_rejected_for_line_and_radar_charts() {
        for storage in [
            ChartSeriesConnectionLineKind::Line2d,
            ChartSeriesConnectionLineKind::Radar2d,
        ] {
            assert!(
                patch_connection_line(
                    &style_with_unknown_fields(),
                    storage,
                    ChartSeriesConnectionLine::Hidden,
                )
                .is_err()
            );
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
