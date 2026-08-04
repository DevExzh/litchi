//! Lossless per-series value-label location CRUD for native bar and column charts.
//!
//! iWork stores the inspector's “Location” choice in sparse series-style
//! objects. Stacked and unstacked charts use distinct generated fields.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::series_style::{
    GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD, chart_series_style_slots,
    generated_chart_series_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const BAR_VALUE_LABEL_LOCATION_FIELD: u32 = 88;
const STACKED_BAR_VALUE_LABEL_LOCATION_FIELD: u32 = 97;

const NATIVE_MIDDLE: u64 = 0;
const NATIVE_OUTSIDE: u64 = 6;
const NATIVE_TOP: u64 = 4;
const NATIVE_BOTTOM: u64 = 8;

/// Placement of data-value labels for one bar or column chart series.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesValueLabelLocation {
    Bottom,
    Middle,
    #[default]
    Top,
    Outside,
}

impl ChartSeriesValueLabelLocation {
    const fn native_value(self) -> u64 {
        match self {
            Self::Bottom => NATIVE_BOTTOM,
            Self::Middle => NATIVE_MIDDLE,
            Self::Top => NATIVE_TOP,
            Self::Outside => NATIVE_OUTSIDE,
        }
    }

    fn from_native(value: u64) -> Result<Self> {
        match value {
            NATIVE_BOTTOM => Ok(Self::Bottom),
            NATIVE_MIDDLE => Ok(Self::Middle),
            NATIVE_TOP => Ok(Self::Top),
            NATIVE_OUTSIDE => Ok(Self::Outside),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported chart series value-label location {value}"
            ))),
        }
    }
}

/// Read value-label locations in native series order.
pub(crate) fn chart_series_value_label_locations(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesValueLabelLocation>> {
    let storage = LocationStorage::for_kind(kind)?;
    let slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slots.len() < series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {} series styles for {series_count} series",
            slots.len()
        )));
    }
    slots
        .iter()
        .take(series_count)
        .map(|slot| slot.read(package, |data| read_location(data, storage)))
        .collect()
}

/// Set value-label locations in native series order.
pub(crate) fn set_chart_series_value_label_locations(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
    expected: &[ChartSeriesValueLabelLocation],
) -> Result<()> {
    if expected.len() != series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {series_count} series, not {} value-label locations",
            expected.len()
        )));
    }
    let storage = LocationStorage::for_kind(kind)?;
    let slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slots.len() < series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {} series styles for {series_count} series",
            slots.len()
        )));
    }
    for slot in slots.iter().take(series_count) {
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    }
    for (slot, &location) in slots.iter().zip(expected) {
        slot.update(package, |data| patch_location(data, storage, location))?;
    }
    if chart_series_value_label_locations(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        kind,
        series_count,
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-label location update failed validation"
        )));
    }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum LocationStorage {
    Bar,
    StackedBar,
}

impl LocationStorage {
    fn for_kind(kind: ChartKind) -> Result<Self> {
        match kind {
            ChartKind::Column2d
            | ChartKind::Column3d
            | ChartKind::Bar2d
            | ChartKind::Bar3d
            | ChartKind::MultiDataColumn2d
            | ChartKind::MultiDataBar2d => Ok(Self::Bar),
            ChartKind::StackedColumn2d
            | ChartKind::StackedColumn3d
            | ChartKind::StackedBar2d
            | ChartKind::StackedBar3d => Ok(Self::StackedBar),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no supported series value-label location"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Bar => BAR_VALUE_LABEL_LOCATION_FIELD,
            Self::StackedBar => STACKED_BAR_VALUE_LABEL_LOCATION_FIELD,
        }
    }
}

fn read_location(data: &[u8], storage: LocationStorage) -> Result<ChartSeriesValueLabelLocation> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(ChartSeriesValueLabelLocation::Top);
    };
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    strict_optional_location(extension, storage.field_number())?
        .map(ChartSeriesValueLabelLocation::from_native)
        .transpose()
        .map(|location| location.unwrap_or_default())
}

fn patch_location(
    data: &[u8],
    storage: LocationStorage,
    location: ChartSeriesValueLabelLocation,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        if location == ChartSeriesValueLabelLocation::Top {
            return Ok(data.to_vec());
        }
        let mut extension = Vec::new();
        crate::wire::append_varint_field(&mut extension, field_number, location.native_value())?;
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patch(&patched, storage, location)?;
        return Ok(patched);
    };

    let present = strict_optional_location(extension, field_number)?.is_some();
    let replacement =
        (location != ChartSeriesValueLabelLocation::Top).then_some(location.native_value());
    let extension = patch_varint_field(extension, field_number, present, replacement)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patch(&patched, storage, location)?;
    Ok(patched)
}

fn strict_optional_location(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series value-label location field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label location field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "chart series value-label location field {field_number} is invalid: {error}"
                ))
            })?;
    if consumed != field.end - field.key_end
        || litchi_iwa_common::varint::encoded_len(value) != consumed
    {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label location field {field_number} is not canonical"
        )));
    }
    ChartSeriesValueLabelLocation::from_native(value)?;
    Ok(Some(value))
}

fn validate_patch(
    data: &[u8],
    storage: LocationStorage,
    expected: ChartSeriesValueLabelLocation,
) -> Result<()> {
    if read_location(data, storage)? != expected {
        return Err(Error::InvalidFormat(
            "chart series value-label location wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    #[test]
    fn bar_and_stacked_bar_use_distinct_native_fields() {
        assert_eq!(LocationStorage::Bar.field_number(), 88);
        assert_eq!(LocationStorage::StackedBar.field_number(), 97);
        assert!(LocationStorage::for_kind(ChartKind::Column2d).is_ok());
        assert!(LocationStorage::for_kind(ChartKind::StackedColumn2d).is_ok());
        assert!(LocationStorage::for_kind(ChartKind::Pie2d).is_err());
    }

    #[test]
    fn all_locations_round_trip_and_top_removes_the_override() {
        let original = tsch::ChartSeriesStyleArchive::default().encode_to_vec();
        assert_eq!(
            read_location(&original, LocationStorage::Bar).unwrap(),
            ChartSeriesValueLabelLocation::Top
        );
        for location in [
            ChartSeriesValueLabelLocation::Bottom,
            ChartSeriesValueLabelLocation::Middle,
            ChartSeriesValueLabelLocation::Outside,
        ] {
            let patched = patch_location(&original, LocationStorage::Bar, location).unwrap();
            assert_eq!(
                read_location(&patched, LocationStorage::Bar).unwrap(),
                location
            );
            let reset = patch_location(
                &patched,
                LocationStorage::Bar,
                ChartSeriesValueLabelLocation::Top,
            )
            .unwrap();
            assert_eq!(
                read_location(&reset, LocationStorage::Bar).unwrap(),
                ChartSeriesValueLabelLocation::Top
            );
        }
    }

    #[test]
    fn unknown_fields_are_retained() {
        let mut generated = Vec::new();
        append_varint_field(&mut generated, 4_097, 42).unwrap();
        let mut original = Vec::new();
        append_varint_field(&mut original, 4_096, 42).unwrap();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        let patched = patch_location(
            &original,
            LocationStorage::Bar,
            ChartSeriesValueLabelLocation::Outside,
        )
        .unwrap();
        let fields = parse_wire_fields(&patched).unwrap();
        assert!(fields.iter().any(|field| field.number == 4_096));
        let extension = generated_chart_series_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert!(
            parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number == 4_097)
        );
    }

    #[test]
    fn malformed_or_unknown_locations_are_rejected() {
        for value in [1, 2, 3, 5, 7, 9] {
            let mut generated = Vec::new();
            append_varint_field(&mut generated, BAR_VALUE_LABEL_LOCATION_FIELD, value).unwrap();
            let mut data = Vec::new();
            append_length_delimited_field(
                &mut data,
                GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
                &generated,
            )
            .unwrap();
            assert!(read_location(&data, LocationStorage::Bar).is_err());
        }
    }
}
