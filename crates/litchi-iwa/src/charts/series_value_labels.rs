//! Lossless per-series value-label visibility CRUD for native charts.
//!
//! iWork stores value-label visibility in sparse series non-style objects,
//! using a distinct generated boolean field for each chart family. Formatter
//! metadata lives in neighboring fields and is deliberately left untouched.

use prost::Message;

use crate::charts::Kind;
use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_varint_field};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_common::chart::series_labels::Visibility;

const AREA_VALUE_LABELS_FIELD: u32 = 38;
const BAR_VALUE_LABELS_FIELD: u32 = 39;
const BUBBLE_VALUE_LABELS_FIELD: u32 = 40;
const LINE_VALUE_LABELS_FIELD: u32 = 42;
const MIXED_VALUE_LABELS_FIELD: u32 = 43;
const PIE_VALUE_LABELS_FIELD: u32 = 44;
const SCATTER_VALUE_LABELS_FIELD: u32 = 45;
const RADAR_VALUE_LABELS_FIELD: u32 = 162;

/// Read value-label visibility for every series in native series order.
pub(crate) fn chart_series_value_label_visibilities(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    series_count: usize,
) -> Result<Vec<Visibility>> {
    let storage = ValueLabelStorage::for_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        storage.default_visibility(),
        |data| read_series_value_label_visibility(data, storage),
    )
}

/// Set value-label visibility for every series in native series order.
pub(crate) fn set_chart_series_value_label_visibilities(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    expected: &[Visibility],
) -> Result<()> {
    let storage = ValueLabelStorage::for_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series value-label visibility",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        storage.default_visibility(),
        |data| read_series_value_label_visibility(data, storage),
        |data, visibility| patch_series_value_label_visibility(data, storage, *visibility),
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ValueLabelStorage {
    Area,
    Bar,
    Bubble,
    Line,
    Mixed,
    Pie,
    Radar,
    Scatter,
}

impl ValueLabelStorage {
    fn for_kind(kind: Kind) -> Result<Self> {
        match kind {
            Kind::Area2d | Kind::Area3d | Kind::StackedArea2d | Kind::StackedArea3d => {
                Ok(Self::Area)
            },
            Kind::Column2d
            | Kind::Column3d
            | Kind::Bar2d
            | Kind::Bar3d
            | Kind::StackedColumn2d
            | Kind::StackedColumn3d
            | Kind::StackedBar2d
            | Kind::StackedBar3d
            | Kind::MultiDataColumn2d
            | Kind::MultiDataBar2d => Ok(Self::Bar),
            Kind::Bubble2d | Kind::MultiDataBubble2d => Ok(Self::Bubble),
            Kind::Line2d | Kind::Line3d => Ok(Self::Line),
            Kind::Mixed2d | Kind::TwoAxis2d => Ok(Self::Mixed),
            Kind::Pie2d | Kind::Pie3d | Kind::Donut2d | Kind::Donut3d => Ok(Self::Pie),
            Kind::Radar2d => Ok(Self::Radar),
            Kind::Scatter2d | Kind::MultiDataScatter2d => Ok(Self::Scatter),
            _ => Err(Error::InvalidFormat(format!(
                "chart kind {kind:?} has no supported series value labels"
            ))),
        }
    }

    const fn field_number(self) -> u32 {
        match self {
            Self::Area => AREA_VALUE_LABELS_FIELD,
            Self::Bar => BAR_VALUE_LABELS_FIELD,
            Self::Bubble => BUBBLE_VALUE_LABELS_FIELD,
            Self::Line => LINE_VALUE_LABELS_FIELD,
            Self::Mixed => MIXED_VALUE_LABELS_FIELD,
            Self::Pie => PIE_VALUE_LABELS_FIELD,
            Self::Radar => RADAR_VALUE_LABELS_FIELD,
            Self::Scatter => SCATTER_VALUE_LABELS_FIELD,
        }
    }

    const fn default_visibility(self) -> Visibility {
        match self {
            Self::Pie => Visibility::Visible,
            Self::Area
            | Self::Bar
            | Self::Bubble
            | Self::Line
            | Self::Mixed
            | Self::Radar
            | Self::Scatter => Visibility::Hidden,
        }
    }
}

fn read_series_value_label_visibility(
    data: &[u8],
    storage: ValueLabelStorage,
) -> Result<Visibility> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(storage.default_visibility());
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    Ok(strict_optional_bool(extension, storage.field_number())?
        .map(Visibility::from)
        .unwrap_or_else(|| storage.default_visibility()))
}

fn patch_series_value_label_visibility(
    data: &[u8],
    storage: ValueLabelStorage,
    visibility: Visibility,
) -> Result<Vec<u8>> {
    let field_number = storage.field_number();
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        if visibility == storage.default_visibility() {
            return Ok(data.to_vec());
        }
        let mut extension = Vec::new();
        crate::wire::append_varint_field(
            &mut extension,
            field_number,
            u64::from(visibility.is_visible()),
        )?;
        let patched =
            patch_chart_series_non_style_extension(data, false, Some(extension.as_slice()))?;
        validate_patched_visibility(&patched, storage, visibility)?;
        return Ok(patched);
    };

    let field_present = strict_optional_bool(extension, field_number)?.is_some();
    let replacement =
        (visibility != storage.default_visibility()).then_some(u64::from(visibility.is_visible()));
    let extension = patch_varint_field(extension, field_number, field_present, replacement)?;
    let patched = patch_chart_series_non_style_extension(
        data,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_visibility(&patched, storage, visibility)?;
    Ok(patched)
}

fn strict_optional_bool(data: &[u8], field_number: u32) -> Result<Option<bool>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series value-label field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end()..field.end()])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "chart series value-label field {field_number} is invalid: {error}"
                ))
            })?;
    if consumed != 1 || consumed != field.end() - field.key_end() || value > 1 {
        return Err(Error::InvalidFormat(format!(
            "chart series value-label field {field_number} is not a canonical boolean"
        )));
    }
    Ok(Some(value == 1))
}

fn validate_patched_visibility(
    data: &[u8],
    storage: ValueLabelStorage,
    expected: Visibility,
) -> Result<()> {
    if read_series_value_label_visibility(data, storage)? != expected {
        return Err(Error::InvalidFormat(
            "chart series value-label wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::{
        GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
        canonical_empty_chart_series_non_style_data,
    };
    use crate::wire::{
        append_length_delimited_field, append_varint_field, parse_wire_fields,
        patch_length_delimited_field,
    };

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;
    const NON_CANONICAL_ZERO: &[u8] = &[0x80, 0x00];

    #[test]
    fn every_known_chart_family_has_typed_value_label_storage() {
        for raw in 1..=tsch::ChartType::RadarChartType2D as i32 {
            assert!(ValueLabelStorage::for_kind(Kind::from_native(raw)).is_ok());
        }
        assert!(ValueLabelStorage::for_kind(Kind::Undefined).is_err());
        assert!(ValueLabelStorage::for_kind(Kind::from_native(9_001)).is_err());
    }

    #[test]
    fn bar_value_labels_create_and_remove_a_minimal_override() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(
            read_series_value_label_visibility(&original, ValueLabelStorage::Bar).unwrap(),
            Visibility::Hidden
        );
        assert_eq!(
            patch_series_value_label_visibility(
                &original,
                ValueLabelStorage::Bar,
                Visibility::Hidden,
            )
            .unwrap(),
            original
        );

        let visible = patch_series_value_label_visibility(
            &original,
            ValueLabelStorage::Bar,
            Visibility::Visible,
        )
        .unwrap();
        assert_eq!(
            read_series_value_label_visibility(&visible, ValueLabelStorage::Bar).unwrap(),
            Visibility::Visible
        );
        assert_eq!(
            patch_series_value_label_visibility(
                &visible,
                ValueLabelStorage::Bar,
                Visibility::Hidden,
            )
            .unwrap(),
            original
        );
    }

    #[test]
    fn pie_value_labels_use_the_native_visible_default() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(
            read_series_value_label_visibility(&original, ValueLabelStorage::Pie).unwrap(),
            Visibility::Visible
        );
        let hidden = patch_series_value_label_visibility(
            &original,
            ValueLabelStorage::Pie,
            Visibility::Hidden,
        )
        .unwrap();
        assert_eq!(
            read_series_value_label_visibility(&hidden, ValueLabelStorage::Pie).unwrap(),
            Visibility::Hidden
        );
        assert_eq!(
            patch_series_value_label_visibility(
                &hidden,
                ValueLabelStorage::Pie,
                Visibility::Visible,
            )
            .unwrap(),
            original
        );
    }

    #[test]
    fn value_label_patch_preserves_unrelated_and_unmapped_fields() {
        let mut generated = Vec::new();
        append_varint_field(&mut generated, LINE_VALUE_LABELS_FIELD, 1).unwrap();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = canonical_empty_chart_series_non_style_data().unwrap();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let visible = patch_series_value_label_visibility(
            &original,
            ValueLabelStorage::Bar,
            Visibility::Visible,
        )
        .unwrap();
        let extension = generated_chart_series_non_style_extension(&visible)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_bool(extension, LINE_VALUE_LABELS_FIELD).unwrap(),
            Some(true)
        );
        assert_eq!(
            raw_field(extension, UNMAPPED_GENERATED_FIELD),
            raw_field(&generated, UNMAPPED_GENERATED_FIELD)
        );
        assert_eq!(
            raw_field(&visible, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
    }

    #[test]
    fn malformed_native_value_label_switches_are_rejected() {
        let base = canonical_empty_chart_series_non_style_data().unwrap();

        let mut duplicate = Vec::new();
        append_varint_field(&mut duplicate, BAR_VALUE_LABELS_FIELD, 0).unwrap();
        append_varint_field(&mut duplicate, BAR_VALUE_LABELS_FIELD, 1).unwrap();
        assert!(read_with_extension(&base, &duplicate).is_err());

        let mut wrong_wire = Vec::new();
        append_length_delimited_field(&mut wrong_wire, BAR_VALUE_LABELS_FIELD, &[]).unwrap();
        assert!(read_with_extension(&base, &wrong_wire).is_err());

        let mut non_boolean = Vec::new();
        append_varint_field(&mut non_boolean, BAR_VALUE_LABELS_FIELD, 2).unwrap();
        assert!(read_with_extension(&base, &non_boolean).is_err());

        let mut non_canonical = Vec::new();
        append_varint_field(&mut non_canonical, BAR_VALUE_LABELS_FIELD, 0).unwrap();
        assert_eq!(non_canonical.pop(), Some(0));
        non_canonical.extend_from_slice(NON_CANONICAL_ZERO);
        assert!(read_with_extension(&base, &non_canonical).is_err());
    }

    fn read_with_extension(base: &[u8], extension: &[u8]) -> Result<Visibility> {
        let data = patch_length_delimited_field(
            base,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(extension),
        )?;
        read_series_value_label_visibility(&data, ValueLabelStorage::Bar)
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number() == number)
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect()
    }
}
