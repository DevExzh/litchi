//! Lossless data-label visibility CRUD for native pie and donut charts.

use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_varint_field};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_common::chart::pie::LabelVisibility;

/// `tschchartseriespieshowserieslabels` in the generated series non-style.
const PIE_SHOW_DATA_POINT_NAMES_FIELD: u32 = 31;
/// `tschchartseriespieshowvaluelabels` in the generated series non-style.
const PIE_SHOW_VALUES_FIELD: u32 = 44;

/// Read label visibility for every wedge in chart-series order.
pub(crate) fn chart_pie_label_visibilities(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<LabelVisibility>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        LabelVisibility::DEFAULT,
        read_series_non_style_labels,
    )
}

/// Set label visibility for every wedge in chart-series order.
pub(crate) fn set_chart_pie_label_visibilities(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[LabelVisibility],
) -> Result<()> {
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "pie label visibility",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        LabelVisibility::DEFAULT,
        read_series_non_style_labels,
        |data, visibility| patch_series_non_style_labels(data, *visibility),
    )
}

fn read_series_non_style_labels(data: &[u8]) -> Result<LabelVisibility> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(LabelVisibility::DEFAULT);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    Ok(LabelVisibility::new(
        strict_optional_bool(extension, PIE_SHOW_DATA_POINT_NAMES_FIELD)?.unwrap_or(false),
        strict_optional_bool(extension, PIE_SHOW_VALUES_FIELD)?.unwrap_or(true),
    ))
}

fn patch_series_non_style_labels(data: &[u8], visibility: LabelVisibility) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        if visibility == LabelVisibility::DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespieshowserieslabels: visibility
                .data_point_names_visible()
                .then_some(true),
            tschchartseriespieshowvaluelabels: (!visibility.values_visible()).then_some(false),
            ..Default::default()
        };
        let patched = patch_chart_series_non_style_extension(
            data,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_labels(&patched, visibility)?;
        return Ok(patched);
    };

    let names_present = strict_optional_bool(extension, PIE_SHOW_DATA_POINT_NAMES_FIELD)?.is_some();
    let values_present = strict_optional_bool(extension, PIE_SHOW_VALUES_FIELD)?.is_some();
    let extension = patch_varint_field(
        extension,
        PIE_SHOW_DATA_POINT_NAMES_FIELD,
        names_present,
        visibility.data_point_names_visible().then_some(1),
    )?;
    let extension = patch_varint_field(
        &extension,
        PIE_SHOW_VALUES_FIELD,
        values_present,
        (!visibility.values_visible()).then_some(0),
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_labels(&patched, visibility)?;
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
            "singular chart pie label field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart pie label field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end()..field.end()])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "chart pie label field {field_number} is invalid: {error}"
                ))
            })?;
    if field.key_end() + consumed != field.end() || consumed != 1 || value > 1 {
        return Err(Error::InvalidFormat(format!(
            "chart pie label field {field_number} is not a canonical boolean"
        )));
    }
    Ok(Some(value == 1))
}

fn validate_patched_labels(data: &[u8], expected: LabelVisibility) -> Result<()> {
    if read_series_non_style_labels(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart pie label-visibility wire patch failed validation".to_owned(),
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

    #[test]
    fn label_visibility_has_explicit_native_defaults() {
        assert_eq!(LabelVisibility::default(), LabelVisibility::VALUES_ONLY);
        assert!(!LabelVisibility::HIDDEN.values_visible());
        assert!(LabelVisibility::DATA_POINT_NAMES_ONLY.data_point_names_visible());
        assert!(LabelVisibility::ALL.values_visible());
    }

    #[test]
    fn label_visibility_patch_is_lossless_and_resets_exactly() {
        let mut generated = tsch::generated::ChartSeriesNonStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = canonical_empty_chart_series_non_style_data().unwrap();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let patched =
            patch_series_non_style_labels(&original, LabelVisibility::DATA_POINT_NAMES_ONLY)
                .unwrap();
        assert_eq!(
            read_series_non_style_labels(&patched).unwrap(),
            LabelVisibility::DATA_POINT_NAMES_ONLY
        );
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_series_non_style_extension(&patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_series_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
        assert_eq!(
            patch_series_non_style_labels(&patched, LabelVisibility::DEFAULT).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_label_flags_are_rejected() {
        for (field, value) in [
            (PIE_SHOW_DATA_POINT_NAMES_FIELD, 2),
            (PIE_SHOW_VALUES_FIELD, 2),
        ] {
            let mut generated = Vec::new();
            append_varint_field(&mut generated, field, value).unwrap();
            let outer = patch_length_delimited_field(
                &canonical_empty_chart_series_non_style_data().unwrap(),
                GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
                false,
                Some(&generated),
            )
            .unwrap();
            assert!(read_series_non_style_labels(&outer).is_err());
        }

        for generated in [[0xf8, 0x01, 0x80, 0x00], [0xe0, 0x02, 0x80, 0x00]] {
            let outer = patch_length_delimited_field(
                &canonical_empty_chart_series_non_style_data().unwrap(),
                GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
                false,
                Some(&generated),
            )
            .unwrap();
            assert!(read_series_non_style_labels(&outer).is_err());
        }
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
