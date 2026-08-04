//! Lossless per-series value-label number-format CRUD for native charts.
//!
//! iWork stores a legacy and a current formatter for every customized chart
//! series. The shared chart-number-format codec validates that both
//! representations agree and patches them together without disturbing
//! affixes or unrelated wire fields.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::number_format::{
    DualNumberFormatFields, patch_dual_number_format, read_dual_number_format,
};
use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::{Error, IWorkPackage, Result};

pub use crate::charts::number_format::{
    ChartDecimalPlaces as ChartSeriesValueLabelDecimalPlaces,
    ChartFixedDecimalPlaces as ChartSeriesValueLabelFixedDecimalPlaces,
    ChartNegativeStyle as ChartSeriesValueLabelNegativeStyle,
    ChartNumberFormat as ChartSeriesValueLabelNumberFormat,
};

const SERIES_NUMBER_FORMAT_FIELDS: DualNumberFormatFields = DualNumberFormatFields {
    legacy: 21,
    format_type: 23,
    current: 98,
};
const FORMAT_CONTEXT: &str = "chart series value-label";

/// Read number formats in native series order.
pub(crate) fn chart_series_value_label_number_formats(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    series_count: usize,
) -> Result<Vec<ChartSeriesValueLabelNumberFormat>> {
    ensure_supported_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
        read_number_format,
    )
}

/// Set number formats in native series order.
pub(crate) fn set_chart_series_value_label_number_formats(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    expected: &[ChartSeriesValueLabelNumberFormat],
) -> Result<()> {
    ensure_supported_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series value-label number formats",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
        read_number_format,
        patch_number_format,
    )
}

fn ensure_supported_kind(kind: ChartKind) -> Result<()> {
    if matches!(kind, ChartKind::Undefined | ChartKind::Unsupported(_)) {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} has no supported series value-label number format"
        )));
    }
    Ok(())
}

fn read_number_format(data: &[u8]) -> Result<ChartSeriesValueLabelNumberFormat> {
    let extension = generated_chart_series_non_style_extension(data)?;
    if let Some(extension) = extension {
        tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    }
    read_dual_number_format(
        extension,
        SERIES_NUMBER_FORMAT_FIELDS,
        ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
        FORMAT_CONTEXT,
    )
}

fn patch_number_format(
    data: &[u8],
    expected: &ChartSeriesValueLabelNumberFormat,
) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none()
        && *expected == ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT
    {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let Some(extension) = patch_dual_number_format(
        extension,
        SERIES_NUMBER_FORMAT_FIELDS,
        *expected,
        ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
        FORMAT_CONTEXT,
    )?
    else {
        return Ok(data.to_vec());
    };
    let patched = patch_chart_series_non_style_extension(
        data,
        existing_extension.is_some(),
        Some(extension.as_slice()),
    )?;
    if read_number_format(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series value-label number-format wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::wire::{append_varint_field, parse_wire_fields};

    fn custom_format() -> ChartSeriesValueLabelNumberFormat {
        ChartSeriesValueLabelNumberFormat::new(
            ChartSeriesValueLabelDecimalPlaces::fixed(2).unwrap(),
            ChartSeriesValueLabelNegativeStyle::Parentheses,
            false,
        )
    }

    #[test]
    fn number_format_round_trips_through_both_native_representations() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let expected = custom_format();
        let patched = patch_number_format(&original, &expected).unwrap();
        assert_eq!(read_number_format(&patched).unwrap(), expected);

        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        for field in [
            SERIES_NUMBER_FORMAT_FIELDS.legacy,
            SERIES_NUMBER_FORMAT_FIELDS.current,
        ] {
            assert_eq!(
                parse_wire_fields(extension)
                    .unwrap()
                    .iter()
                    .filter(|wire| wire.number == field)
                    .count(),
                1
            );
        }
    }

    #[test]
    fn number_format_patch_preserves_unknown_extension_fields() {
        const UNKNOWN_FIELD: u32 = 9_001;
        let mut extension = Vec::new();
        append_varint_field(&mut extension, UNKNOWN_FIELD, 73).unwrap();
        let original = patch_chart_series_non_style_extension(
            &canonical_empty_chart_series_non_style_data().unwrap(),
            false,
            Some(extension.as_slice()),
        )
        .unwrap();

        let patched = patch_number_format(&original, &custom_format()).unwrap();
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        let unknown = parse_wire_fields(extension)
            .unwrap()
            .into_iter()
            .find(|field| field.number == UNKNOWN_FIELD)
            .unwrap();
        assert_eq!(
            &extension[unknown.key_end..unknown.end],
            litchi_iwa_common::varint::encode_varint(73)
        );
    }

    #[test]
    fn default_format_does_not_materialize_sparse_storage() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(
            patch_number_format(
                &original,
                &ChartSeriesValueLabelNumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT
            )
            .unwrap(),
            original
        );
    }
}
