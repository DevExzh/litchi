//! Lossless per-series value-label prefix and suffix CRUD for native charts.
//!
//! Series value labels and axis labels use the same dual native formatter
//! representation. The shared formatter codec keeps both copies synchronized
//! while preserving number-format settings and unknown fields.

use prost::Message;

use litchi_iwa_common::chart::number_format::LabelAffixes;

use crate::charts::Kind;
use crate::charts::number_format::{DualNumberFormatFields, patch_dual_affixes, read_dual_affixes};
use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::{Error, IWorkPackage, Result};

const SERIES_NUMBER_FORMAT_FIELDS: DualNumberFormatFields = DualNumberFormatFields {
    legacy: 21,
    format_type: 23,
    current: 98,
};
const FORMAT_CONTEXT: &str = "chart series value-label";

/// Read value-label affixes in native series order.
pub(crate) fn chart_series_value_label_affixes(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    series_count: usize,
) -> Result<Vec<LabelAffixes>> {
    ensure_supported_kind(kind)?;
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        LabelAffixes::default(),
        read_affixes,
    )
}

/// Set value-label affixes in native series order.
pub(crate) fn set_chart_series_value_label_affixes(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    expected: &[LabelAffixes],
) -> Result<()> {
    ensure_supported_kind(kind)?;
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "series value-label affixes",
        NewChartSeriesNonStyleBase::Styled,
        expected,
        LabelAffixes::default(),
        read_affixes,
        patch_affixes,
    )
}

fn ensure_supported_kind(kind: Kind) -> Result<()> {
    if kind == Kind::Undefined || kind.is_unsupported() {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} has no supported series value-label affixes"
        )));
    }
    Ok(())
}

fn read_affixes(data: &[u8]) -> Result<LabelAffixes> {
    let extension = generated_chart_series_non_style_extension(data)?;
    if let Some(extension) = extension {
        tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    }
    read_dual_affixes(extension, SERIES_NUMBER_FORMAT_FIELDS, FORMAT_CONTEXT)
}

fn patch_affixes(data: &[u8], expected: &LabelAffixes) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_non_style_extension(data)?;
    if existing_extension.is_none() && expected.is_empty() {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let Some(extension) = patch_dual_affixes(
        extension,
        SERIES_NUMBER_FORMAT_FIELDS,
        expected,
        crate::charts::NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
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
    if read_affixes(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart series value-label affix wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::canonical_empty_chart_series_non_style_data;
    use crate::charts::{DecimalPlaces, NegativeStyle, NumberFormat};
    use crate::wire::{append_varint_field, parse_wire_fields};

    fn custom_affixes() -> LabelAffixes {
        LabelAffixes::new("$", " net").unwrap()
    }

    #[test]
    fn affixes_round_trip_through_both_native_representations() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let expected = custom_affixes();
        let patched = patch_affixes(&original, &expected).unwrap();
        assert_eq!(read_affixes(&patched).unwrap(), expected);

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
                    .filter(|wire| wire.number() == field)
                    .count(),
                1
            );
        }
    }

    #[test]
    fn affix_patch_preserves_number_format_and_unknown_fields() {
        const UNKNOWN_FIELD: u32 = 9_001;
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        let mut extension = Vec::new();
        append_varint_field(&mut extension, UNKNOWN_FIELD, 73).unwrap();
        let original =
            patch_chart_series_non_style_extension(&original, false, Some(extension.as_slice()))
                .unwrap();
        let number_format = NumberFormat::new(
            DecimalPlaces::fixed(3).unwrap(),
            NegativeStyle::Parentheses,
            false,
        );
        let extension = generated_chart_series_non_style_extension(&original)
            .unwrap()
            .unwrap();
        let extension = crate::charts::number_format::patch_dual_number_format(
            extension,
            SERIES_NUMBER_FORMAT_FIELDS,
            number_format,
            NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
            FORMAT_CONTEXT,
        )
        .unwrap()
        .unwrap();
        let formatted =
            patch_chart_series_non_style_extension(&original, true, Some(extension.as_slice()))
                .unwrap();

        let patched = patch_affixes(&formatted, &custom_affixes()).unwrap();
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            crate::charts::number_format::read_dual_number_format(
                Some(extension),
                SERIES_NUMBER_FORMAT_FIELDS,
                NumberFormat::SERIES_VALUE_LABEL_NATIVE_DEFAULT,
                FORMAT_CONTEXT,
            )
            .unwrap(),
            number_format
        );
        assert!(
            parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNKNOWN_FIELD)
        );
    }

    #[test]
    fn empty_affixes_do_not_materialize_sparse_storage() {
        let original = canonical_empty_chart_series_non_style_data().unwrap();
        assert_eq!(
            patch_affixes(&original, &LabelAffixes::default()).unwrap(),
            original
        );
    }
}
