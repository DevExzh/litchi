//! Lossless prefix and suffix CRUD for native chart-axis labels.
//!
//! Pages, Numbers, and Keynote persist axis-label affixes inside both the
//! legacy and current number formatter objects on an axis non-style archive.

use prost::Message;

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::charts::number_format::{DualNumberFormatFields, patch_dual_affixes, read_dual_affixes};
use crate::charts::{Axis, ChartLabelAffixes, ChartNumberFormat};
use crate::protobuf::tsch;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const AXIS_NUMBER_FORMAT_FIELDS: DualNumberFormatFields = DualNumberFormatFields {
    legacy: 2,
    format_type: 3,
    current: 42,
};
const FORMAT_CONTEXT: &str = "chart axis-label";

/// Read the text placed before and after labels on one native chart axis.
pub(crate) fn chart_axis_label_affixes(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
) -> Result<ChartLabelAffixes> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, read_axis_label_affixes)
}

/// Set or clear the text placed before and after labels on one chart axis.
pub(crate) fn set_chart_axis_label_affixes(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
    affixes: &ChartLabelAffixes,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, read_axis_label_affixes)? == *affixes {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_axis_label_affixes(data, affixes))?;
    if slot.read(package, read_axis_label_affixes)? != *affixes {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis label-affix update failed validation",
            axis.as_str()
        )));
    }
    Ok(())
}

fn read_axis_label_affixes(data: &[u8]) -> Result<ChartLabelAffixes> {
    let extension = generated_axis_non_style_extension(data)?;
    if let Some(extension) = extension {
        tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    }
    read_dual_affixes(extension, AXIS_NUMBER_FORMAT_FIELDS, FORMAT_CONTEXT)
}

fn patch_axis_label_affixes(data: &[u8], expected: &ChartLabelAffixes) -> Result<Vec<u8>> {
    let existing_extension = generated_axis_non_style_extension(data)?;
    if existing_extension.is_none() && expected.is_empty() {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let Some(patched_extension) = patch_dual_affixes(
        extension,
        AXIS_NUMBER_FORMAT_FIELDS,
        expected,
        ChartNumberFormat::AXIS_NATIVE_DEFAULT,
        FORMAT_CONTEXT,
    )?
    else {
        return Ok(data.to_vec());
    };
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        Some(patched_extension.as_slice()),
    )?;
    if read_axis_label_affixes(&patched)? != *expected {
        return Err(Error::InvalidFormat(
            "chart axis-label affix wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::{ChartDecimalPlaces, ChartNegativeStyle};
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    fn custom_affixes() -> ChartLabelAffixes {
        ChartLabelAffixes::new("USD ", " net")
    }

    #[test]
    fn app_authored_axis_affixes_round_trip_through_both_formatters() {
        let original = axis_non_style_with_unknown_fields();
        let customized = patch_axis_label_affixes(&original, &custom_affixes()).unwrap();
        assert_eq!(
            read_axis_label_affixes(&customized).unwrap(),
            custom_affixes()
        );
        let extension = generated_axis_non_style_extension(&customized)
            .unwrap()
            .unwrap();
        for field in [
            AXIS_NUMBER_FORMAT_FIELDS.legacy,
            AXIS_NUMBER_FORMAT_FIELDS.current,
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
    fn clearing_axis_affixes_preserves_number_format_and_unknown_fields() {
        let original = axis_non_style_with_unknown_fields();
        let format = ChartNumberFormat::new(
            ChartDecimalPlaces::fixed(2).unwrap(),
            ChartNegativeStyle::Parentheses,
            true,
        );
        let formatted =
            super::super::axis_number_format::patch_axis_number_format(&original, format).unwrap();
        let customized = patch_axis_label_affixes(&formatted, &custom_affixes()).unwrap();
        let cleared = patch_axis_label_affixes(&customized, &ChartLabelAffixes::default()).unwrap();
        assert_eq!(
            read_axis_label_affixes(&cleared).unwrap(),
            ChartLabelAffixes::default()
        );
        assert_eq!(
            super::super::axis_number_format::read_axis_number_format(&cleared).unwrap(),
            format
        );
        let extension = generated_axis_non_style_extension(&cleared)
            .unwrap()
            .unwrap();
        assert!(
            parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNKNOWN_GENERATED_FIELD)
        );
    }

    #[test]
    fn malformed_or_conflicting_axis_affixes_are_rejected() {
        let original = axis_non_style_with_unknown_fields();
        let customized = patch_axis_label_affixes(&original, &custom_affixes()).unwrap();
        let extension = generated_axis_non_style_extension(&customized)
            .unwrap()
            .unwrap();
        let legacy = parse_wire_fields(extension)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == AXIS_NUMBER_FORMAT_FIELDS.legacy)
            .unwrap();
        let mut duplicate_extension = extension.to_vec();
        append_length_delimited_field(
            &mut duplicate_extension,
            AXIS_NUMBER_FORMAT_FIELDS.legacy,
            &extension[legacy.payload_start()..legacy.end()],
        )
        .unwrap();
        let duplicate = patch_length_delimited_field(
            &customized,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(duplicate_extension.as_slice()),
        )
        .unwrap();
        assert!(read_axis_label_affixes(&duplicate).is_err());
    }

    fn axis_non_style_with_unknown_fields() -> Vec<u8> {
        let mut generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisvaluenumberofmajorgridlines: Some(5),
            tschchartaxisvalueshowlabels: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 73).unwrap();
        let mut outer = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 91).unwrap();
        outer
    }
}
