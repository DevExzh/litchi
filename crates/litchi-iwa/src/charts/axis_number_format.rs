//! Lossless decimal-number-format CRUD for native chart axes.
//!
//! iWork stores both legacy and current format payloads on the selected axis
//! non-style object. The default is automatic decimal places, a minus sign,
//! and no thousands separator.

use prost::Message;

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::charts::number_format::{
    DualNumberFormatFields, clear_dual_number_format, patch_dual_number_format,
    read_dual_number_format,
};
use crate::charts::{ChartAxis, ChartNumberFormat};
use crate::protobuf::tsch;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const AXIS_NUMBER_FORMAT_FIELDS: DualNumberFormatFields = DualNumberFormatFields {
    legacy: 2,
    format_type: 3,
    current: 42,
};
const FORMAT_CONTEXT: &str = "chart axis-label";

/// Read the decimal-number format for one native chart axis.
pub(crate) fn chart_axis_number_format(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
) -> Result<ChartNumberFormat> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, read_axis_number_format)
}

/// Set or reset the decimal-number format for one native chart axis.
pub(crate) fn set_chart_axis_number_format(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: ChartAxis,
    format: ChartNumberFormat,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, read_axis_number_format)? == format {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_axis_number_format(data, format))?;
    if slot.read(package, read_axis_number_format)? != format {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis number-format update failed validation",
            axis.label()
        )));
    }
    Ok(())
}

fn read_axis_number_format(data: &[u8]) -> Result<ChartNumberFormat> {
    let extension = generated_axis_non_style_extension(data)?;
    if let Some(extension) = extension {
        tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    }
    read_dual_number_format(
        extension,
        AXIS_NUMBER_FORMAT_FIELDS,
        ChartNumberFormat::AXIS_NATIVE_DEFAULT,
        FORMAT_CONTEXT,
    )
}

fn patch_axis_number_format(data: &[u8], expected: ChartNumberFormat) -> Result<Vec<u8>> {
    let existing_extension = generated_axis_non_style_extension(data)?;
    if existing_extension.is_none() && expected == ChartNumberFormat::AXIS_NATIVE_DEFAULT {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let patched_extension = if expected == ChartNumberFormat::AXIS_NATIVE_DEFAULT {
        clear_dual_number_format(extension, AXIS_NUMBER_FORMAT_FIELDS, FORMAT_CONTEXT)?
    } else {
        patch_dual_number_format(
            extension,
            AXIS_NUMBER_FORMAT_FIELDS,
            expected,
            ChartNumberFormat::AXIS_NATIVE_DEFAULT,
            FORMAT_CONTEXT,
        )?
    };
    let Some(patched_extension) = patched_extension else {
        return Ok(data.to_vec());
    };
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        Some(patched_extension.as_slice()),
    )?;
    if read_axis_number_format(&patched)? != expected {
        return Err(Error::InvalidFormat(
            "chart axis-label number-format wire patch failed validation".to_owned(),
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

    fn custom_format() -> ChartNumberFormat {
        ChartNumberFormat::new(
            ChartDecimalPlaces::fixed(2).unwrap(),
            ChartNegativeStyle::Parentheses,
            true,
        )
    }

    #[test]
    fn app_authored_axis_number_format_round_trips_exactly() {
        let original = axis_non_style_with_unknown_fields();
        let customized = patch_axis_number_format(&original, custom_format()).unwrap();
        assert_eq!(
            read_axis_number_format(&customized).unwrap(),
            custom_format()
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
                    .filter(|wire| wire.number == field)
                    .count(),
                1
            );
        }
    }

    #[test]
    fn reset_removes_only_axis_number_format_fields() {
        let original = axis_non_style_with_unknown_fields();
        let customized = patch_axis_number_format(&original, custom_format()).unwrap();
        let reset =
            patch_axis_number_format(&customized, ChartNumberFormat::AXIS_NATIVE_DEFAULT).unwrap();
        assert_eq!(reset, original);
    }

    #[test]
    fn malformed_or_conflicting_axis_number_formats_are_rejected() {
        let original = axis_non_style_with_unknown_fields();
        let customized = patch_axis_number_format(&original, custom_format()).unwrap();
        let extension = generated_axis_non_style_extension(&customized)
            .unwrap()
            .unwrap();
        let legacy = parse_wire_fields(extension)
            .unwrap()
            .into_iter()
            .find(|field| field.number == AXIS_NUMBER_FORMAT_FIELDS.legacy)
            .unwrap();
        let mut duplicate_extension = extension.to_vec();
        append_length_delimited_field(
            &mut duplicate_extension,
            AXIS_NUMBER_FORMAT_FIELDS.legacy,
            &extension[legacy.payload_start..legacy.end],
        )
        .unwrap();
        let duplicate = patch_length_delimited_field(
            &customized,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            true,
            Some(duplicate_extension.as_slice()),
        )
        .unwrap();
        assert!(read_axis_number_format(&duplicate).is_err());
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
