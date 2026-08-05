//! Lossless native chart-gap storage and mutation.
//!
//! iWork stores the `Between Columns`/`Between Bars` and `Between Sets`
//! inspector percentages in the generated extension of
//! `TSCH.ChartStyleArchive`. The public types model those percentages directly
//! while this module preserves every unrelated protobuf field byte-for-byte.

use prost::Message;

use litchi_iwa_common::chart::gaps::{Percentage, Spacing};

use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultinterbargap` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BETWEEN_ITEMS_GAP_FIELD: u32 = 16;
/// `tschchartinfodefaultintersetgap` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BETWEEN_SETS_GAP_FIELD: u32 = 17;
/// Read the gap spacing of one native chart.
pub(crate) fn chart_gap_spacing(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Spacing> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_gap_spacing)
}

/// Set the gap spacing of one native chart.
pub(crate) fn set_chart_gap_spacing(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    spacing: Spacing,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_gap_spacing)? == spacing {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_gap_spacing(data, spacing))?;
    if slot.read(package, read_chart_gap_spacing)? != spacing {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} gap update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_gap_spacing(data: &[u8]) -> Result<Spacing> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(Spacing::DEFAULT);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    Ok(Spacing::new(
        generated
            .tschchartinfodefaultinterbargap
            .map(native_percentage)
            .transpose()?
            .unwrap_or(Spacing::DEFAULT.between_items()),
        generated
            .tschchartinfodefaultintersetgap
            .map(native_percentage)
            .transpose()?
            .unwrap_or(Spacing::DEFAULT.between_sets()),
    ))
}

fn native_percentage(value: f32) -> Result<Percentage> {
    Percentage::new(value).map_err(|error| {
        Error::InvalidFormat(format!("native chart gap {value} is invalid: {error}"))
    })
}

fn patch_chart_gap_spacing(data: &[u8], spacing: Spacing) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if spacing == Spacing::DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultinterbargap: (spacing.between_items()
                != Spacing::DEFAULT.between_items())
            .then(|| spacing.between_items().percent()),
            tschchartinfodefaultintersetgap: (spacing.between_sets()
                != Spacing::DEFAULT.between_sets())
            .then(|| spacing.between_sets().percent()),
            ..Default::default()
        };
        let extension = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_chart_gap_spacing(&patched, spacing)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let between_items_present = generated.tschchartinfodefaultinterbargap.is_some();
    let between_sets_present = generated.tschchartinfodefaultintersetgap.is_some();
    let between_items = if spacing == Spacing::DEFAULT {
        None
    } else {
        (between_items_present || spacing.between_items() != Spacing::DEFAULT.between_items())
            .then(|| spacing.between_items().percent().to_bits())
    };
    let between_sets = if spacing == Spacing::DEFAULT {
        None
    } else {
        (between_sets_present || spacing.between_sets() != Spacing::DEFAULT.between_sets())
            .then(|| spacing.between_sets().percent().to_bits())
    };
    let extension = patch_fixed32_field(
        extension,
        CHART_BETWEEN_ITEMS_GAP_FIELD,
        between_items_present,
        between_items,
    )?;
    let extension = patch_fixed32_field(
        &extension,
        CHART_BETWEEN_SETS_GAP_FIELD,
        between_sets_present,
        between_sets,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_gap_spacing(&patched, spacing)?;
    Ok(patched)
}

fn validate_patched_chart_gap_spacing(data: &[u8], expected: Spacing) -> Result<()> {
    if read_chart_gap_spacing(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart gap wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn gap_patch_retains_other_style_fields_and_unmapped_data() {
        let original_spacing = spacing(15.0, 45.0);
        let replacement = spacing(25.0, 70.0);
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(original_spacing.between_items().percent()),
            tschchartinfodefaultintersetgap: Some(original_spacing.between_sets().percent()),
            tschchartinfodefaultroundedcornerradius: Some(0.2),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let patched = patch_chart_gap_spacing(&original, replacement).unwrap();
        assert_eq!(read_chart_gap_spacing(&patched).unwrap(), replacement);
        let patched_generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&patched).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(patched_generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(
            patched_generated.tschchartinfodefaultroundedcornerradius,
            Some(0.2)
        );
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_style_extension(&patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(&original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );

        let restored = patch_chart_gap_spacing(&patched, original_spacing).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn gaps_default_natively_and_create_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let replacement = spacing(25.0, 70.0);

        assert_eq!(read_chart_gap_spacing(&original).unwrap(), Spacing::DEFAULT);
        assert_eq!(
            patch_chart_gap_spacing(&original, Spacing::DEFAULT).unwrap(),
            original
        );

        let patched = patch_chart_gap_spacing(&original, replacement).unwrap();
        assert_eq!(read_chart_gap_spacing(&patched).unwrap(), replacement);
        assert!(generated_chart_style_extension(&patched).unwrap().is_some());
    }

    #[test]
    fn resetting_gaps_retains_other_style_fields() {
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(25.0),
            tschchartinfodefaultintersetgap: Some(70.0),
            ..Default::default()
        });

        let reset = patch_chart_gap_spacing(&original, Spacing::DEFAULT).unwrap();
        assert_eq!(read_chart_gap_spacing(&reset).unwrap(), Spacing::DEFAULT);
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&reset).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultinterbargap, None);
        assert_eq!(generated.tschchartinfodefaultintersetgap, None);
        assert_unknown_fields_retained(&original, &reset);
    }

    #[test]
    fn malformed_native_gap_values_are_rejected() {
        let negative = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultinterbargap: Some(-0.1),
            ..Default::default()
        });
        assert!(read_chart_gap_spacing(&negative).is_err());

        let excessive = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultintersetgap: Some(1_000.0),
            ..Default::default()
        });
        assert!(read_chart_gap_spacing(&excessive).is_err());
    }

    fn spacing(between_items: f32, between_sets: f32) -> Spacing {
        Spacing::new(
            Percentage::new(between_items).unwrap(),
            Percentage::new(between_sets).unwrap(),
        )
    }

    fn style_with_unknown_fields(generated: tsch::generated::ChartStyleArchive) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut data = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(&mut data, GENERATED_CHART_STYLE_EXTENSION_FIELD, &extension)
            .unwrap();
        append_varint_field(&mut data, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        data
    }

    fn assert_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        assert_eq!(
            raw_field(patched, UNMAPPED_OUTER_FIELD),
            raw_field(original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_style_extension(patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
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
