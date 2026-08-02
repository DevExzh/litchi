//! Lossless native chart-border storage and mutation.
//!
//! iWork stores the Chart Options `Border` switch in the generated extension
//! of a chart's `TSCH.ChartStyleArchive`. This module resolves the private
//! style object, preserves both protobuf layers losslessly, and changes only
//! the native border switch.

use prost::Message;

use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultshowborder` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BORDER_VISIBLE_FIELD: u32 = 18;

/// Read whether one native chart shows its chart-area border.
pub(crate) fn chart_border_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_border_visible)
}

/// Set whether one native chart shows its chart-area border.
pub(crate) fn set_chart_border_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_border_visible)? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_border_visibility(data, visible))?;
    if slot.read(package, read_chart_border_visible)? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} border update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_border_visible(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    Ok(generated.tschchartinfodefaultshowborder.unwrap_or(false))
}

fn patch_chart_border_visibility(data: &[u8], visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if !visible {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            ..Default::default()
        };
        let extension = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_chart_border_visibility(&patched, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let visible_present = generated.tschchartinfodefaultshowborder.is_some();
    let value = (visible_present || visible).then_some(u64::from(visible));
    let extension = patch_varint_field(
        extension,
        CHART_BORDER_VISIBLE_FIELD,
        visible_present,
        value,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_border_visibility(&patched, visible)?;
    Ok(patched)
}

fn validate_patched_chart_border_visibility(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_border_visible(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart border wire patch failed validation".to_owned(),
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
    fn border_patch_retains_other_style_fields_and_unmapped_data() {
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(false),
            tschchartinfodefaultgridbackgroundopacity: Some(1.0),
            tschchartinfodefaultinterbargap: Some(0.2),
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

        let visible = patch_chart_border_visibility(&original, true).unwrap();
        assert!(read_chart_border_visible(&visible).unwrap());
        let patched_generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&visible).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(
            patched_generated.tschchartinfodefaultgridbackgroundopacity,
            Some(1.0)
        );
        assert_eq!(patched_generated.tschchartinfodefaultinterbargap, Some(0.2));
        assert_eq!(
            raw_field(&visible, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_style_extension(&visible).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(&original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );

        let restored = patch_chart_border_visibility(&visible, false).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn borders_default_hidden_and_create_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert!(!read_chart_border_visible(&original).unwrap());
        assert_eq!(
            patch_chart_border_visibility(&original, false).unwrap(),
            original
        );

        let visible = patch_chart_border_visibility(&original, true).unwrap();
        assert!(read_chart_border_visible(&visible).unwrap());
        assert!(generated_chart_style_extension(&visible).unwrap().is_some());
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number == number)
            .map(|field| data[field.start..field.end].to_vec())
            .collect()
    }
}
