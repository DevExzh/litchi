//! Lossless native chart-legend fill storage and mutation.
//!
//! Legend fill is an independently inherited `TSD.FillArchive` in the native
//! legend-style extension. The exact override state is retained so callers can
//! distinguish inheritance from an explicit “No Fill”.

use prost::Message;

use crate::charts::legend_style::{
    GENERATED_LEGEND_STYLE_EXTENSION_FIELD, LegendStyleSlot, generated_legend_style_extension,
    legend_style_slot,
};
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::component_identifier_for_entry;
use crate::protobuf::tsch;
use crate::shapes::{
    ShapeFill, fill_from_native, fill_to_native, image_data_identifier,
    remove_orphaned_image_asset, validate_image_asset,
};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschlegendmodeldefaultfill` in `TSCH.Generated.LegendStyleArchive`.
const LEGEND_FILL_FIELD: u32 = 1;

/// Exact direct fill state for a native chart legend.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum ChartLegendFill {
    /// No direct override; iWork resolves the legend style's parent chain.
    #[default]
    Inherited,
    /// A direct fill override. `ShapeFill::None` is the explicit “No Fill”.
    Fill(ShapeFill),
}

/// Read the exact direct legend-fill state of one native chart.
pub(crate) fn chart_legend_fill(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendFill> {
    legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_legend_fill)
}

/// Set or remove the direct legend-fill override of one native chart.
pub(crate) fn set_chart_legend_fill(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    fill: &ChartLegendFill,
) -> Result<()> {
    let slot = legend_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = slot.read(package, read_legend_fill)?;
    if &current == fill {
        return Ok(());
    }
    if let ChartLegendFill::Fill(fill) = fill {
        validate_image_asset(package, fill)?;
    }
    let old_data_identifier = direct_image_data_identifier(&current);
    let new_data_identifier = direct_image_data_identifier(fill);
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_legend_fill(data, fill))?;
    adjust_legend_style_data_reference(package, &slot, old_data_identifier, new_data_identifier)?;
    if slot.read(package, read_legend_fill)? != *fill {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend fill update failed validation"
        )));
    }
    remove_orphaned_image_asset(package, old_data_identifier)?;
    Ok(())
}

fn read_legend_fill(data: &[u8]) -> Result<ChartLegendFill> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        return Ok(ChartLegendFill::Inherited);
    };
    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    generated
        .tschlegendmodeldefaultfill
        .as_ref()
        .map(fill_from_native)
        .transpose()
        .map(|fill| fill.map_or(ChartLegendFill::Inherited, ChartLegendFill::Fill))
}

fn patch_legend_fill(data: &[u8], fill: &ChartLegendFill) -> Result<Vec<u8>> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        let ChartLegendFill::Fill(fill) = fill else {
            return Ok(data.to_vec());
        };
        let generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultfill: Some(fill_to_native(fill)),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_legend_fill(&patched, &ChartLegendFill::Fill(fill.clone()))?;
        return Ok(patched);
    };

    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let fill_present = generated.tschlegendmodeldefaultfill.is_some();
    let native = match fill {
        ChartLegendFill::Inherited => None,
        ChartLegendFill::Fill(fill) => Some(fill_to_native(fill).encode_to_vec()),
    };
    let extension = patch_length_delimited_field(
        extension,
        LEGEND_FILL_FIELD,
        fill_present,
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_fill(&patched, fill)?;
    Ok(patched)
}

fn adjust_legend_style_data_reference(
    package: &mut IWorkPackage,
    slot: &LegendStyleSlot,
    old_data_identifier: Option<u64>,
    new_data_identifier: Option<u64>,
) -> Result<()> {
    if old_data_identifier == new_data_identifier {
        return Ok(());
    }
    let component_id = component_identifier_for_entry(package, slot.archive_name())?
        .ok_or_else(|| Error::InvalidFormat("Legend style has no owning component".to_owned()))?;
    if let Some(identifier) = old_data_identifier {
        remove_component_data_reference(package, component_id, identifier, slot.object_id())?;
    }
    if let Some(identifier) = new_data_identifier {
        add_component_data_reference(package, component_id, identifier, slot.object_id())?;
    }
    Ok(())
}

fn direct_image_data_identifier(fill: &ChartLegendFill) -> Option<u64> {
    match fill {
        ChartLegendFill::Inherited => None,
        ChartLegendFill::Fill(fill) => image_data_identifier(fill),
    }
}

fn validate_patched_legend_fill(data: &[u8], expected: &ChartLegendFill) -> Result<()> {
    if &read_legend_fill(data)? != expected {
        return Err(Error::InvalidFormat(
            "Chart legend fill wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{RgbColorSpace, RgbaColor};
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn legend_fill_is_exact_and_preserves_unknown_fields() {
        let mut generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultopacity: Some(0.8),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::LegendStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        assert_eq!(
            read_legend_fill(&original).unwrap(),
            ChartLegendFill::Inherited
        );
        let solid = ChartLegendFill::Fill(ShapeFill::Solid(
            RgbaColor::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
        ));
        let colored = patch_legend_fill(&original, &solid).unwrap();
        assert_eq!(read_legend_fill(&colored).unwrap(), solid);
        let generated_after = generated_legend_style_extension(&colored).unwrap().unwrap();
        assert_eq!(
            tsch::generated::LegendStyleArchive::decode(generated_after)
                .unwrap()
                .tschlegendmodeldefaultopacity,
            Some(0.8)
        );
        assert!(
            parse_wire_fields(&colored)
                .unwrap()
                .iter()
                .any(|field| { field.number == UNMAPPED_OUTER_FIELD && field.wire_type == 0 })
        );
        assert!(
            parse_wire_fields(generated_after)
                .unwrap()
                .iter()
                .any(|field| field.number == UNMAPPED_GENERATED_FIELD && field.wire_type == 0)
        );

        let no_fill = patch_legend_fill(&colored, &ChartLegendFill::Fill(ShapeFill::None)).unwrap();
        assert_eq!(
            read_legend_fill(&no_fill).unwrap(),
            ChartLegendFill::Fill(ShapeFill::None)
        );
        let inherited = patch_legend_fill(&no_fill, &ChartLegendFill::Inherited).unwrap();
        assert_eq!(inherited, original);
    }
}
