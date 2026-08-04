//! Lossless native radar-grid shape storage and mutation.
//!
//! Radar charts expose `Straight` and `Curved` gridlines in the Chart
//! inspector. iWork stores the selection as a generated chart-style boolean;
//! this module exposes a closed enum and preserves unrelated protobuf fields
//! byte-for-byte.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultradarradiusgridlinecurve` in
/// `TSCH.Generated.ChartStyleArchive`.
const RADAR_RADIUS_GRIDLINE_CURVE_FIELD: u32 = 29;
const STRAIGHT_NATIVE_VALUE: u64 = 0;
const CURVED_NATIVE_VALUE: u64 = 1;

/// Geometry used for the radius gridlines of a native radar chart.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartRadarGridShape {
    /// Straight segments form a polygon, the native default.
    #[default]
    Straight,
    /// Curved segments form concentric circles.
    Curved,
}

impl ChartRadarGridShape {
    const fn native(self) -> u64 {
        match self {
            Self::Straight => STRAIGHT_NATIVE_VALUE,
            Self::Curved => CURVED_NATIVE_VALUE,
        }
    }

    fn from_native(value: u64) -> Result<Self> {
        match value {
            STRAIGHT_NATIVE_VALUE => Ok(Self::Straight),
            CURVED_NATIVE_VALUE => Ok(Self::Curved),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native radar grid shape {value}"
            ))),
        }
    }
}

/// Read one radar chart's effective native grid shape.
pub(crate) fn chart_radar_grid_shape(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
) -> Result<ChartRadarGridShape> {
    require_radar_chart(kind)?;
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_radar_grid_shape)
}

/// Set one radar chart's native grid shape.
pub(crate) fn set_chart_radar_grid_shape(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    shape: ChartRadarGridShape,
) -> Result<()> {
    require_radar_chart(kind)?;
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_radar_grid_shape)? == shape {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_radar_grid_shape(data, shape))?;
    if slot.read(package, read_chart_radar_grid_shape)? != shape {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} radar grid-shape update failed validation"
        )));
    }
    Ok(())
}

fn require_radar_chart(kind: ChartKind) -> Result<()> {
    if !kind.supports_radar_grid_shape() {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} does not expose radar grid shape"
        )));
    }
    Ok(())
}

fn read_chart_radar_grid_shape(data: &[u8]) -> Result<ChartRadarGridShape> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(ChartRadarGridShape::Straight);
    };
    let value = read_optional_native_shape(extension)?;
    Ok(value
        .map(ChartRadarGridShape::from_native)
        .transpose()?
        .unwrap_or_default())
}

fn read_optional_native_shape(extension: &[u8]) -> Result<Option<u64>> {
    tsch::generated::ChartStyleArchive::decode(extension)?;
    let fields = parse_wire_fields(extension)?;
    let mut matches = fields
        .iter()
        .filter(|field| field.number == RADAR_RADIUS_GRIDLINE_CURVE_FIELD);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "radar grid-shape field {RADAR_RADIUS_GRIDLINE_CURVE_FIELD} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "radar grid-shape field {RADAR_RADIUS_GRIDLINE_CURVE_FIELD} is not a varint"
        )));
    }
    let (value, length) = litchi_iwa_common::varint::decode_varint_from_bytes(
        &extension[field.payload_start..field.end],
    )
    .map_err(|error| Error::InvalidFormat(format!("invalid radar grid-shape value: {error}")))?;
    if field.payload_start + length != field.end {
        return Err(Error::InvalidFormat(
            "radar grid-shape varint has trailing bytes".to_owned(),
        ));
    }
    Ok(Some(value))
}

fn patch_chart_radar_grid_shape(data: &[u8], shape: ChartRadarGridShape) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if shape == ChartRadarGridShape::Straight {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultradarradiusgridlinecurve: Some(true),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_radar_grid_shape(&patched, shape)?;
        return Ok(patched);
    };

    let present = read_optional_native_shape(extension)?.is_some();
    let native = (shape != ChartRadarGridShape::Straight).then(|| shape.native());
    let extension = patch_varint_field(
        extension,
        RADAR_RADIUS_GRIDLINE_CURVE_FIELD,
        present,
        native,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_radar_grid_shape(&patched, shape)?;
    Ok(patched)
}

fn validate_patched_chart_radar_grid_shape(
    data: &[u8],
    expected: ChartRadarGridShape,
) -> Result<()> {
    if read_chart_radar_grid_shape(data)? != expected {
        return Err(Error::InvalidFormat(
            "radar grid-shape wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;
    const UNKNOWN_VALUE: u64 = 42;

    #[test]
    fn radar_grid_shapes_are_closed_and_straight_is_the_native_default() {
        assert_eq!(
            ChartRadarGridShape::default(),
            ChartRadarGridShape::Straight
        );
        assert_eq!(
            ChartRadarGridShape::from_native(STRAIGHT_NATIVE_VALUE).unwrap(),
            ChartRadarGridShape::Straight
        );
        assert_eq!(
            ChartRadarGridShape::from_native(CURVED_NATIVE_VALUE).unwrap(),
            ChartRadarGridShape::Curved
        );
        assert!(ChartRadarGridShape::from_native(2).is_err());
    }

    #[test]
    fn radar_grid_shape_patch_preserves_other_fields_and_unknown_data() {
        let mut extension = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultradarradiusgridlinecurve: Some(false),
            tschchartinfodefaultshowborder: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
        let mut original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, UNKNOWN_VALUE).unwrap();

        let curved = patch_chart_radar_grid_shape(&original, ChartRadarGridShape::Curved).unwrap();
        assert_eq!(
            read_chart_radar_grid_shape(&curved).unwrap(),
            ChartRadarGridShape::Curved
        );
        assert!(has_field(&curved, UNKNOWN_OUTER_FIELD));
        let extension = generated_chart_style_extension(&curved).unwrap().unwrap();
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
        assert_eq!(
            tsch::generated::ChartStyleArchive::decode(extension)
                .unwrap()
                .tschchartinfodefaultshowborder,
            Some(true)
        );

        let straight =
            patch_chart_radar_grid_shape(&curved, ChartRadarGridShape::Straight).unwrap();
        assert_eq!(
            read_chart_radar_grid_shape(&straight).unwrap(),
            ChartRadarGridShape::Straight
        );
        assert_eq!(
            read_optional_native_shape(
                generated_chart_style_extension(&straight).unwrap().unwrap()
            )
            .unwrap(),
            None
        );
    }

    #[test]
    fn sparse_default_patch_is_a_no_op() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            patch_chart_radar_grid_shape(&original, ChartRadarGridShape::Straight).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_boolean_is_rejected() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, RADAR_RADIUS_GRIDLINE_CURVE_FIELD, 2).unwrap();
        let mut outer = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        assert!(read_chart_radar_grid_shape(&outer).is_err());
    }

    fn has_field(data: &[u8], field_number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number == field_number)
    }
}
