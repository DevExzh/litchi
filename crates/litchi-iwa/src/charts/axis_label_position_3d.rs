//! Lossless native 3D value-axis label-position storage and mutation.
//!
//! Axis-bearing 3D charts expose Automatic plus two orientation-dependent
//! locations for their primary value-axis labels. The inspector names them
//! Left/Right for vertical charts and Top/Bottom for horizontal bar charts.

use prost::Message;

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::charts::{ChartAxis, ChartKind};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisdefault3dlabelposition` in the generated axis non-style.
const VALUE_AXIS_3D_LABEL_POSITION_FIELD: u32 = 1;
const AUTOMATIC_NATIVE_VALUE: i32 = 1;
const LEADING_NATIVE_VALUE: i32 = 2;
const TRAILING_NATIVE_VALUE: i32 = 3;

/// Position of primary value-axis labels in a native 3D chart.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Chart3dAxisLabelPosition {
    /// Let iWork choose the side from chart orientation.
    #[default]
    Automatic,
    /// Show labels at Left, or Top for a horizontal bar chart.
    Leading,
    /// Show labels at Right, or Bottom for a horizontal bar chart.
    Trailing,
    /// Preserve an unrecognized future native value.
    Unsupported(i32),
}

impl Chart3dAxisLabelPosition {
    /// Decode the integer stored by the iWork protobuf schema.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            AUTOMATIC_NATIVE_VALUE => Self::Automatic,
            LEADING_NATIVE_VALUE => Self::Leading,
            TRAILING_NATIVE_VALUE => Self::Trailing,
            value => Self::Unsupported(value),
        }
    }

    /// Return the integer used by the iWork protobuf schema.
    pub const fn into_raw(self) -> i32 {
        match self {
            Self::Automatic => AUTOMATIC_NATIVE_VALUE,
            Self::Leading => LEADING_NATIVE_VALUE,
            Self::Trailing => TRAILING_NATIVE_VALUE,
            Self::Unsupported(value) => value,
        }
    }
}

/// Read one chart's effective 3D value-axis label position.
pub(crate) fn chart_3d_value_axis_label_position(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
) -> Result<Chart3dAxisLabelPosition> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?
    .read(package, read_3d_value_axis_label_position)
}

/// Set one chart's 3D value-axis label position.
pub(crate) fn set_chart_3d_value_axis_label_position(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    position: Chart3dAxisLabelPosition,
) -> Result<()> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?;
    if slot.read(package, read_3d_value_axis_label_position)? == position {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_3d_value_axis_label_position(data, position)
    })?;
    if slot.read(package, read_3d_value_axis_label_position)? != position {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D value-axis label-position update failed validation"
        )));
    }
    Ok(())
}

fn require_supported_kind(
    kind: ChartKind,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<()> {
    if !kind.supports_3d_value_axis_label_position() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} kind {kind:?} has no 3D value-axis label position"
        )));
    }
    Ok(())
}

fn read_3d_value_axis_label_position(data: &[u8]) -> Result<Chart3dAxisLabelPosition> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(Chart3dAxisLabelPosition::default());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(generated
        .tschchartaxisdefault3dlabelposition
        .map(Chart3dAxisLabelPosition::from_raw)
        .unwrap_or_default())
}

fn patch_3d_value_axis_label_position(
    data: &[u8],
    position: Chart3dAxisLabelPosition,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if position == Chart3dAxisLabelPosition::default() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisdefault3dlabelposition: Some(position.into_raw()),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_3d_value_axis_label_position(&patched, position)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let present = generated.tschchartaxisdefault3dlabelposition.is_some();
    let native = (present || position != Chart3dAxisLabelPosition::default())
        .then_some(position.into_raw() as u64);
    let extension = patch_varint_field(
        extension,
        VALUE_AXIS_3D_LABEL_POSITION_FIELD,
        present,
        native,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_3d_value_axis_label_position(&patched, position)?;
    Ok(patched)
}

fn validate_patched_3d_value_axis_label_position(
    data: &[u8],
    expected: Chart3dAxisLabelPosition,
) -> Result<()> {
    if read_3d_value_axis_label_position(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D value-axis label-position wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;
    const FUTURE_NATIVE_VALUE: i32 = 9_001;

    #[test]
    fn native_values_are_typed_defaulted_and_forward_compatible() {
        let cases = [
            (AUTOMATIC_NATIVE_VALUE, Chart3dAxisLabelPosition::Automatic),
            (LEADING_NATIVE_VALUE, Chart3dAxisLabelPosition::Leading),
            (TRAILING_NATIVE_VALUE, Chart3dAxisLabelPosition::Trailing),
            (
                FUTURE_NATIVE_VALUE,
                Chart3dAxisLabelPosition::Unsupported(FUTURE_NATIVE_VALUE),
            ),
        ];
        for (native, position) in cases {
            assert_eq!(Chart3dAxisLabelPosition::from_raw(native), position);
            assert_eq!(position.into_raw(), native);
        }
        assert_eq!(
            Chart3dAxisLabelPosition::default(),
            Chart3dAxisLabelPosition::Automatic
        );
    }

    #[test]
    fn capability_matches_all_axis_bearing_3d_kinds() {
        for kind in [
            ChartKind::Column3d,
            ChartKind::Bar3d,
            ChartKind::Line3d,
            ChartKind::Area3d,
            ChartKind::StackedColumn3d,
            ChartKind::StackedBar3d,
            ChartKind::StackedArea3d,
        ] {
            assert!(kind.supports_3d_value_axis_label_position(), "{kind:?}");
        }
        for kind in [ChartKind::Column2d, ChartKind::Pie3d, ChartKind::Donut3d] {
            assert!(!kind.supports_3d_value_axis_label_position(), "{kind:?}");
        }
    }

    #[test]
    fn position_patch_preserves_neighboring_and_unknown_fields() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisdefault3dlabelposition: Some(AUTOMATIC_NATIVE_VALUE),
                tschchartaxisvaluescale: Some(2),
                tschchartaxisvaluenumberofmajorgridlines: Some(7),
                ..Default::default()
            });
        let patched =
            patch_3d_value_axis_label_position(&original, Chart3dAxisLabelPosition::Trailing)
                .unwrap();
        assert_eq!(
            read_3d_value_axis_label_position(&patched).unwrap(),
            Chart3dAxisLabelPosition::Trailing
        );
        let extension = generated_axis_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension).unwrap();
        assert_eq!(generated.tschchartaxisvaluescale, Some(2));
        assert_eq!(generated.tschchartaxisvaluenumberofmajorgridlines, Some(7));
        assert!(has_field(&patched, UNKNOWN_OUTER_FIELD));
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));

        let restored =
            patch_3d_value_axis_label_position(&patched, Chart3dAxisLabelPosition::Automatic)
                .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn absent_default_is_an_exact_no_op_and_nondefault_creates_storage() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            read_3d_value_axis_label_position(&original).unwrap(),
            Chart3dAxisLabelPosition::Automatic
        );
        assert_eq!(
            patch_3d_value_axis_label_position(&original, Chart3dAxisLabelPosition::Automatic)
                .unwrap(),
            original
        );
        let leading =
            patch_3d_value_axis_label_position(&original, Chart3dAxisLabelPosition::Leading)
                .unwrap();
        assert_eq!(
            read_3d_value_axis_label_position(&leading).unwrap(),
            Chart3dAxisLabelPosition::Leading
        );
    }

    fn axis_non_style_with_unknown_fields(
        generated: tsch::generated::ChartAxisNonStyleArchive,
    ) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 43).unwrap();
        let mut data = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut data, UNKNOWN_OUTER_FIELD, 44).unwrap();
        data
    }

    fn has_field(data: &[u8], number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number == number)
    }
}
