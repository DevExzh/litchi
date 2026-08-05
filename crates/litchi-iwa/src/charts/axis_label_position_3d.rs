//! Lossless native 3D value-axis label-position storage and mutation.
//!
//! Axis-bearing 3D charts expose Automatic plus two orientation-dependent
//! locations for their primary value-axis labels. The inspector names them
//! Left/Right for vertical charts and Top/Bottom for horizontal bar charts.

use prost::Message;

use litchi_iwa_common::chart::axis::label_position_3d::LabelPosition3d;

use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::charts::{Axis, Kind};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisdefault3dlabelposition` in the generated axis non-style.
const VALUE_AXIS_3D_LABEL_POSITION_FIELD: u32 = 1;
/// Read one chart's effective 3D value-axis label position.
pub(crate) fn chart_3d_value_axis_label_position(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
) -> Result<LabelPosition3d> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
    )?
    .read(package, read_3d_value_axis_label_position)
}

/// Set one chart's 3D value-axis label position.
pub(crate) fn set_chart_3d_value_axis_label_position(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    position: LabelPosition3d,
) -> Result<()> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Value,
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

fn require_supported_kind(kind: Kind, drawable_object_id: u64, drawable_label: &str) -> Result<()> {
    if !kind.supports_3d_value_axis_label_position() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} kind {kind:?} has no 3D value-axis label position"
        )));
    }
    Ok(())
}

fn read_3d_value_axis_label_position(data: &[u8]) -> Result<LabelPosition3d> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(LabelPosition3d::default());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    Ok(generated
        .tschchartaxisdefault3dlabelposition
        .map(LabelPosition3d::from_native)
        .unwrap_or_default())
}

fn patch_3d_value_axis_label_position(data: &[u8], position: LabelPosition3d) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if position == LabelPosition3d::default() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisdefault3dlabelposition: Some(position.native_value()),
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
    let native = (present || position != LabelPosition3d::default())
        .then_some(position.native_value() as u64);
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
    expected: LabelPosition3d,
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
            (
                LabelPosition3d::Automatic.native_value(),
                LabelPosition3d::Automatic,
            ),
            (
                LabelPosition3d::Leading.native_value(),
                LabelPosition3d::Leading,
            ),
            (
                LabelPosition3d::Trailing.native_value(),
                LabelPosition3d::Trailing,
            ),
            (
                FUTURE_NATIVE_VALUE,
                LabelPosition3d::Unsupported(FUTURE_NATIVE_VALUE),
            ),
        ];
        for (native, position) in cases {
            assert_eq!(LabelPosition3d::from_native(native), position);
            assert_eq!(position.native_value(), native);
        }
        assert_eq!(LabelPosition3d::default(), LabelPosition3d::Automatic);
    }

    #[test]
    fn capability_matches_all_axis_bearing_3d_kinds() {
        for kind in [
            Kind::Column3d,
            Kind::Bar3d,
            Kind::Line3d,
            Kind::Area3d,
            Kind::StackedColumn3d,
            Kind::StackedBar3d,
            Kind::StackedArea3d,
        ] {
            assert!(kind.supports_3d_value_axis_label_position(), "{kind:?}");
        }
        for kind in [Kind::Column2d, Kind::Pie3d, Kind::Donut3d] {
            assert!(!kind.supports_3d_value_axis_label_position(), "{kind:?}");
        }
    }

    #[test]
    fn position_patch_preserves_neighboring_and_unknown_fields() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisdefault3dlabelposition: Some(
                    LabelPosition3d::Automatic.native_value(),
                ),
                tschchartaxisvaluescale: Some(2),
                tschchartaxisvaluenumberofmajorgridlines: Some(7),
                ..Default::default()
            });
        let patched =
            patch_3d_value_axis_label_position(&original, LabelPosition3d::Trailing).unwrap();
        assert_eq!(
            read_3d_value_axis_label_position(&patched).unwrap(),
            LabelPosition3d::Trailing
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
            patch_3d_value_axis_label_position(&patched, LabelPosition3d::Automatic).unwrap();
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
            LabelPosition3d::Automatic
        );
        assert_eq!(
            patch_3d_value_axis_label_position(&original, LabelPosition3d::Automatic).unwrap(),
            original
        );
        let leading =
            patch_3d_value_axis_label_position(&original, LabelPosition3d::Leading).unwrap();
        assert_eq!(
            read_3d_value_axis_label_position(&leading).unwrap(),
            LabelPosition3d::Leading
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
            .any(|field| field.number() == number)
    }
}
