//! Lossless native chart-axis gridline stroke storage and mutation.

use prost::Message;

use crate::charts::Axis;
use crate::charts::axis_style::{
    GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD, axis_style_slot, generated_axis_style_extension,
};
use crate::protobuf::{tsch, tsd};
use crate::shapes::{ShapeStroke, empty_stroke_archive, stroke_from_native, stroke_to_native};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const CATEGORY_MAJOR_GRIDLINE_STROKE_FIELD: u32 = 16;
const VALUE_MAJOR_GRIDLINE_STROKE_FIELD: u32 = 17;
const CATEGORY_MINOR_GRIDLINE_STROKE_FIELD: u32 = 22;
const VALUE_MINOR_GRIDLINE_STROKE_FIELD: u32 = 23;

/// Which gridline family on one chart axis is being styled.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartAxisGridline {
    Major,
    Minor,
}

impl ChartAxisGridline {
    const fn field(self, axis: Axis) -> u32 {
        match (self, axis) {
            (Self::Major, Axis::Category) => CATEGORY_MAJOR_GRIDLINE_STROKE_FIELD,
            (Self::Major, Axis::Value) => VALUE_MAJOR_GRIDLINE_STROKE_FIELD,
            (Self::Minor, Axis::Category) => CATEGORY_MINOR_GRIDLINE_STROKE_FIELD,
            (Self::Minor, Axis::Value) => VALUE_MINOR_GRIDLINE_STROKE_FIELD,
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::Major => "major",
            Self::Minor => "minor",
        }
    }

    fn native(
        self,
        generated: &tsch::generated::ChartAxisStyleArchive,
        axis: Axis,
    ) -> Option<&tsd::StrokeArchive> {
        match (self, axis) {
            (Self::Major, Axis::Category) => {
                generated.tschchartaxiscategorymajorgridlinestroke.as_ref()
            },
            (Self::Major, Axis::Value) => generated.tschchartaxisvaluemajorgridlinestroke.as_ref(),
            (Self::Minor, Axis::Category) => {
                generated.tschchartaxiscategoryminorgridlinestroke.as_ref()
            },
            (Self::Minor, Axis::Value) => generated.tschchartaxisvalueminorgridlinestroke.as_ref(),
        }
    }

    fn set_native(
        self,
        generated: &mut tsch::generated::ChartAxisStyleArchive,
        axis: Axis,
        stroke: Option<tsd::StrokeArchive>,
    ) {
        match (self, axis) {
            (Self::Major, Axis::Category) => {
                generated.tschchartaxiscategorymajorgridlinestroke = stroke
            },
            (Self::Major, Axis::Value) => generated.tschchartaxisvaluemajorgridlinestroke = stroke,
            (Self::Minor, Axis::Category) => {
                generated.tschchartaxiscategoryminorgridlinestroke = stroke
            },
            (Self::Minor, Axis::Value) => generated.tschchartaxisvalueminorgridlinestroke = stroke,
        }
    }
}

/// Exact native stroke state for one chart-axis gridline family.
///
/// `Inherited` preserves the absence of an override. `NoStroke` mirrors the
/// inspector's “None” choice. Visibility remains independently editable
/// through the existing gridline visibility APIs.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartAxisGridlineStroke {
    #[default]
    Inherited,
    NoStroke,
    Stroke(ShapeStroke),
}

/// Read the exact native gridline stroke state for one chart axis.
pub(crate) fn chart_axis_gridline_stroke(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
    gridline: ChartAxisGridline,
) -> Result<ChartAxisGridlineStroke> {
    axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?
    .read(package, |data| {
        read_axis_gridline_stroke(data, axis, gridline)
    })
}

/// Set one chart axis' gridline stroke without changing its visibility.
pub(crate) fn set_chart_axis_gridline_stroke(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    axis: Axis,
    gridline: ChartAxisGridline,
    stroke: ChartAxisGridlineStroke,
) -> Result<()> {
    let slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        axis,
    )?;
    if slot.read(package, |data| {
        read_axis_gridline_stroke(data, axis, gridline)
    })? == stroke
    {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_axis_gridline_stroke(data, axis, gridline, stroke)
    })?;
    if slot.read(package, |data| {
        read_axis_gridline_stroke(data, axis, gridline)
    })? != stroke
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} {}-axis {} gridline stroke update failed validation",
            axis.as_str(),
            gridline.label(),
        )));
    }
    Ok(())
}

fn read_axis_gridline_stroke(
    data: &[u8],
    axis: Axis,
    gridline: ChartAxisGridline,
) -> Result<ChartAxisGridlineStroke> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok(ChartAxisGridlineStroke::Inherited);
    };
    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let Some(native) = gridline.native(&generated, axis) else {
        return Ok(ChartAxisGridlineStroke::Inherited);
    };
    Ok(match stroke_from_native(native)? {
        Some(stroke) => ChartAxisGridlineStroke::Stroke(stroke),
        None => ChartAxisGridlineStroke::NoStroke,
    })
}

fn patch_axis_gridline_stroke(
    data: &[u8],
    axis: Axis,
    gridline: ChartAxisGridline,
    stroke: ChartAxisGridlineStroke,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        if stroke == ChartAxisGridlineStroke::Inherited {
            return Ok(data.to_vec());
        }
        let mut generated = tsch::generated::ChartAxisStyleArchive::default();
        gridline.set_native(&mut generated, axis, native_stroke(stroke));
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_axis_gridline_stroke(&patched, axis, gridline, stroke)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let field_present = gridline.native(&generated, axis).is_some();
    let encoded = native_stroke(stroke).map(|stroke| stroke.encode_to_vec());
    let extension = patch_length_delimited_field(
        extension,
        gridline.field(axis),
        field_present,
        encoded.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_axis_gridline_stroke(&patched, axis, gridline, stroke)?;
    Ok(patched)
}

fn native_stroke(stroke: ChartAxisGridlineStroke) -> Option<tsd::StrokeArchive> {
    match stroke {
        ChartAxisGridlineStroke::Inherited => None,
        ChartAxisGridlineStroke::NoStroke => Some(empty_stroke_archive()),
        ChartAxisGridlineStroke::Stroke(stroke) => Some(stroke_to_native(stroke)),
    }
}

fn validate_patched_axis_gridline_stroke(
    data: &[u8],
    axis: Axis,
    gridline: ChartAxisGridline,
    expected: ChartAxisGridlineStroke,
) -> Result<()> {
    if read_axis_gridline_stroke(data, axis, gridline)? != expected {
        return Err(Error::InvalidFormat(format!(
            "{}-axis {} gridline stroke wire patch failed validation",
            axis.as_str(),
            gridline.label(),
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::{RgbColorSpace, RgbaColor, StrokePattern, StrokeWidth};
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn gridline_strokes_are_exact_and_preserve_unknown_fields() {
        let original_stroke = test_stroke(StrokePattern::Solid, 1.0);
        let replacement = test_stroke(StrokePattern::MediumDash, 3.0);
        let mut extension = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategorymajorgridlinestroke: Some(stroke_to_native(original_stroke)),
            tschchartaxiscategoryshowmajorgridlines: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let changed = patch_axis_gridline_stroke(
            &original,
            Axis::Category,
            ChartAxisGridline::Major,
            ChartAxisGridlineStroke::Stroke(replacement),
        )
        .unwrap();
        assert_eq!(
            read_axis_gridline_stroke(&changed, Axis::Category, ChartAxisGridline::Major,).unwrap(),
            ChartAxisGridlineStroke::Stroke(replacement)
        );
        assert_eq!(
            raw_field(&changed, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_axis_style_extension(&changed).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(&extension, UNMAPPED_GENERATED_FIELD)
        );

        let empty = patch_axis_gridline_stroke(
            &changed,
            Axis::Category,
            ChartAxisGridline::Major,
            ChartAxisGridlineStroke::NoStroke,
        )
        .unwrap();
        assert_eq!(
            read_axis_gridline_stroke(&empty, Axis::Category, ChartAxisGridline::Major,).unwrap(),
            ChartAxisGridlineStroke::NoStroke
        );

        let inherited = patch_axis_gridline_stroke(
            &empty,
            Axis::Category,
            ChartAxisGridline::Major,
            ChartAxisGridlineStroke::Inherited,
        )
        .unwrap();
        assert_eq!(
            read_axis_gridline_stroke(&inherited, Axis::Category, ChartAxisGridline::Major,)
                .unwrap(),
            ChartAxisGridlineStroke::Inherited
        );
    }

    fn test_stroke(pattern: StrokePattern, width: f32) -> ShapeStroke {
        ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
            StrokeWidth::new(width).unwrap(),
            pattern,
        )
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
