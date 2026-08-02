//! Lossless native 3D chart-depth storage and mutation.
//!
//! The Chart inspector exposes depth as a percentage. iWork stores the
//! corresponding depth ratio in the Z component of a chart-kind-specific
//! `TSCH.Chart3DVectorArchive`. X, Y, W, and unrelated protobuf fields are
//! retained byte-for-byte when an existing vector is updated.

use prost::Message;

use crate::charts::ChartKind;
use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{
    patch_fixed32_field, patch_length_delimited_field, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const AREA_3D_SCALE_FIELD: u32 = 5;
const BAR_3D_SCALE_FIELD: u32 = 6;
const COLUMN_3D_SCALE_FIELD: u32 = 7;
const LINE_3D_SCALE_FIELD: u32 = 9;
const PIE_3D_SCALE_FIELD: u32 = 10;
const VECTOR_Z_FIELD: u32 = 3;

const MINIMUM_DEPTH_PERCENT: f32 = 0.0;
const MAXIMUM_DEPTH_PERCENT: f32 = 100.0;
const DEFAULT_DEPTH_PERCENT: f32 = 11.020_468;
const BAR_MINIMUM_DEPTH_RATIO: f32 = 1.0 / 6.0;
const BAR_MAXIMUM_DEPTH_RATIO: f32 = 4.0;
const COLUMN_MINIMUM_DEPTH_RATIO: f32 = 1.0 / 6.0;
const COLUMN_MAXIMUM_DEPTH_RATIO: f32 = 2.5;
const LINEAR_MINIMUM_DEPTH_RATIO: f32 = 1.0 / 3.0;
const LINEAR_MAXIMUM_DEPTH_RATIO: f32 = 3.0;
const PIE_MINIMUM_DEPTH_RATIO: f32 = 0.25;
const PIE_MAXIMUM_DEPTH_RATIO: f32 = 4.0;
const DEPTH_PERCENT_EQUALITY_EPSILON: f32 = 0.001;

// A complete scale is required when a chart has not yet been opened in an
// iWork application. These are the native default scene dimensions; later
// layout passes may adjust them by a few ULPs.
const DEFAULT_SCALE_X: f32 = 1.5;
const DEFAULT_SCALE_Y: f32 = 2.2;
const DEFAULT_SCALE_W: f32 = f32::from_bits(0x3f22_b0c9);

/// Depth of a native 3D chart, in Chart-inspector percent.
#[derive(Debug, Clone, Copy)]
pub struct Chart3dDepth {
    percent: f32,
}

impl PartialEq for Chart3dDepth {
    fn eq(&self, other: &Self) -> bool {
        (self.percent - other.percent).abs() <= DEPTH_PERCENT_EQUALITY_EPSILON
    }
}

impl Chart3dDepth {
    /// Native depth used when no chart-specific scale is stored.
    pub const DEFAULT: Self = Self {
        percent: DEFAULT_DEPTH_PERCENT,
    };

    /// Construct a chart depth in the inspector's inclusive `0%..=100%` range.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite()
            || !(MINIMUM_DEPTH_PERCENT..=MAXIMUM_DEPTH_PERCENT).contains(&percent)
        {
            return Err(Error::InvalidFormat(format!(
                "chart 3D depth must be finite and within {MINIMUM_DEPTH_PERCENT}%..={MAXIMUM_DEPTH_PERCENT}%"
            )));
        }
        Ok(Self { percent })
    }

    /// Return the Chart-inspector percentage.
    pub const fn percent(self) -> f32 {
        self.percent
    }

    const fn depth_ratio(self, field: ScaleField) -> f32 {
        let (minimum, maximum) = field.depth_ratio_bounds();
        minimum + (maximum - minimum) * (self.percent / MAXIMUM_DEPTH_PERCENT)
    }

    fn from_depth_ratio(field: ScaleField, depth_ratio: f32) -> Result<Self> {
        if !depth_ratio.is_finite() {
            return Err(Error::InvalidFormat(
                "native chart 3D depth ratio must be finite".to_owned(),
            ));
        }
        let (minimum, maximum) = field.depth_ratio_bounds();
        let normalized = ((depth_ratio - minimum) / (maximum - minimum)).clamp(0.0, 1.0);
        Self::from_percent(normalized * MAXIMUM_DEPTH_PERCENT)
    }
}

impl Default for Chart3dDepth {
    fn default() -> Self {
        Self::DEFAULT
    }
}

#[derive(Debug, Clone, Copy)]
enum ScaleField {
    Area,
    Bar,
    Column,
    Line,
    Pie,
}

impl ScaleField {
    const fn for_kind(kind: ChartKind) -> Option<Self> {
        match kind {
            ChartKind::Area3d | ChartKind::StackedArea3d => Some(Self::Area),
            ChartKind::Bar3d | ChartKind::StackedBar3d => Some(Self::Bar),
            ChartKind::Column3d | ChartKind::StackedColumn3d => Some(Self::Column),
            ChartKind::Line3d => Some(Self::Line),
            ChartKind::Pie3d | ChartKind::Donut3d => Some(Self::Pie),
            _ => None,
        }
    }

    const fn number(self) -> u32 {
        match self {
            Self::Area => AREA_3D_SCALE_FIELD,
            Self::Bar => BAR_3D_SCALE_FIELD,
            Self::Column => COLUMN_3D_SCALE_FIELD,
            Self::Line => LINE_3D_SCALE_FIELD,
            Self::Pie => PIE_3D_SCALE_FIELD,
        }
    }

    const fn depth_ratio_bounds(self) -> (f32, f32) {
        match self {
            Self::Area | Self::Line => (LINEAR_MINIMUM_DEPTH_RATIO, LINEAR_MAXIMUM_DEPTH_RATIO),
            Self::Bar => (BAR_MINIMUM_DEPTH_RATIO, BAR_MAXIMUM_DEPTH_RATIO),
            Self::Column => (COLUMN_MINIMUM_DEPTH_RATIO, COLUMN_MAXIMUM_DEPTH_RATIO),
            Self::Pie => (PIE_MINIMUM_DEPTH_RATIO, PIE_MAXIMUM_DEPTH_RATIO),
        }
    }

    const fn ratio_denominator(self, vector: &tsch::Chart3DVectorArchive) -> f32 {
        match self {
            // TSCH3DChartType.depthRatioDimension is `isHorizontal`; the
            // vector maps dimension zero to X and dimension one to Y.
            Self::Bar => vector.y,
            Self::Area | Self::Column | Self::Line | Self::Pie => vector.x,
        }
    }

    fn vector(
        self,
        generated: &tsch::generated::ChartNonStyleArchive,
    ) -> Option<&tsch::Chart3DVectorArchive> {
        match self {
            Self::Area => generated.tschchartinfoarea3dscale.as_ref(),
            Self::Bar => generated.tschchartinfobar3dscale.as_ref(),
            Self::Column => generated.tschchartinfocolumn3dscale.as_ref(),
            Self::Line => generated.tschchartinfoline3dscale.as_ref(),
            Self::Pie => generated.tschchartinfopie3dscale.as_ref(),
        }
    }

    fn assign(
        self,
        generated: &mut tsch::generated::ChartNonStyleArchive,
        vector: tsch::Chart3DVectorArchive,
    ) {
        match self {
            Self::Area => generated.tschchartinfoarea3dscale = Some(vector),
            Self::Bar => generated.tschchartinfobar3dscale = Some(vector),
            Self::Column => generated.tschchartinfocolumn3dscale = Some(vector),
            Self::Line => generated.tschchartinfoline3dscale = Some(vector),
            Self::Pie => generated.tschchartinfopie3dscale = Some(vector),
        }
    }
}

/// Read one chart's effective native 3D depth.
pub(crate) fn chart_3d_depth(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
) -> Result<Chart3dDepth> {
    let field = ScaleField::for_kind(kind).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is not a 3D chart"
        ))
    })?;
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, |data| read_chart_3d_depth(data, field))
}

/// Set one chart's native 3D depth.
pub(crate) fn set_chart_3d_depth(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    depth: Chart3dDepth,
) -> Result<()> {
    let field = ScaleField::for_kind(kind).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is not a 3D chart"
        ))
    })?;
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, |data| read_chart_3d_depth(data, field))? == depth {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_3d_depth(data, field, depth))?;
    if slot.read(package, |data| read_chart_3d_depth(data, field))? != depth {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D depth update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_3d_depth(data: &[u8], field: ScaleField) -> Result<Chart3dDepth> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(Chart3dDepth::DEFAULT);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let Some(scale) = field.vector(&generated) else {
        return Ok(Chart3dDepth::DEFAULT);
    };
    if !scale.x.is_finite() || !scale.y.is_finite() || !scale.w.is_finite() {
        return Err(Error::InvalidFormat(
            "native chart 3D scale X/Y/W components must be finite".to_owned(),
        ));
    }
    let denominator = field.ratio_denominator(scale);
    if !denominator.is_finite() || denominator <= 0.0 {
        return Err(Error::InvalidFormat(
            "native chart 3D depth denominator must be finite and positive".to_owned(),
        ));
    }
    Chart3dDepth::from_depth_ratio(field, scale.z / denominator)
}

fn patch_chart_3d_depth(data: &[u8], field: ScaleField, depth: Chart3dDepth) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if depth == Chart3dDepth::DEFAULT {
            return Ok(data.to_vec());
        }
        let mut generated = tsch::generated::ChartNonStyleArchive::default();
        field.assign(&mut generated, native_scale(field, depth));
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_3d_depth(&patched, field, depth)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let extension = if let Some(scale) = field.vector(&generated) {
        let denominator = field.ratio_denominator(scale);
        if !denominator.is_finite() || denominator <= 0.0 {
            return Err(Error::InvalidFormat(
                "native chart 3D depth denominator must be finite and positive".to_owned(),
            ));
        }
        let native_depth = depth.depth_ratio(field) * denominator;
        transform_length_delimited_field(extension, field.number(), |vector| {
            patch_fixed32_field(vector, VECTOR_Z_FIELD, true, Some(native_depth.to_bits()))
        })?
    } else if depth == Chart3dDepth::DEFAULT {
        extension.to_vec()
    } else {
        let encoded = native_scale(field, depth).encode_to_vec();
        patch_length_delimited_field(extension, field.number(), false, Some(encoded.as_slice()))?
    };
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_3d_depth(&patched, field, depth)?;
    Ok(patched)
}

const fn native_scale(field: ScaleField, depth: Chart3dDepth) -> tsch::Chart3DVectorArchive {
    let mut scale = tsch::Chart3DVectorArchive {
        x: DEFAULT_SCALE_X,
        y: DEFAULT_SCALE_Y,
        z: 0.0,
        w: DEFAULT_SCALE_W,
    };
    scale.z = depth.depth_ratio(field) * field.ratio_denominator(&scale);
    scale
}

fn validate_patched_chart_3d_depth(
    data: &[u8],
    field: ScaleField,
    expected: Chart3dDepth,
) -> Result<()> {
    if read_chart_3d_depth(data, field)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D depth wire patch failed validation".to_owned(),
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
    const UNKNOWN_VECTOR_FIELD: u32 = 4_098;

    #[test]
    fn depth_is_bounded_and_has_the_native_default() {
        assert_eq!(Chart3dDepth::default(), Chart3dDepth::DEFAULT);
        assert_eq!(Chart3dDepth::DEFAULT.percent(), DEFAULT_DEPTH_PERCENT);
        assert!(Chart3dDepth::from_percent(f32::NAN).is_err());
        assert!(Chart3dDepth::from_percent(f32::INFINITY).is_err());
        assert!(Chart3dDepth::from_percent(-0.001).is_err());
        assert!(Chart3dDepth::from_percent(100.001).is_err());
        assert_eq!(
            Chart3dDepth::from_percent(0.0)
                .unwrap()
                .depth_ratio(ScaleField::Column),
            COLUMN_MINIMUM_DEPTH_RATIO
        );
        assert_eq!(
            Chart3dDepth::from_percent(100.0)
                .unwrap()
                .depth_ratio(ScaleField::Column),
            COLUMN_MAXIMUM_DEPTH_RATIO
        );
    }

    #[test]
    fn depth_patch_preserves_unknown_and_other_vector_fields() {
        let original_depth = Chart3dDepth::from_percent(25.0).unwrap();
        let mut vector = native_scale(ScaleField::Column, original_depth).encode_to_vec();
        append_varint_field(&mut vector, UNKNOWN_VECTOR_FIELD, 42).unwrap();
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfocolumn3dscale: Some(native_scale(ScaleField::Column, original_depth)),
            tschchartinfodefaultshowlegend: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        extension =
            patch_length_delimited_field(&extension, COLUMN_3D_SCALE_FIELD, true, Some(&vector))
                .unwrap();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 43).unwrap();
        let mut original = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, 44).unwrap();

        let expected = Chart3dDepth::from_percent(75.0).unwrap();
        let patched = patch_chart_3d_depth(&original, ScaleField::Column, expected).unwrap();
        assert_eq!(
            read_chart_3d_depth(&patched, ScaleField::Column).unwrap(),
            expected
        );
        assert!(has_field(&patched, UNKNOWN_OUTER_FIELD));
        let extension = generated_chart_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
        let vector = find_field_bytes(extension, COLUMN_3D_SCALE_FIELD);
        assert!(has_field(vector, UNKNOWN_VECTOR_FIELD));
        let decoded = tsch::Chart3DVectorArchive::decode(vector).unwrap();
        assert_eq!(decoded.x, DEFAULT_SCALE_X);
        assert_eq!(decoded.y, DEFAULT_SCALE_Y);
        assert_eq!(decoded.w, DEFAULT_SCALE_W);
    }

    #[test]
    fn every_native_3d_kind_selects_its_scale_slot() {
        let cases = [
            (ChartKind::Area3d, AREA_3D_SCALE_FIELD),
            (ChartKind::StackedArea3d, AREA_3D_SCALE_FIELD),
            (ChartKind::Bar3d, BAR_3D_SCALE_FIELD),
            (ChartKind::StackedBar3d, BAR_3D_SCALE_FIELD),
            (ChartKind::Column3d, COLUMN_3D_SCALE_FIELD),
            (ChartKind::StackedColumn3d, COLUMN_3D_SCALE_FIELD),
            (ChartKind::Line3d, LINE_3D_SCALE_FIELD),
            (ChartKind::Pie3d, PIE_3D_SCALE_FIELD),
            (ChartKind::Donut3d, PIE_3D_SCALE_FIELD),
        ];
        for (kind, expected) in cases {
            assert_eq!(ScaleField::for_kind(kind).unwrap().number(), expected);
        }
        assert!(ScaleField::for_kind(ChartKind::Column2d).is_none());
    }

    #[test]
    fn every_scale_family_round_trips_inspector_endpoints() {
        for field in [
            ScaleField::Area,
            ScaleField::Bar,
            ScaleField::Column,
            ScaleField::Line,
            ScaleField::Pie,
        ] {
            for percent in [0.0, 25.0, 50.0, 75.0, 100.0] {
                let expected = Chart3dDepth::from_percent(percent).unwrap();
                let scale = native_scale(field, expected);
                let actual = Chart3dDepth::from_depth_ratio(
                    field,
                    scale.z / field.ratio_denominator(&scale),
                )
                .unwrap();
                assert_eq!(actual, expected, "{field:?} at {percent}%");
            }
        }
    }

    fn has_field(data: &[u8], number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number == number)
    }

    fn find_field_bytes(data: &[u8], number: u32) -> &[u8] {
        let fields = parse_wire_fields(data).unwrap();
        let field = fields.iter().find(|field| field.number == number).unwrap();
        &data[field.payload_start..field.end]
    }
}
