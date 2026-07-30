//! Lossless native 3D chart-lighting style storage and mutation.
//!
//! The Chart inspector exposes seven fixed lighting packages. Non-default
//! packages are stored in a chart-family-specific field of the generated chart
//! style extension. Unrelated and unknown protobuf fields are retained.

mod presets;

use prost::Message;

use self::presets::native_lighting_package;
use crate::charts::ChartKind;
use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

const AREA_3D_LIGHTING_FIELD: u32 = 2;
const BAR_3D_LIGHTING_FIELD: u32 = 3;
const COLUMN_3D_LIGHTING_FIELD: u32 = 4;
const LINE_3D_LIGHTING_FIELD: u32 = 6;
const PIE_3D_LIGHTING_FIELD: u32 = 7;

const DEFAULT_NATIVE_NAME: &str = "Default";
const SOFT_LIGHT_NATIVE_NAME: &str = "Soft Light";
const SOFT_FILL_NATIVE_NAME: &str = "Soft Fill";
const MEDIUM_CENTER_NATIVE_NAME: &str = "Medium Center";
const MEDIUM_RIGHT_NATIVE_NAME: &str = "Medium Right";
const MEDIUM_LEFT_NATIVE_NAME: &str = "Medium Left";
const GLOSSY_NATIVE_NAME: &str = "Glossy";

/// One native choice from the Chart inspector's Lighting Style menu.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Chart3dLightingStyle {
    #[default]
    Default,
    SoftLight,
    SoftFill,
    MediumCenter,
    MediumRight,
    MediumLeft,
    Glossy,
}

impl Chart3dLightingStyle {
    /// Return the label shown by the native Chart inspector.
    pub const fn inspector_label(self) -> &'static str {
        self.native_name()
    }

    pub(super) const fn native_name(self) -> &'static str {
        match self {
            Self::Default => DEFAULT_NATIVE_NAME,
            Self::SoftLight => SOFT_LIGHT_NATIVE_NAME,
            Self::SoftFill => SOFT_FILL_NATIVE_NAME,
            Self::MediumCenter => MEDIUM_CENTER_NATIVE_NAME,
            Self::MediumRight => MEDIUM_RIGHT_NATIVE_NAME,
            Self::MediumLeft => MEDIUM_LEFT_NATIVE_NAME,
            Self::Glossy => GLOSSY_NATIVE_NAME,
        }
    }

    fn from_native_name(name: &str) -> Result<Self> {
        match name {
            DEFAULT_NATIVE_NAME => Ok(Self::Default),
            SOFT_LIGHT_NATIVE_NAME => Ok(Self::SoftLight),
            SOFT_FILL_NATIVE_NAME => Ok(Self::SoftFill),
            MEDIUM_CENTER_NATIVE_NAME => Ok(Self::MediumCenter),
            MEDIUM_RIGHT_NATIVE_NAME => Ok(Self::MediumRight),
            MEDIUM_LEFT_NATIVE_NAME => Ok(Self::MediumLeft),
            GLOSSY_NATIVE_NAME => Ok(Self::Glossy),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported native chart 3D lighting package {name:?}"
            ))),
        }
    }
}

impl std::fmt::Display for Chart3dLightingStyle {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.inspector_label())
    }
}

#[derive(Debug, Clone, Copy)]
enum LightingField {
    Area,
    Bar,
    Column,
    Line,
    Pie,
}

impl LightingField {
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
            Self::Area => AREA_3D_LIGHTING_FIELD,
            Self::Bar => BAR_3D_LIGHTING_FIELD,
            Self::Column => COLUMN_3D_LIGHTING_FIELD,
            Self::Line => LINE_3D_LIGHTING_FIELD,
            Self::Pie => PIE_3D_LIGHTING_FIELD,
        }
    }

    fn package(
        self,
        generated: &tsch::generated::ChartStyleArchive,
    ) -> Option<&tsch::Chart3DLightingPackageArchive> {
        match self {
            Self::Area => generated.tschchartinfoarea3dlightingpackage.as_ref(),
            Self::Bar => generated.tschchartinfobar3dlightingpackage.as_ref(),
            Self::Column => generated.tschchartinfocolumn3dlightingpackage.as_ref(),
            Self::Line => generated.tschchartinfoline3dlightingpackage.as_ref(),
            Self::Pie => generated.tschchartinfopie3dlightingpackage.as_ref(),
        }
    }

    fn assign(
        self,
        generated: &mut tsch::generated::ChartStyleArchive,
        package: tsch::Chart3DLightingPackageArchive,
    ) {
        match self {
            Self::Area => generated.tschchartinfoarea3dlightingpackage = Some(package),
            Self::Bar => generated.tschchartinfobar3dlightingpackage = Some(package),
            Self::Column => generated.tschchartinfocolumn3dlightingpackage = Some(package),
            Self::Line => generated.tschchartinfoline3dlightingpackage = Some(package),
            Self::Pie => generated.tschchartinfopie3dlightingpackage = Some(package),
        }
    }
}

/// Read one chart's effective native 3D lighting style.
pub(crate) fn chart_3d_lighting_style(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
) -> Result<Chart3dLightingStyle> {
    let field = require_lighting_field(kind, drawable_object_id, drawable_label)?;
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, |data| read_chart_3d_lighting_style(data, field))
}

/// Set one chart's native 3D lighting style.
pub(crate) fn set_chart_3d_lighting_style(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: ChartKind,
    style: Chart3dLightingStyle,
) -> Result<()> {
    let field = require_lighting_field(kind, drawable_object_id, drawable_label)?;
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, |data| read_chart_3d_lighting_style(data, field))? == style {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_3d_lighting_style(data, field, style)
    })?;
    if slot.read(package, |data| read_chart_3d_lighting_style(data, field))? != style {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D lighting-style update failed validation"
        )));
    }
    Ok(())
}

fn require_lighting_field(
    kind: ChartKind,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<LightingField> {
    LightingField::for_kind(kind).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} kind {kind:?} has no 3D lighting style"
        ))
    })
}

fn read_chart_3d_lighting_style(data: &[u8], field: LightingField) -> Result<Chart3dLightingStyle> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(Chart3dLightingStyle::Default);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let Some(package) = field.package(&generated) else {
        return Ok(Chart3dLightingStyle::Default);
    };
    validate_native_package(package)?;
    Chart3dLightingStyle::from_native_name(&package.name)
}

fn patch_chart_3d_lighting_style(
    data: &[u8],
    field: LightingField,
    style: Chart3dLightingStyle,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        let Some(package) = native_lighting_package(style) else {
            return Ok(data.to_vec());
        };
        let mut generated = tsch::generated::ChartStyleArchive::default();
        field.assign(&mut generated, package);
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_3d_lighting_style(&patched, field, style)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let present = field.package(&generated).is_some();
    let native = native_lighting_package(style).map(|package| package.encode_to_vec());
    let extension =
        patch_length_delimited_field(extension, field.number(), present, native.as_deref())?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_3d_lighting_style(&patched, field, style)?;
    Ok(patched)
}

fn validate_native_package(package: &tsch::Chart3DLightingPackageArchive) -> Result<()> {
    if package.name.is_empty() || package.lights.is_empty() {
        return Err(Error::InvalidFormat(
            "native chart 3D lighting package must have a name and at least one light".to_owned(),
        ));
    }
    for light in &package.lights {
        if light.name.is_empty()
            || !light.intensity.is_finite()
            || light.intensity < 0.0
            || !valid_vector(light.ambient_color)
            || !valid_vector(light.diffuse_color)
            || !valid_vector(light.specular_color)
            || !valid_vector(light.attenuation)
            || light.coordinate_space > 1
        {
            return Err(Error::InvalidFormat(
                "native chart 3D lighting package contains an invalid light".to_owned(),
            ));
        }
        let source_count = usize::from(light.point_light.is_some())
            + usize::from(light.directional_light.is_some())
            + usize::from(light.spot_light.is_some());
        if source_count != 1
            || light
                .point_light
                .is_some_and(|source| !valid_vector(source.position))
            || light
                .directional_light
                .is_some_and(|source| !valid_vector(source.direction))
            || light.spot_light.is_some_and(|source| {
                !valid_vector(source.position)
                    || !valid_vector(source.direction)
                    || !source.cutoff.is_finite()
                    || !source.dropoff.is_finite()
            })
        {
            return Err(Error::InvalidFormat(
                "native chart 3D lighting package contains an invalid light source".to_owned(),
            ));
        }
    }
    Ok(())
}

const fn valid_vector(vector: tsch::Chart3DVectorArchive) -> bool {
    vector.x.is_finite() && vector.y.is_finite() && vector.z.is_finite() && vector.w.is_finite()
}

fn validate_patched_chart_3d_lighting_style(
    data: &[u8],
    field: LightingField,
    expected: Chart3dLightingStyle,
) -> Result<()> {
    if read_chart_3d_lighting_style(data, field)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D lighting-style wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_varint_field, parse_wire_fields};

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_EXTENSION_FIELD: u32 = 4_097;

    #[test]
    fn all_native_presets_round_trip_in_every_family_slot() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        for field in [
            LightingField::Area,
            LightingField::Bar,
            LightingField::Column,
            LightingField::Line,
            LightingField::Pie,
        ] {
            let mut current = original.clone();
            for style in [
                Chart3dLightingStyle::SoftLight,
                Chart3dLightingStyle::SoftFill,
                Chart3dLightingStyle::MediumCenter,
                Chart3dLightingStyle::MediumRight,
                Chart3dLightingStyle::MediumLeft,
                Chart3dLightingStyle::Glossy,
                Chart3dLightingStyle::Default,
            ] {
                current = patch_chart_3d_lighting_style(&current, field, style).unwrap();
                assert_eq!(
                    read_chart_3d_lighting_style(&current, field).unwrap(),
                    style
                );
            }
        }
    }

    #[test]
    fn every_native_3d_kind_selects_its_lighting_family() {
        let cases = [
            (ChartKind::Area3d, AREA_3D_LIGHTING_FIELD),
            (ChartKind::StackedArea3d, AREA_3D_LIGHTING_FIELD),
            (ChartKind::Bar3d, BAR_3D_LIGHTING_FIELD),
            (ChartKind::StackedBar3d, BAR_3D_LIGHTING_FIELD),
            (ChartKind::Column3d, COLUMN_3D_LIGHTING_FIELD),
            (ChartKind::StackedColumn3d, COLUMN_3D_LIGHTING_FIELD),
            (ChartKind::Line3d, LINE_3D_LIGHTING_FIELD),
            (ChartKind::Pie3d, PIE_3D_LIGHTING_FIELD),
            (ChartKind::Donut3d, PIE_3D_LIGHTING_FIELD),
        ];
        for (kind, expected) in cases {
            assert_eq!(LightingField::for_kind(kind).unwrap().number(), expected);
        }
        assert!(LightingField::for_kind(ChartKind::Column2d).is_none());
    }

    #[test]
    fn patch_preserves_other_fields_and_unknown_wire_data() {
        let mut extension = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(42.0),
            tschchartinfocolumn3dlightingpackage: native_lighting_package(
                Chart3dLightingStyle::SoftLight,
            ),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, 43).unwrap();
        let mut original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        original = patch_length_delimited_field(
            &original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(&extension),
        )
        .unwrap();
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, 44).unwrap();

        let patched = patch_chart_3d_lighting_style(
            &original,
            LightingField::Column,
            Chart3dLightingStyle::Glossy,
        )
        .unwrap();
        assert_eq!(
            read_chart_3d_lighting_style(&patched, LightingField::Column).unwrap(),
            Chart3dLightingStyle::Glossy
        );
        assert!(has_field(&patched, UNKNOWN_OUTER_FIELD));
        let extension = generated_chart_style_extension(&patched).unwrap().unwrap();
        assert!(has_field(extension, UNKNOWN_EXTENSION_FIELD));
        let generated = tsch::generated::ChartStyleArchive::decode(extension).unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultinterbargap, Some(42.0));
    }

    #[test]
    fn malformed_or_unknown_native_packages_are_rejected() {
        let mut package = native_lighting_package(Chart3dLightingStyle::SoftLight).unwrap();
        package.name = "Future Native Lighting".to_owned();
        let data = style_with_package(package);
        assert!(read_chart_3d_lighting_style(&data, LightingField::Bar).is_err());

        let mut package = native_lighting_package(Chart3dLightingStyle::SoftLight).unwrap();
        package.lights[0].intensity = f32::NAN;
        let data = style_with_package(package);
        assert!(read_chart_3d_lighting_style(&data, LightingField::Bar).is_err());
    }

    fn style_with_package(package: tsch::Chart3DLightingPackageArchive) -> Vec<u8> {
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfobar3dlightingpackage: Some(package),
            ..Default::default()
        }
        .encode_to_vec();
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        patch_length_delimited_field(
            &original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(&generated),
        )
        .unwrap()
    }

    fn has_field(data: &[u8], number: u32) -> bool {
        parse_wire_fields(data)
            .unwrap()
            .iter()
            .any(|field| field.number == number)
    }
}
