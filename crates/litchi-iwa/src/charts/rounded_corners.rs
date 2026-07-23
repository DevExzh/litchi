//! Lossless native chart rounded-corner storage and mutation.
//!
//! iWork stores a chart's `Corner Radius` percentage and `Outside Corners
//! Only` switch in the generated extension of `TSCH.ChartStyleArchive`. The
//! public types model those inspector values directly while this module
//! preserves every unrelated protobuf field byte-for-byte.

use prost::Message;

use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultroundedcornerradius` in
/// `TSCH.Generated.ChartStyleArchive`.
const CHART_CORNER_RADIUS_FIELD: u32 = 122;
/// `tschchartinfodefaultroundedcornerouterendonly` in
/// `TSCH.Generated.ChartStyleArchive`.
const CHART_OUTSIDE_CORNERS_ONLY_FIELD: u32 = 123;
/// The smallest selectable chart corner-radius percentage.
const MINIMUM_CORNER_RADIUS_PERCENT: f32 = 0.0;
/// The largest selectable chart corner-radius percentage.
const MAXIMUM_CORNER_RADIUS_PERCENT: f32 = 100.0;
/// The native stored ratio corresponding to a 100% inspector radius.
const MAXIMUM_NATIVE_CORNER_RADIUS: f32 = 1.0;

/// A chart corner radius shown by iWork as a percentage.
///
/// The native archive stores this as a normalized ratio. This type keeps that
/// ratio internally so values read from a document compare exactly with values
/// supplied through [`Self::new`].
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartCornerRadius(f32);

impl ChartCornerRadius {
    /// Square chart corners.
    pub const ZERO: Self = Self(MINIMUM_CORNER_RADIUS_PERCENT);

    /// Build a native chart corner radius from the inspector percentage.
    pub fn new(percent: f32) -> Result<Self> {
        if !percent.is_finite()
            || !(MINIMUM_CORNER_RADIUS_PERCENT..=MAXIMUM_CORNER_RADIUS_PERCENT).contains(&percent)
        {
            return Err(Error::InvalidFormat(format!(
                "chart corner radius must be finite and within {MINIMUM_CORNER_RADIUS_PERCENT}%..={MAXIMUM_CORNER_RADIUS_PERCENT}%"
            )));
        }
        Ok(Self(percent / MAXIMUM_CORNER_RADIUS_PERCENT))
    }

    /// Return the percentage displayed by iWork's `Corner Radius` field.
    pub const fn percent(self) -> f32 {
        self.0 * MAXIMUM_CORNER_RADIUS_PERCENT
    }

    fn from_native(value: f32) -> Result<Self> {
        if !value.is_finite()
            || !(MINIMUM_CORNER_RADIUS_PERCENT..=MAXIMUM_NATIVE_CORNER_RADIUS).contains(&value)
        {
            return Err(Error::InvalidFormat(
                "native chart corner radius must be a finite normalized percentage".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    const fn native(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for ChartCornerRadius {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::new(percent)
    }
}

/// Native rounded-corner settings for a chart.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartRoundedCorners {
    radius: ChartCornerRadius,
    outside_corners_only: bool,
}

impl ChartRoundedCorners {
    /// The native default: square corners with no outside-corners override.
    pub const NONE: Self = Self::new(ChartCornerRadius::ZERO, false);

    /// Construct rounded-corner settings from an inspector radius and switch.
    pub const fn new(radius: ChartCornerRadius, outside_corners_only: bool) -> Self {
        Self {
            radius,
            outside_corners_only,
        }
    }

    /// Return the chart's `Corner Radius` setting.
    pub const fn radius(self) -> ChartCornerRadius {
        self.radius
    }

    /// Return whether `Outside Corners Only` is enabled.
    pub const fn outside_corners_only(self) -> bool {
        self.outside_corners_only
    }
}

impl Default for ChartRoundedCorners {
    fn default() -> Self {
        Self::NONE
    }
}

/// Read the rounded-corner settings of one native chart.
pub(crate) fn chart_rounded_corners(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartRoundedCorners> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_rounded_corners)
}

/// Set the rounded-corner settings of one native chart.
pub(crate) fn set_chart_rounded_corners(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    rounded_corners: ChartRoundedCorners,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_rounded_corners)? == rounded_corners {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_rounded_corners(data, rounded_corners)
    })?;
    if slot.read(package, read_chart_rounded_corners)? != rounded_corners {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} rounded-corner update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_rounded_corners(data: &[u8]) -> Result<ChartRoundedCorners> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(ChartRoundedCorners::NONE);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    Ok(ChartRoundedCorners::new(
        generated
            .tschchartinfodefaultroundedcornerradius
            .map(ChartCornerRadius::from_native)
            .transpose()?
            .unwrap_or(ChartCornerRadius::ZERO),
        generated
            .tschchartinfodefaultroundedcornerouterendonly
            .unwrap_or(false),
    ))
}

fn patch_chart_rounded_corners(
    data: &[u8],
    rounded_corners: ChartRoundedCorners,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        if rounded_corners == ChartRoundedCorners::NONE {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultroundedcornerradius: (rounded_corners.radius()
                != ChartCornerRadius::ZERO)
                .then(|| rounded_corners.radius().native()),
            tschchartinfodefaultroundedcornerouterendonly: rounded_corners
                .outside_corners_only()
                .then_some(true),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_chart_rounded_corners(&patched, rounded_corners)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let radius_present = generated.tschchartinfodefaultroundedcornerradius.is_some();
    let outside_corners_only_present = generated
        .tschchartinfodefaultroundedcornerouterendonly
        .is_some();
    let radius = if rounded_corners == ChartRoundedCorners::NONE {
        None
    } else {
        (radius_present || rounded_corners.radius() != ChartCornerRadius::ZERO)
            .then(|| rounded_corners.radius().native().to_bits())
    };
    let outside_corners_only = if rounded_corners == ChartRoundedCorners::NONE {
        None
    } else {
        (outside_corners_only_present || rounded_corners.outside_corners_only())
            .then_some(u64::from(rounded_corners.outside_corners_only()))
    };
    let extension =
        patch_fixed32_field(extension, CHART_CORNER_RADIUS_FIELD, radius_present, radius)?;
    let extension = patch_varint_field(
        &extension,
        CHART_OUTSIDE_CORNERS_ONLY_FIELD,
        outside_corners_only_present,
        outside_corners_only,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_rounded_corners(&patched, rounded_corners)?;
    Ok(patched)
}

fn validate_patched_chart_rounded_corners(
    data: &[u8],
    expected: ChartRoundedCorners,
) -> Result<()> {
    if read_chart_rounded_corners(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart rounded-corner wire patch failed validation".to_owned(),
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
    fn corner_radius_rejects_invalid_percentages_and_native_values() {
        assert!(ChartCornerRadius::new(f32::NAN).is_err());
        assert!(ChartCornerRadius::new(f32::INFINITY).is_err());
        assert!(ChartCornerRadius::new(-0.1).is_err());
        assert!(ChartCornerRadius::new(100.1).is_err());
        assert!(ChartCornerRadius::from_native(-0.1).is_err());
        assert!(ChartCornerRadius::from_native(1.1).is_err());
        assert_eq!(ChartCornerRadius::new(20.0).unwrap().percent(), 20.0);
    }

    #[test]
    fn rounded_corner_patch_retains_other_style_fields_and_unmapped_data() {
        let original_settings =
            ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);
        let replacement = ChartRoundedCorners::new(ChartCornerRadius::new(35.0).unwrap(), false);
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(0.2),
            tschchartinfodefaultroundedcornerradius: Some(original_settings.radius().native()),
            tschchartinfodefaultroundedcornerouterendonly: Some(
                original_settings.outside_corners_only(),
            ),
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

        let patched = patch_chart_rounded_corners(&original, replacement).unwrap();
        assert_eq!(read_chart_rounded_corners(&patched).unwrap(), replacement);
        let patched_generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&patched).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(patched_generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(patched_generated.tschchartinfodefaultinterbargap, Some(0.2));
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

        let restored = patch_chart_rounded_corners(&patched, original_settings).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn rounded_corners_default_to_none_and_create_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let rounded_corners = ChartRoundedCorners::new(ChartCornerRadius::new(20.0).unwrap(), true);

        assert_eq!(
            read_chart_rounded_corners(&original).unwrap(),
            ChartRoundedCorners::NONE
        );
        assert_eq!(
            patch_chart_rounded_corners(&original, ChartRoundedCorners::NONE).unwrap(),
            original
        );

        let patched = patch_chart_rounded_corners(&original, rounded_corners).unwrap();
        assert_eq!(
            read_chart_rounded_corners(&patched).unwrap(),
            rounded_corners
        );
        assert!(generated_chart_style_extension(&patched).unwrap().is_some());
    }

    #[test]
    fn resetting_rounded_corners_retains_other_style_fields() {
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultroundedcornerradius: Some(0.2),
            tschchartinfodefaultroundedcornerouterendonly: Some(true),
            ..Default::default()
        });

        let reset = patch_chart_rounded_corners(&original, ChartRoundedCorners::NONE).unwrap();
        assert_eq!(
            read_chart_rounded_corners(&reset).unwrap(),
            ChartRoundedCorners::NONE
        );
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&reset).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultroundedcornerradius, None);
        assert_eq!(
            generated.tschchartinfodefaultroundedcornerouterendonly,
            None
        );
        assert_unknown_fields_retained(&original, &reset);
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
            .filter(|field| field.number == number)
            .map(|field| data[field.start..field.end].to_vec())
            .collect()
    }
}
