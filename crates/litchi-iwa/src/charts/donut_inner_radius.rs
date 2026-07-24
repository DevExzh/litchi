//! Lossless native donut inner-radius storage and mutation.
//!
//! iWork exposes this value as `Inner Radius` in the Segments inspector and
//! stores the chart-level default in `TSCH.ChartNonStyleArchive`. The public
//! scalar uses the inspector percentage while the wire patch preserves every
//! unrelated field byte-for-byte.

use prost::Message;

use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultinnerradius` in the generated non-style archive.
const CHART_DONUT_INNER_RADIUS_FIELD: u32 = 27;
const MINIMUM_DONUT_INNER_RADIUS_PERCENT: f32 = 1.0;
const DEFAULT_DONUT_INNER_RADIUS_PERCENT: f32 = 75.0;
const MAXIMUM_DONUT_INNER_RADIUS_PERCENT: f32 = 99.0;
const PERCENT_SCALE: f32 = 100.0;

/// Size of the center hole in a native donut chart.
///
/// Values use the percentage displayed by the Segments inspector and must be
/// finite in the inclusive range `1%..=99%`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartDonutInnerRadius(f32);

impl ChartDonutInnerRadius {
    /// Smallest center hole accepted by the native inspector.
    pub const MINIMUM: Self = Self(MINIMUM_DONUT_INNER_RADIUS_PERCENT / PERCENT_SCALE);
    /// Native center-hole size used when the property is absent.
    pub const DEFAULT: Self = Self(DEFAULT_DONUT_INNER_RADIUS_PERCENT / PERCENT_SCALE);
    /// Largest center hole accepted by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_DONUT_INNER_RADIUS_PERCENT / PERCENT_SCALE);

    /// Construct an inner radius from an inspector percentage.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite()
            || !(MINIMUM_DONUT_INNER_RADIUS_PERCENT..=MAXIMUM_DONUT_INNER_RADIUS_PERCENT)
                .contains(&percent)
        {
            return Err(Error::InvalidFormat(format!(
                "chart donut inner radius must be finite and within {MINIMUM_DONUT_INNER_RADIUS_PERCENT}%..={MAXIMUM_DONUT_INNER_RADIUS_PERCENT}%"
            )));
        }
        Ok(Self(percent / PERCENT_SCALE))
    }

    /// Return the percentage displayed by iWork.
    pub fn percent(self) -> f32 {
        self.0 * PERCENT_SCALE
    }

    fn from_native(fraction: f32) -> Result<Self> {
        if !fraction.is_finite() || !(Self::MINIMUM.0..=Self::MAXIMUM.0).contains(&fraction) {
            return Err(Error::InvalidFormat(format!(
                "native chart donut inner radius {fraction} must be finite and within {}..={}",
                Self::MINIMUM.0,
                Self::MAXIMUM.0
            )));
        }
        Ok(Self(fraction))
    }

    const fn native_fraction(self) -> f32 {
        self.0
    }
}

impl Default for ChartDonutInnerRadius {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<f32> for ChartDonutInnerRadius {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

/// Read the effective inner radius of one native donut chart.
pub(crate) fn chart_donut_inner_radius(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartDonutInnerRadius> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_donut_inner_radius)
}

/// Set the inner radius of one native donut chart.
pub(crate) fn set_chart_donut_inner_radius(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    radius: ChartDonutInnerRadius,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_donut_inner_radius)? == radius {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_donut_inner_radius(data, radius))?;
    if slot.read(package, read_chart_donut_inner_radius)? != radius {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} donut inner-radius update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_donut_inner_radius(data: &[u8]) -> Result<ChartDonutInnerRadius> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(ChartDonutInnerRadius::DEFAULT);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    generated
        .tschchartinfodefaultinnerradius
        .map(ChartDonutInnerRadius::from_native)
        .transpose()
        .map(|radius| radius.unwrap_or(ChartDonutInnerRadius::DEFAULT))
}

fn patch_chart_donut_inner_radius(data: &[u8], radius: ChartDonutInnerRadius) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if radius == ChartDonutInnerRadius::DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultinnerradius: Some(radius.native_fraction()),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_donut_inner_radius(&patched, radius)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let radius_present = generated.tschchartinfodefaultinnerradius.is_some();
    let native =
        (radius != ChartDonutInnerRadius::DEFAULT).then(|| radius.native_fraction().to_bits());
    let extension = patch_fixed32_field(
        extension,
        CHART_DONUT_INNER_RADIUS_FIELD,
        radius_present,
        native,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_donut_inner_radius(&patched, radius)?;
    Ok(patched)
}

fn validate_patched_chart_donut_inner_radius(
    data: &[u8],
    expected: ChartDonutInnerRadius,
) -> Result<()> {
    if read_chart_donut_inner_radius(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart donut inner-radius wire patch failed validation".to_owned(),
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
    fn donut_inner_radii_are_strict_and_default_to_seventy_five_percent() {
        assert_eq!(
            ChartDonutInnerRadius::default(),
            ChartDonutInnerRadius::DEFAULT
        );
        assert_eq!(
            ChartDonutInnerRadius::from_percent(42.0).unwrap().percent(),
            42.0
        );
        for invalid in [
            f32::NEG_INFINITY,
            0.0,
            0.999,
            99.001,
            100.0,
            f32::INFINITY,
            f32::NAN,
        ] {
            assert!(ChartDonutInnerRadius::from_percent(invalid).is_err());
        }

        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            read_chart_donut_inner_radius(&base).unwrap(),
            ChartDonutInnerRadius::DEFAULT
        );
        assert_eq!(
            patch_chart_donut_inner_radius(&base, ChartDonutInnerRadius::DEFAULT).unwrap(),
            base
        );
    }

    #[test]
    fn donut_inner_radius_patch_is_lossless_and_resets_exactly() {
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(true),
            tschchartinfodefaulttitle: Some("Regional revenue".to_owned()),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
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
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        let customized = ChartDonutInnerRadius::from_percent(42.0).unwrap();

        let patched = patch_chart_donut_inner_radius(&original, customized).unwrap();
        assert_eq!(read_chart_donut_inner_radius(&patched).unwrap(), customized);
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
        assert_eq!(
            patch_chart_donut_inner_radius(&patched, ChartDonutInnerRadius::DEFAULT).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_donut_inner_radii_are_rejected() {
        for invalid in [0.0, 0.009, 0.991, 1.0, f32::INFINITY, f32::NAN] {
            let generated = tsch::generated::ChartNonStyleArchive {
                tschchartinfodefaultinnerradius: Some(invalid),
                ..Default::default()
            };
            let mut data = tsch::ChartNonStyleArchive {
                super_: Some(tss::StyleArchive::default()),
            }
            .encode_to_vec();
            append_length_delimited_field(
                &mut data,
                GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
                &generated.encode_to_vec(),
            )
            .unwrap();
            assert!(read_chart_donut_inner_radius(&data).is_err());
        }
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
