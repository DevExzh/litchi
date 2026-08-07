//! Lossless native radar start-angle storage and mutation.
//!
//! iWork exposes this value as `Rotation Angle` in the Radar chart inspector
//! and stores the chart-level default in `TSCH.ChartNonStyleArchive`. The
//! public scalar rejects noncanonical values while the wire patch preserves
//! every unrelated field byte-for-byte.

use prost::Message;

use crate::charts::Kind;
use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultradarstartangle` in the generated non-style archive.
const CHART_RADAR_START_ANGLE_FIELD: u32 = 29;
const MINIMUM_RADAR_START_ANGLE_DEGREES: f32 = 0.0;
const FULL_TURN_DEGREES: f32 = 360.0;

/// Rotation of the first radius in a native radar chart.
///
/// Angles use the clockwise degrees displayed by the iWork inspector. Values
/// must be finite and lie in the canonical half-open range `0°..360°`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartRadarStartAngle(f32);

impl ChartRadarStartAngle {
    pub const ZERO: Self = Self(MINIMUM_RADAR_START_ANGLE_DEGREES);
    pub const QUARTER_TURN: Self = Self(90.0);
    pub const HALF_TURN: Self = Self(180.0);
    pub const THREE_QUARTER_TURN: Self = Self(270.0);

    /// Construct a canonical angle accepted by the native Radar inspector.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite()
            || !(MINIMUM_RADAR_START_ANGLE_DEGREES..FULL_TURN_DEGREES).contains(&degrees)
        {
            return Err(Error::InvalidFormat(format!(
                "chart radar start angle must be finite and within {MINIMUM_RADAR_START_ANGLE_DEGREES}°..{FULL_TURN_DEGREES}°"
            )));
        }
        Ok(if degrees == MINIMUM_RADAR_START_ANGLE_DEGREES {
            Self::ZERO
        } else {
            Self(degrees)
        })
    }

    /// Return the clockwise degrees displayed by iWork.
    pub const fn degrees(self) -> f32 {
        self.0
    }

    fn from_native(degrees: f32) -> Result<Self> {
        Self::from_degrees(degrees).map_err(|_| {
            Error::InvalidFormat(format!(
                "native chart radar start angle {degrees} must be finite and within {MINIMUM_RADAR_START_ANGLE_DEGREES}°..{FULL_TURN_DEGREES}°"
            ))
        })
    }
}

impl Default for ChartRadarStartAngle {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for ChartRadarStartAngle {
    type Error = Error;

    fn try_from(degrees: f32) -> Result<Self> {
        Self::from_degrees(degrees)
    }
}

/// Read the effective start angle of one native radar chart.
pub(crate) fn chart_radar_start_angle(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
) -> Result<ChartRadarStartAngle> {
    require_radar_chart(kind)?;
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_radar_start_angle)
}

/// Set the start angle of one native radar chart.
pub(crate) fn set_chart_radar_start_angle(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    angle: ChartRadarStartAngle,
) -> Result<()> {
    require_radar_chart(kind)?;
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_radar_start_angle)? == angle {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_radar_start_angle(data, angle))?;
    if slot.read(package, read_chart_radar_start_angle)? != angle {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} radar start-angle update failed validation"
        )));
    }
    Ok(())
}

fn require_radar_chart(kind: Kind) -> Result<()> {
    if !kind.supports_radar_start_angle() {
        return Err(Error::InvalidFormat(format!(
            "chart kind {kind:?} does not expose a radar rotation angle"
        )));
    }
    Ok(())
}

fn read_chart_radar_start_angle(data: &[u8]) -> Result<ChartRadarStartAngle> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(ChartRadarStartAngle::ZERO);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    generated
        .tschchartinfodefaultradarstartangle
        .map(ChartRadarStartAngle::from_native)
        .transpose()
        .map(|angle| angle.unwrap_or(ChartRadarStartAngle::ZERO))
}

fn patch_chart_radar_start_angle(data: &[u8], angle: ChartRadarStartAngle) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if angle == ChartRadarStartAngle::ZERO {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultradarstartangle: Some(angle.degrees()),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_radar_start_angle(&patched, angle)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let present = generated.tschchartinfodefaultradarstartangle.is_some();
    let native = (angle != ChartRadarStartAngle::ZERO).then(|| angle.degrees().to_bits());
    let extension = patch_fixed32_field(extension, CHART_RADAR_START_ANGLE_FIELD, present, native)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_radar_start_angle(&patched, angle)?;
    Ok(patched)
}

fn validate_patched_chart_radar_start_angle(
    data: &[u8],
    expected: ChartRadarStartAngle,
) -> Result<()> {
    if read_chart_radar_start_angle(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart radar start-angle wire patch failed validation".to_owned(),
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
    const UNKNOWN_VALUE: u64 = 42;

    #[test]
    fn radar_start_angles_are_strict_and_native_default_is_zero() {
        assert_eq!(ChartRadarStartAngle::default(), ChartRadarStartAngle::ZERO);
        assert_eq!(
            ChartRadarStartAngle::from_degrees(359.5).unwrap().degrees(),
            359.5
        );
        assert_eq!(
            ChartRadarStartAngle::from_degrees(-0.0)
                .unwrap()
                .degrees()
                .to_bits(),
            0.0f32.to_bits()
        );
        for invalid in [f32::NEG_INFINITY, -0.1, 360.0, f32::INFINITY, f32::NAN] {
            assert!(ChartRadarStartAngle::from_degrees(invalid).is_err());
        }

        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            read_chart_radar_start_angle(&base).unwrap(),
            ChartRadarStartAngle::ZERO
        );
        assert_eq!(
            patch_chart_radar_start_angle(&base, ChartRadarStartAngle::ZERO).unwrap(),
            base
        );
    }

    #[test]
    fn radar_start_angle_patch_is_lossless_and_resets_exactly() {
        let mut extension = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(true),
            tschchartinfodefaulttitle: Some("Radar comparison".to_owned()),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
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
        append_varint_field(&mut original, UNKNOWN_OUTER_FIELD, UNKNOWN_VALUE).unwrap();
        let customized = ChartRadarStartAngle::from_degrees(45.0).unwrap();

        let patched = patch_chart_radar_start_angle(&original, customized).unwrap();
        assert_eq!(read_chart_radar_start_angle(&patched).unwrap(), customized);
        assert_eq!(
            raw_field(&patched, UNKNOWN_OUTER_FIELD),
            raw_field(&original, UNKNOWN_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&patched)
                    .unwrap()
                    .unwrap(),
                UNKNOWN_EXTENSION_FIELD,
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNKNOWN_EXTENSION_FIELD,
            )
        );
        assert_eq!(
            patch_chart_radar_start_angle(&patched, ChartRadarStartAngle::ZERO).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_radar_start_angles_are_rejected() {
        for invalid in [-1.0, 360.0, f32::INFINITY, f32::NAN] {
            let generated = tsch::generated::ChartNonStyleArchive {
                tschchartinfodefaultradarstartangle: Some(invalid),
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
            assert!(read_chart_radar_start_angle(&data).is_err());
        }
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
