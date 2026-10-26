//! Lossless per-wedge position CRUD for native pie and donut charts.
//!
//! iWork stores each wedge's distance from the chart center in a sparse array
//! of private `TSCH.ChartSeriesNonStyleArchive` objects. The shared series
//! non-style layer owns that graph; this module only validates and losslessly
//! patches the generated wedge-explosion scalar.

use litchi_iwa_common::chart::pie::{WedgeExplosion, WedgeIndex};
use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::patch_fixed32_field;
use crate::{Error, HostPackage, Result};

/// `tschchartseriespiewedgeexplosion` in the generated series non-style.
const PIE_WEDGE_EXPLOSION_FIELD: u32 = 63;
const MINIMUM_WEDGE_EXPLOSION_PERCENT: f32 = 0.0;
const MAXIMUM_WEDGE_EXPLOSION_PERCENT: f32 = 100.0;

/// Distance of one pie or donut wedge from the chart center.
///
/// Values use the percentage displayed by the Wedges inspector and must be
/// finite in the inclusive range `0%..=100%`. The value is the shared
/// `litchi_iwa_common::chart::pie::WedgeExplosion` (record 0375).
#[derive(Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartPieWedgeExplosion(WedgeExplosion);

impl std::fmt::Debug for ChartPieWedgeExplosion {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ChartPieWedgeExplosion")
            .field(&self.0.native_fraction())
            .finish()
    }
}

impl ChartPieWedgeExplosion {
    /// Wedge at its native, unseparated position.
    pub const ZERO: Self = Self(WedgeExplosion::ZERO);
    /// Wedge at the inspector's maximum distance from the center.
    pub const MAXIMUM: Self = Self(WedgeExplosion::MAXIMUM);

    /// Construct a wedge position from an inspector percentage.
    pub fn from_percent(percent: f32) -> Result<Self> {
        WedgeExplosion::from_percent(percent).map(Self).map_err(|_| {
            Error::InvalidFormat(format!(
                "chart pie wedge explosion must be finite and within {MINIMUM_WEDGE_EXPLOSION_PERCENT}%..={MAXIMUM_WEDGE_EXPLOSION_PERCENT}%"
            ))
        })
    }

    /// Return the percentage displayed by iWork.
    pub fn percent(self) -> f32 {
        self.0.percent()
    }

    fn from_native(fraction: f32) -> Result<Self> {
        WedgeExplosion::from_native(fraction).map(Self).map_err(|_| {
            Error::InvalidFormat(format!(
                "native chart pie wedge explosion {fraction} must be finite and within 0.0..=1.0"
            ))
        })
    }

    const fn native_fraction(self) -> f32 {
        self.0.native_fraction()
    }
}

impl Default for ChartPieWedgeExplosion {
    fn default() -> Self {
        Self::ZERO
    }
}

impl TryFrom<f32> for ChartPieWedgeExplosion {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

/// Zero-based index of one pie or donut wedge in chart series order (the
/// shared `litchi_iwa_common::chart::pie::WedgeIndex`).
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ChartPieWedgeIndex(WedgeIndex);

impl std::fmt::Debug for ChartPieWedgeIndex {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_tuple("ChartPieWedgeIndex")
            .field(&self.0.zero_based())
            .finish()
    }
}

impl ChartPieWedgeIndex {
    /// Construct a zero-based wedge index.
    pub const fn from_zero_based(index: usize) -> Self {
        Self(WedgeIndex::from_zero_based(index))
    }

    /// Return the zero-based wedge index.
    pub const fn zero_based(self) -> usize {
        self.0.zero_based()
    }
}

/// Read every wedge position in chart series order.
pub(crate) fn chart_pie_wedge_explosions(
    package: &HostPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartPieWedgeExplosion>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartPieWedgeExplosion::ZERO,
        read_series_non_style_explosion,
    )
}

/// Set every wedge position in chart series order.
pub(crate) fn set_chart_pie_wedge_explosions(
    package: &mut HostPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartPieWedgeExplosion],
) -> Result<()> {
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "pie wedge explosion",
        NewChartSeriesNonStyleBase::Unstyled,
        expected,
        ChartPieWedgeExplosion::ZERO,
        read_series_non_style_explosion,
        |data, explosion| patch_series_non_style_explosion(data, *explosion),
    )
}

fn read_series_non_style_explosion(data: &[u8]) -> Result<ChartPieWedgeExplosion> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartPieWedgeExplosion::ZERO);
    };
    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    generated
        .tschchartseriespiewedgeexplosion
        .map(ChartPieWedgeExplosion::from_native)
        .transpose()
        .map(|value| value.unwrap_or(ChartPieWedgeExplosion::ZERO))
}

fn patch_series_non_style_explosion(
    data: &[u8],
    explosion: ChartPieWedgeExplosion,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        if explosion == ChartPieWedgeExplosion::ZERO {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespiewedgeexplosion: Some(explosion.native_fraction()),
            ..Default::default()
        };
        let patched = patch_chart_series_non_style_extension(
            data,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_explosion(&patched, explosion)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let present = generated.tschchartseriespiewedgeexplosion.is_some();
    let native =
        (explosion != ChartPieWedgeExplosion::ZERO).then(|| explosion.native_fraction().to_bits());
    let extension = patch_fixed32_field(extension, PIE_WEDGE_EXPLOSION_FIELD, present, native)?;
    let patched = patch_chart_series_non_style_extension(
        data,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_explosion(&patched, explosion)?;
    Ok(patched)
}

fn validate_patched_explosion(data: &[u8], expected: ChartPieWedgeExplosion) -> Result<()> {
    if read_series_non_style_explosion(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart pie wedge-explosion wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::series_non_style::{
        GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
        canonical_empty_chart_series_non_style_data,
    };
    use crate::protobuf::tss;
    use crate::wire::{
        append_length_delimited_field, append_varint_field, parse_wire_fields,
        patch_length_delimited_field,
    };

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn wedge_explosions_are_strict_percentages() {
        assert_eq!(
            ChartPieWedgeExplosion::default(),
            ChartPieWedgeExplosion::ZERO
        );
        assert_eq!(
            ChartPieWedgeExplosion::from_percent(25.0)
                .unwrap()
                .percent(),
            25.0
        );
        assert_eq!(
            ChartPieWedgeExplosion::from_percent(100.0).unwrap(),
            ChartPieWedgeExplosion::MAXIMUM
        );
        for invalid in [f32::NEG_INFINITY, -0.1, 100.1, f32::INFINITY, f32::NAN] {
            assert!(ChartPieWedgeExplosion::from_percent(invalid).is_err());
        }
    }

    #[test]
    fn wedge_explosion_patch_is_lossless_and_resets_exactly() {
        let mut generated = tsch::generated::ChartSeriesNonStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = tsch::ChartSeriesNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        let customized = ChartPieWedgeExplosion::from_percent(25.0).unwrap();

        let patched = patch_series_non_style_explosion(&original, customized).unwrap();
        assert_eq!(
            read_series_non_style_explosion(&patched).unwrap(),
            customized
        );
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_series_non_style_extension(&patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_series_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
        assert_eq!(
            patch_series_non_style_explosion(&patched, ChartPieWedgeExplosion::ZERO).unwrap(),
            original
        );
    }

    #[test]
    fn malformed_native_wedge_explosions_are_rejected() {
        for invalid in [-0.1, 1.1, f32::INFINITY, f32::NAN] {
            let mut outer = canonical_empty_chart_series_non_style_data().unwrap();
            let generated = tsch::generated::ChartSeriesNonStyleArchive {
                tschchartseriespiewedgeexplosion: Some(invalid),
                ..Default::default()
            };
            outer = patch_length_delimited_field(
                &outer,
                GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
                false,
                Some(generated.encode_to_vec().as_slice()),
            )
            .unwrap();
            assert!(read_series_non_style_explosion(&outer).is_err());
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
