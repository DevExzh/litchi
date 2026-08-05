//! Lossless native 3D line/area series-gap storage and mutation.
//!
//! The Chart inspector exposes this value as the whole-percentage
//! `Between Series` control for 3D line and area charts. iWork stores the
//! normalized ratio in the generated chart non-style archive.

use prost::Message;

use crate::charts::Kind;
use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefault3dintersetdepthgap` in the generated chart non-style.
const CHART_3D_SERIES_GAP_FIELD: u32 = 3;
const MINIMUM_GAP_PERCENT: u16 = 0;
const MAXIMUM_GAP_PERCENT: u16 = 200;
const DEFAULT_GAP_PERCENT: u16 = 100;
const PERCENT_RATIO: f32 = 100.0;
/// Accommodates ordinary `f32` division/multiplication noise without
/// accepting a genuinely fractional native inspector value.
const NATIVE_WHOLE_PERCENT_TOLERANCE: f32 = 0.001;

/// Spacing between series along the depth axis of a 3D line or area chart.
///
/// iWork exposes only whole percentages in the inclusive `0%..=200%` range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Chart3dSeriesGap(u16);

impl Chart3dSeriesGap {
    /// No spacing between adjacent series.
    pub const ZERO: Self = Self(MINIMUM_GAP_PERCENT);

    /// The native value used when storage is absent.
    pub const NATIVE_DEFAULT: Self = Self(DEFAULT_GAP_PERCENT);

    /// The largest spacing accepted by the native inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_GAP_PERCENT);

    /// Construct a native 3D series-gap percentage.
    pub fn new(percent: u16) -> Result<Self> {
        if percent > MAXIMUM_GAP_PERCENT {
            return Err(Error::InvalidFormat(format!(
                "3D chart series gap must be within {MINIMUM_GAP_PERCENT}%..={MAXIMUM_GAP_PERCENT}%"
            )));
        }
        Ok(Self(percent))
    }

    /// Return the whole percentage displayed by iWork.
    pub const fn percent(self) -> u16 {
        self.0
    }

    const fn native_ratio(self) -> f32 {
        self.0 as f32 / PERCENT_RATIO
    }

    fn from_native_ratio(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::InvalidFormat(
                "native 3D chart series-gap ratio must be finite".to_owned(),
            ));
        }
        let percent = ratio * PERCENT_RATIO;
        let rounded = percent.round();
        if !(f32::from(MINIMUM_GAP_PERCENT)..=f32::from(MAXIMUM_GAP_PERCENT)).contains(&rounded)
            || (percent - rounded).abs() > NATIVE_WHOLE_PERCENT_TOLERANCE
        {
            return Err(Error::InvalidFormat(format!(
                "native 3D chart series gap {percent}% is outside the whole-percent {MINIMUM_GAP_PERCENT}%..={MAXIMUM_GAP_PERCENT}% range"
            )));
        }
        Ok(Self(rounded as u16))
    }
}

impl TryFrom<u16> for Chart3dSeriesGap {
    type Error = Error;

    fn try_from(percent: u16) -> Result<Self> {
        Self::new(percent)
    }
}

impl Default for Chart3dSeriesGap {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

/// Read one chart's effective native 3D line/area series gap.
pub(crate) fn chart_3d_series_gap(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
) -> Result<Chart3dSeriesGap> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_3d_series_gap)
}

/// Set one chart's native 3D line/area series gap.
pub(crate) fn set_chart_3d_series_gap(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    kind: Kind,
    gap: Chart3dSeriesGap,
) -> Result<()> {
    require_supported_kind(kind, drawable_object_id, drawable_label)?;
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_3d_series_gap)? == gap {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_3d_series_gap(data, gap))?;
    if slot.read(package, read_chart_3d_series_gap)? != gap {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} 3D series-gap update failed validation"
        )));
    }
    Ok(())
}

fn require_supported_kind(kind: Kind, drawable_object_id: u64, drawable_label: &str) -> Result<()> {
    if !kind.supports_3d_series_gap() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} kind {kind:?} has no 3D Between Series gap"
        )));
    }
    Ok(())
}

fn read_chart_3d_series_gap(data: &[u8]) -> Result<Chart3dSeriesGap> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(Chart3dSeriesGap::NATIVE_DEFAULT);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    generated
        .tschchartinfodefault3dintersetdepthgap
        .map(Chart3dSeriesGap::from_native_ratio)
        .transpose()
        .map(|gap| gap.unwrap_or(Chart3dSeriesGap::NATIVE_DEFAULT))
}

fn patch_chart_3d_series_gap(data: &[u8], gap: Chart3dSeriesGap) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        if gap == Chart3dSeriesGap::NATIVE_DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dintersetdepthgap: Some(gap.native_ratio()),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_3d_series_gap(&patched, gap)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let present = generated.tschchartinfodefault3dintersetdepthgap.is_some();
    let native = (gap != Chart3dSeriesGap::NATIVE_DEFAULT).then(|| gap.native_ratio().to_bits());
    let extension = patch_fixed32_field(extension, CHART_3D_SERIES_GAP_FIELD, present, native)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_3d_series_gap(&patched, gap)?;
    Ok(patched)
}

fn validate_patched_chart_3d_series_gap(data: &[u8], expected: Chart3dSeriesGap) -> Result<()> {
    if read_chart_3d_series_gap(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart 3D series-gap wire patch failed validation".to_owned(),
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
    fn percentages_match_native_bounds_default_and_ratio() {
        assert_eq!(
            Chart3dSeriesGap::default(),
            Chart3dSeriesGap::NATIVE_DEFAULT
        );
        assert_eq!(Chart3dSeriesGap::NATIVE_DEFAULT.percent(), 100);
        assert_eq!(Chart3dSeriesGap::MAXIMUM.percent(), 200);
        assert!(Chart3dSeriesGap::new(201).is_err());
        for percent in MINIMUM_GAP_PERCENT..=MAXIMUM_GAP_PERCENT {
            let gap = Chart3dSeriesGap::new(percent).unwrap();
            assert_eq!(
                Chart3dSeriesGap::from_native_ratio(gap.native_ratio()).unwrap(),
                gap
            );
        }
    }

    #[test]
    fn malformed_native_ratios_are_rejected() {
        for ratio in [f32::NAN, f32::INFINITY, -0.01, 2.01, 0.255] {
            assert!(Chart3dSeriesGap::from_native_ratio(ratio).is_err());
        }
    }

    #[test]
    fn capability_matches_only_unstacked_3d_line_and_area() {
        for kind in [Kind::Line3d, Kind::Area3d] {
            assert!(kind.supports_3d_series_gap(), "{kind:?}");
        }
        for kind in [
            Kind::Line2d,
            Kind::Area2d,
            Kind::Column3d,
            Kind::Bar3d,
            Kind::StackedArea3d,
            Kind::Pie3d,
        ] {
            assert!(!kind.supports_3d_series_gap(), "{kind:?}");
        }
    }

    #[test]
    fn patch_preserves_neighboring_and_unknown_fields() {
        let original = non_style_with_unknown_fields(tsch::generated::ChartNonStyleArchive {
            tschchartinfodefault3dintersetdepthgap: Some(0.25),
            tschchartinfodefault3dbarshape: Some(1),
            tschchartinfodefaultshowlegend: Some(true),
            ..Default::default()
        });
        let changed =
            patch_chart_3d_series_gap(&original, Chart3dSeriesGap::new(175).unwrap()).unwrap();
        assert_eq!(read_chart_3d_series_gap(&changed).unwrap().percent(), 175);
        assert_eq!(
            raw_field(&changed, UNKNOWN_OUTER_FIELD),
            raw_field(&original, UNKNOWN_OUTER_FIELD)
        );
        let changed_extension = generated_chart_non_style_extension(&changed)
            .unwrap()
            .unwrap();
        let original_extension = generated_chart_non_style_extension(&original)
            .unwrap()
            .unwrap();
        assert_eq!(
            raw_field(changed_extension, UNKNOWN_EXTENSION_FIELD),
            raw_field(original_extension, UNKNOWN_EXTENSION_FIELD)
        );
        let generated = tsch::generated::ChartNonStyleArchive::decode(changed_extension).unwrap();
        assert_eq!(generated.tschchartinfodefault3dbarshape, Some(1));
        assert_eq!(generated.tschchartinfodefaultshowlegend, Some(true));
    }

    #[test]
    fn default_is_sparse_and_nondefault_creates_storage() {
        let original = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        assert_eq!(
            read_chart_3d_series_gap(&original).unwrap(),
            Chart3dSeriesGap::NATIVE_DEFAULT
        );
        assert_eq!(
            patch_chart_3d_series_gap(&original, Chart3dSeriesGap::NATIVE_DEFAULT).unwrap(),
            original
        );

        let changed =
            patch_chart_3d_series_gap(&original, Chart3dSeriesGap::new(25).unwrap()).unwrap();
        assert_eq!(read_chart_3d_series_gap(&changed).unwrap().percent(), 25);
        let reset = patch_chart_3d_series_gap(&changed, Chart3dSeriesGap::NATIVE_DEFAULT).unwrap();
        let generated = tsch::generated::ChartNonStyleArchive::decode(
            generated_chart_non_style_extension(&reset)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefault3dintersetdepthgap, None);
    }

    fn non_style_with_unknown_fields(generated: tsch::generated::ChartNonStyleArchive) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNKNOWN_EXTENSION_FIELD, UNKNOWN_VALUE).unwrap();
        let mut outer = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, UNKNOWN_VALUE).unwrap();
        outer
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<u8> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number() == number)
            .flat_map(|field| data[field.start()..field.end()].iter().copied())
            .collect()
    }
}
