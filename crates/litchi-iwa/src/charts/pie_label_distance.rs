//! Lossless per-wedge data-label distance CRUD for native pie and donut charts.

use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_fixed32_field};
use crate::{Error, IWorkPackage, Result};

/// Legacy `tschchartseriespie2_3labelexplosion`.
const LEGACY_PIE_LABEL_DISTANCE_FIELD: u32 = 16;
/// Modern `tschchartseriespielabelexplosion`.
const PIE_LABEL_DISTANCE_FIELD: u32 = 147;
const MINIMUM_LABEL_DISTANCE_PERCENT: f32 = 30.0;
const DEFAULT_LABEL_DISTANCE_PERCENT: f32 = 67.0;
const MAXIMUM_LABEL_DISTANCE_PERCENT: f32 = 200.0;

/// Distance of a pie or donut wedge's label from the chart center.
///
/// Values use the percentage displayed by iWork and must be finite in the
/// inclusive native range `30%..=200%`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartPieLabelDistance(f32);

impl ChartPieLabelDistance {
    /// Nearest label position accepted by the iWork inspector.
    pub const MINIMUM: Self = Self(MINIMUM_LABEL_DISTANCE_PERCENT);
    /// Native label position when no per-wedge override is stored.
    pub const DEFAULT: Self = Self(DEFAULT_LABEL_DISTANCE_PERCENT);
    /// Farthest label position accepted by the iWork inspector.
    pub const MAXIMUM: Self = Self(MAXIMUM_LABEL_DISTANCE_PERCENT);

    /// Construct a label distance from an inspector percentage.
    pub fn from_percent(percent: f32) -> Result<Self> {
        if !percent.is_finite()
            || !(MINIMUM_LABEL_DISTANCE_PERCENT..=MAXIMUM_LABEL_DISTANCE_PERCENT).contains(&percent)
        {
            return Err(Error::InvalidFormat(format!(
                "chart pie label distance must be finite and within {MINIMUM_LABEL_DISTANCE_PERCENT}%..={MAXIMUM_LABEL_DISTANCE_PERCENT}%"
            )));
        }
        Ok(Self(percent))
    }

    /// Return the percentage displayed by iWork.
    pub const fn percent(self) -> f32 {
        self.0
    }

    fn native_percent(self) -> f32 {
        self.0
    }
}

impl Default for ChartPieLabelDistance {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<f32> for ChartPieLabelDistance {
    type Error = Error;

    fn try_from(percent: f32) -> Result<Self> {
        Self::from_percent(percent)
    }
}

/// Read every wedge's label distance in chart-series order.
pub(crate) fn chart_pie_label_distances(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartPieLabelDistance>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartPieLabelDistance::DEFAULT,
        read_series_non_style_label_distance,
    )
}

/// Set every wedge's label distance in chart-series order.
pub(crate) fn set_chart_pie_label_distances(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartPieLabelDistance],
) -> Result<()> {
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "pie label distance",
        NewChartSeriesNonStyleBase::Unstyled,
        expected,
        ChartPieLabelDistance::DEFAULT,
        read_series_non_style_label_distance,
        |data, distance| patch_series_non_style_label_distance(data, *distance),
    )
}

fn read_series_non_style_label_distance(data: &[u8]) -> Result<ChartPieLabelDistance> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartPieLabelDistance::DEFAULT);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    let legacy = strict_optional_distance(extension, LEGACY_PIE_LABEL_DISTANCE_FIELD)?;
    let modern = strict_optional_distance(extension, PIE_LABEL_DISTANCE_FIELD)?;
    match (legacy, modern) {
        (Some(legacy), Some(modern)) if legacy != modern => Err(Error::InvalidFormat(format!(
            "native chart pie label distances disagree: {}% != {}%",
            legacy.percent(),
            modern.percent()
        ))),
        (Some(distance), _) | (_, Some(distance)) => Ok(distance),
        (None, None) => Ok(ChartPieLabelDistance::DEFAULT),
    }
}

fn patch_series_non_style_label_distance(
    data: &[u8],
    distance: ChartPieLabelDistance,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        if distance == ChartPieLabelDistance::DEFAULT {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespie2_3labelexplosion: Some(distance.native_percent()),
            tschchartseriespielabelexplosion: Some(distance.native_percent()),
            ..Default::default()
        };
        let patched = patch_chart_series_non_style_extension(
            data,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_label_distance(&patched, distance)?;
        return Ok(patched);
    };

    let legacy_present =
        strict_optional_distance(extension, LEGACY_PIE_LABEL_DISTANCE_FIELD)?.is_some();
    let modern_present = strict_optional_distance(extension, PIE_LABEL_DISTANCE_FIELD)?.is_some();
    let replacement =
        (distance != ChartPieLabelDistance::DEFAULT).then(|| distance.native_percent().to_bits());
    let extension = patch_fixed32_field(
        extension,
        LEGACY_PIE_LABEL_DISTANCE_FIELD,
        legacy_present,
        replacement,
    )?;
    let extension = patch_fixed32_field(
        &extension,
        PIE_LABEL_DISTANCE_FIELD,
        modern_present,
        replacement,
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_label_distance(&patched, distance)?;
    Ok(patched)
}

fn strict_optional_distance(
    data: &[u8],
    field_number: u32,
) -> Result<Option<ChartPieLabelDistance>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart pie label distance field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 5 || field.end - field.key_end != size_of::<f32>() {
        return Err(Error::InvalidFormat(format!(
            "chart pie label distance field {field_number} is not fixed32"
        )));
    }
    let bytes: [u8; size_of::<f32>()] = data[field.key_end..field.end]
        .try_into()
        .map_err(|_| Error::InvalidFormat("chart pie label distance is truncated".to_owned()))?;
    ChartPieLabelDistance::from_percent(f32::from_le_bytes(bytes)).map(Some)
}

fn validate_patched_label_distance(data: &[u8], expected: ChartPieLabelDistance) -> Result<()> {
    if read_series_non_style_label_distance(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart pie label-distance wire patch failed validation".to_owned(),
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
    use crate::wire::{
        append_length_delimited_field, append_varint_field, parse_wire_fields,
        patch_length_delimited_field,
    };

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn label_distances_are_strict_native_percentages() {
        assert_eq!(
            ChartPieLabelDistance::default(),
            ChartPieLabelDistance::DEFAULT
        );
        assert_eq!(
            ChartPieLabelDistance::from_percent(150.0)
                .unwrap()
                .percent(),
            150.0
        );
        assert_eq!(
            ChartPieLabelDistance::from_percent(30.0).unwrap(),
            ChartPieLabelDistance::MINIMUM
        );
        assert_eq!(
            ChartPieLabelDistance::from_percent(200.0).unwrap(),
            ChartPieLabelDistance::MAXIMUM
        );
        for invalid in [f32::NEG_INFINITY, 29.9, 200.1, f32::INFINITY, f32::NAN] {
            assert!(ChartPieLabelDistance::from_percent(invalid).is_err());
        }
    }

    #[test]
    fn label_distance_patch_is_lossless_and_resets_exactly() {
        let mut generated = tsch::generated::ChartSeriesNonStyleArchive::default().encode_to_vec();
        append_varint_field(&mut generated, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut original = canonical_empty_chart_series_non_style_data().unwrap();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        let customized = ChartPieLabelDistance::from_percent(150.0).unwrap();

        let patched = patch_series_non_style_label_distance(&original, customized).unwrap();
        assert_eq!(
            read_series_non_style_label_distance(&patched).unwrap(),
            customized
        );
        assert_eq!(
            raw_field(&patched, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        let extension = generated_chart_series_non_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            raw_field(extension, UNMAPPED_GENERATED_FIELD),
            raw_field(&generated, UNMAPPED_GENERATED_FIELD)
        );
        assert_eq!(
            raw_field(extension, LEGACY_PIE_LABEL_DISTANCE_FIELD).len(),
            1
        );
        assert_eq!(raw_field(extension, PIE_LABEL_DISTANCE_FIELD).len(), 1);
        assert_eq!(
            patch_series_non_style_label_distance(&patched, ChartPieLabelDistance::DEFAULT)
                .unwrap(),
            original
        );
    }

    #[test]
    fn malformed_or_inconsistent_native_label_distances_are_rejected() {
        for invalid in [29.9, 200.1, f32::INFINITY, f32::NAN] {
            let generated = tsch::generated::ChartSeriesNonStyleArchive {
                tschchartseriespielabelexplosion: Some(invalid),
                ..Default::default()
            };
            assert!(
                read_series_non_style_label_distance(&outer_with_extension(
                    &generated.encode_to_vec()
                ))
                .is_err()
            );
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespie2_3labelexplosion: Some(100.0),
            tschchartseriespielabelexplosion: Some(150.0),
            ..Default::default()
        };
        assert!(
            read_series_non_style_label_distance(&outer_with_extension(&generated.encode_to_vec()))
                .is_err()
        );
    }

    fn outer_with_extension(extension: &[u8]) -> Vec<u8> {
        patch_length_delimited_field(
            &canonical_empty_chart_series_non_style_data().unwrap(),
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(extension),
        )
        .unwrap()
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
