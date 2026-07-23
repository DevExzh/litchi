//! Lossless native value-axis scale-bound storage and mutation.
//!
//! iWork stores the manual minimum and maximum of a chart's primary value
//! axis in the generated extension of `TSCH.ChartAxisNonStyleArchive`. This
//! module represents automatic versus fixed bounds explicitly, preserves both
//! protobuf layers losslessly, and patches only those two nested messages.

use prost::Message;

use crate::charts::ChartAxis;
use crate::charts::axis::{
    GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD, axis_non_style_slot,
    generated_axis_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkPackage, Result};

/// `tschchartaxisdefaultusermax` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MAXIMUM_FIELD: u32 = 17;
/// `tschchartaxisdefaultusermin` in `TSCH.Generated.ChartAxisNonStyleArchive`.
const VALUE_AXIS_MINIMUM_FIELD: u32 = 18;

/// One finite manual bound for an iWork chart value axis.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartAxisBound(f64);

impl ChartAxisBound {
    /// Create a finite native chart-axis bound.
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::InvalidFormat(
                "chart axis bound must be finite".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the numeric value used by iWork's Axis Scale inspector.
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for ChartAxisBound {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::new(value)
    }
}

/// The manual bounds for a native chart's primary value axis.
///
/// A missing bound is iWork's `Auto` value. Construct this type through
/// [`Self::new`] to reject inverted ranges, or use [`Self::automatic`] for the
/// default automatic scale.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartValueAxisBounds {
    minimum: Option<ChartAxisBound>,
    maximum: Option<ChartAxisBound>,
}

impl ChartValueAxisBounds {
    /// Build a value-axis range from optional manual endpoints.
    pub fn new(minimum: Option<ChartAxisBound>, maximum: Option<ChartAxisBound>) -> Result<Self> {
        let bounds = Self { minimum, maximum };
        bounds.validate()?;
        Ok(bounds)
    }

    /// Use iWork's automatic lower and upper value-axis bounds.
    pub const fn automatic() -> Self {
        Self {
            minimum: None,
            maximum: None,
        }
    }

    /// Build a fully manual value-axis range.
    pub fn fixed(minimum: ChartAxisBound, maximum: ChartAxisBound) -> Result<Self> {
        Self::new(Some(minimum), Some(maximum))
    }

    /// Return the optional manual lower bound.
    pub const fn minimum(self) -> Option<ChartAxisBound> {
        self.minimum
    }

    /// Return the optional manual upper bound.
    pub const fn maximum(self) -> Option<ChartAxisBound> {
        self.maximum
    }

    fn validate(self) -> Result<()> {
        if let (Some(minimum), Some(maximum)) = (self.minimum, self.maximum)
            && minimum.value() > maximum.value()
        {
            return Err(Error::InvalidFormat(
                "chart value-axis minimum exceeds maximum".to_owned(),
            ));
        }
        Ok(())
    }
}

/// Read the manual bounds of one native chart's primary value axis.
pub(crate) fn chart_value_axis_bounds(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartValueAxisBounds> {
    axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?
    .read(package, read_value_axis_bounds)
}

/// Set the manual bounds of one native chart's primary value axis.
pub(crate) fn set_chart_value_axis_bounds(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    bounds: ChartValueAxisBounds,
) -> Result<()> {
    let slot = axis_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Value,
    )?;
    if slot.read(package, read_value_axis_bounds)? == bounds {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_value_axis_bounds(data, bounds))?;
    if slot.read(package, read_value_axis_bounds)? != bounds {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-axis bounds update failed validation"
        )));
    }
    Ok(())
}

fn read_value_axis_bounds(data: &[u8]) -> Result<ChartValueAxisBounds> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        return Ok(ChartValueAxisBounds::automatic());
    };
    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    ChartValueAxisBounds::new(
        decode_bound(generated.tschchartaxisdefaultusermin.as_ref(), "minimum")?,
        decode_bound(generated.tschchartaxisdefaultusermax.as_ref(), "maximum")?,
    )
}

fn decode_bound(
    archive: Option<&tsch::ChartsNsNumberDoubleArchive>,
    label: &str,
) -> Result<Option<ChartAxisBound>> {
    let Some(archive) = archive else {
        return Ok(None);
    };
    let value = archive.number_archive.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "chart value-axis {label} is missing its numeric value"
        ))
    })?;
    ChartAxisBound::new(value).map(Some)
}

fn encoded_bound(bound: Option<ChartAxisBound>) -> Option<Vec<u8>> {
    bound.map(|bound| {
        tsch::ChartsNsNumberDoubleArchive {
            number_archive: Some(bound.value()),
        }
        .encode_to_vec()
    })
}

fn patch_value_axis_bounds(data: &[u8], bounds: ChartValueAxisBounds) -> Result<Vec<u8>> {
    let Some(extension) = generated_axis_non_style_extension(data)? else {
        if bounds == ChartValueAxisBounds::automatic() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartAxisNonStyleArchive {
            tschchartaxisdefaultusermin: bounds.minimum().map(number_archive),
            tschchartaxisdefaultusermax: bounds.maximum().map(number_archive),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_value_axis_bounds(&patched, bounds)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartAxisNonStyleArchive::decode(extension)?;
    let minimum = encoded_bound(bounds.minimum());
    let maximum = encoded_bound(bounds.maximum());
    let extension = patch_length_delimited_field(
        extension,
        VALUE_AXIS_MINIMUM_FIELD,
        generated.tschchartaxisdefaultusermin.is_some(),
        minimum.as_deref(),
    )?;
    let extension = patch_length_delimited_field(
        &extension,
        VALUE_AXIS_MAXIMUM_FIELD,
        generated.tschchartaxisdefaultusermax.is_some(),
        maximum.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_value_axis_bounds(&patched, bounds)?;
    Ok(patched)
}

fn number_archive(bound: ChartAxisBound) -> tsch::ChartsNsNumberDoubleArchive {
    tsch::ChartsNsNumberDoubleArchive {
        number_archive: Some(bound.value()),
    }
}

fn validate_patched_value_axis_bounds(data: &[u8], expected: ChartValueAxisBounds) -> Result<()> {
    if read_value_axis_bounds(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart value-axis bounds wire patch failed validation".to_owned(),
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
    fn bounds_reject_nonfinite_and_inverted_ranges() {
        assert!(ChartAxisBound::new(f64::NAN).is_err());
        assert!(ChartAxisBound::new(f64::INFINITY).is_err());

        let low = ChartAxisBound::new(-1.0).unwrap();
        let high = ChartAxisBound::new(1.0).unwrap();
        assert!(ChartValueAxisBounds::fixed(high, low).is_err());
        assert_eq!(
            ChartValueAxisBounds::fixed(low, high).unwrap().minimum(),
            Some(low)
        );
    }

    #[test]
    fn value_axis_bounds_patch_retains_other_fields_and_unmapped_data() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxiscategoryshowtitle: Some(true),
                tschchartaxiscategorytitle: Some("Month".to_owned()),
                tschchartaxisdefaultusermin: Some(number_archive(
                    ChartAxisBound::new(-5.0).unwrap(),
                )),
                tschchartaxisdefaultusermax: Some(number_archive(
                    ChartAxisBound::new(50.0).unwrap(),
                )),
                ..Default::default()
            });
        let replacement = ChartValueAxisBounds::fixed(
            ChartAxisBound::new(0.0).unwrap(),
            ChartAxisBound::new(30.0).unwrap(),
        )
        .unwrap();

        let patched = patch_value_axis_bounds(&original, replacement).unwrap();
        assert_eq!(read_value_axis_bounds(&patched).unwrap(), replacement);
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(
            generated_axis_non_style_extension(&patched)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            generated.tschchartaxiscategorytitle.as_deref(),
            Some("Month")
        );
        assert_unknown_fields_retained(&original, &patched);

        let restored = patch_value_axis_bounds(
            &patched,
            ChartValueAxisBounds::fixed(
                ChartAxisBound::new(-5.0).unwrap(),
                ChartAxisBound::new(50.0).unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn automatic_bounds_remove_only_the_requested_fields() {
        let original =
            axis_non_style_with_unknown_fields(tsch::generated::ChartAxisNonStyleArchive {
                tschchartaxisdefaultusermin: Some(number_archive(
                    ChartAxisBound::new(0.0).unwrap(),
                )),
                tschchartaxisdefaultusermax: Some(number_archive(
                    ChartAxisBound::new(30.0).unwrap(),
                )),
                tschchartaxisvalueshowtitle: Some(true),
                tschchartaxisvaluetitle: Some("Revenue".to_owned()),
                ..Default::default()
            });

        let automatic =
            patch_value_axis_bounds(&original, ChartValueAxisBounds::automatic()).unwrap();
        assert_eq!(
            read_value_axis_bounds(&automatic).unwrap(),
            ChartValueAxisBounds::automatic()
        );
        let generated = tsch::generated::ChartAxisNonStyleArchive::decode(
            generated_axis_non_style_extension(&automatic)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            generated.tschchartaxisvaluetitle.as_deref(),
            Some("Revenue")
        );
        assert_unknown_fields_retained(&original, &automatic);
    }

    #[test]
    fn value_axis_bounds_patch_creates_an_extension_when_missing() {
        let original = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let bounds =
            ChartValueAxisBounds::new(Some(ChartAxisBound::new(-2.0).unwrap()), None).unwrap();

        let patched = patch_value_axis_bounds(&original, bounds).unwrap();
        assert_eq!(read_value_axis_bounds(&patched).unwrap(), bounds);
        assert_eq!(
            patch_value_axis_bounds(&original, ChartValueAxisBounds::automatic()).unwrap(),
            original
        );
    }

    fn axis_non_style_with_unknown_fields(
        generated: tsch::generated::ChartAxisNonStyleArchive,
    ) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartAxisNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut data = base.encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_AXIS_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
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
                generated_axis_non_style_extension(patched)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_axis_non_style_extension(original)
                    .unwrap()
                    .unwrap(),
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
