//! Lossless leader-line visibility CRUD for native pie and donut charts.

use prost::Message;

use crate::charts::series_non_style::{
    NewChartSeriesNonStyleBase, chart_series_non_style_values,
    generated_chart_series_non_style_extension, patch_chart_series_non_style_extension,
    set_chart_series_non_style_values,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartseriespieenablecalloutline` in the generated series non-style.
const PIE_LEADER_LINE_VISIBILITY_FIELD: u32 = 102;
const NATIVE_HIDDEN: u64 = 0;
const NATIVE_VISIBLE: u64 = 2;

/// Whether iWork draws a leader line between a pie label and its wedge.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ChartPieLeaderLineVisibility {
    Hidden,
    /// The native default.
    #[default]
    Visible,
}

impl ChartPieLeaderLineVisibility {
    fn from_native(value: u64) -> Result<Self> {
        match value {
            NATIVE_HIDDEN => Ok(Self::Hidden),
            NATIVE_VISIBLE => Ok(Self::Visible),
            value => Err(Error::InvalidFormat(format!(
                "unsupported native pie leader-line visibility {value}"
            ))),
        }
    }

    const fn native(self) -> u64 {
        match self {
            Self::Hidden => NATIVE_HIDDEN,
            Self::Visible => NATIVE_VISIBLE,
        }
    }
}

/// Read leader-line visibility for every wedge in chart-series order.
pub(crate) fn chart_pie_leader_line_visibilities(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartPieLeaderLineVisibility>> {
    chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
        ChartPieLeaderLineVisibility::Visible,
        read_series_non_style_leader_line_visibility,
    )
}

/// Set leader-line visibility for every wedge in chart-series order.
pub(crate) fn set_chart_pie_leader_line_visibilities(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartPieLeaderLineVisibility],
) -> Result<()> {
    set_chart_series_non_style_values(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        "pie leader-line visibility",
        NewChartSeriesNonStyleBase::Unstyled,
        expected,
        ChartPieLeaderLineVisibility::Visible,
        read_series_non_style_leader_line_visibility,
        |data, visibility| patch_series_non_style_leader_line_visibility(data, *visibility),
    )
}

fn read_series_non_style_leader_line_visibility(
    data: &[u8],
) -> Result<ChartPieLeaderLineVisibility> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        return Ok(ChartPieLeaderLineVisibility::Visible);
    };
    tsch::generated::ChartSeriesNonStyleArchive::decode(extension)?;
    Ok(strict_optional_visibility(extension)?.unwrap_or(ChartPieLeaderLineVisibility::Visible))
}

fn patch_series_non_style_leader_line_visibility(
    data: &[u8],
    visibility: ChartPieLeaderLineVisibility,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_series_non_style_extension(data)? else {
        if visibility == ChartPieLeaderLineVisibility::Visible {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesNonStyleArchive {
            tschchartseriespieenablecalloutline: Some(visibility.native() as i32),
            ..Default::default()
        };
        let patched = patch_chart_series_non_style_extension(
            data,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_visibility(&patched, visibility)?;
        return Ok(patched);
    };

    let visibility_present = strict_optional_visibility(extension)?.is_some();
    let replacement =
        (visibility != ChartPieLeaderLineVisibility::Visible).then_some(visibility.native());
    let extension = patch_varint_field(
        extension,
        PIE_LEADER_LINE_VISIBILITY_FIELD,
        visibility_present,
        replacement,
    )?;
    let patched = patch_chart_series_non_style_extension(
        data,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_visibility(&patched, visibility)?;
    Ok(patched)
}

fn strict_optional_visibility(data: &[u8]) -> Result<Option<ChartPieLeaderLineVisibility>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields
        .iter()
        .filter(|field| field.number == PIE_LEADER_LINE_VISIBILITY_FIELD);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(
            "singular chart pie leader-line visibility occurs more than once".to_owned(),
        ));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(
            "chart pie leader-line visibility is not a varint".to_owned(),
        ));
    }
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart pie leader-line visibility is invalid: {error}"
        ))
    })?;
    if field.key_end + consumed != field.end {
        return Err(Error::InvalidFormat(
            "chart pie leader-line visibility is not canonical".to_owned(),
        ));
    }
    ChartPieLeaderLineVisibility::from_native(value).map(Some)
}

fn validate_patched_visibility(data: &[u8], expected: ChartPieLeaderLineVisibility) -> Result<()> {
    if read_series_non_style_leader_line_visibility(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart pie leader-line visibility wire patch failed validation".to_owned(),
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
    fn leader_line_visibility_matches_native_values() {
        assert_eq!(
            ChartPieLeaderLineVisibility::default(),
            ChartPieLeaderLineVisibility::Visible
        );
        assert_eq!(
            ChartPieLeaderLineVisibility::from_native(NATIVE_HIDDEN).unwrap(),
            ChartPieLeaderLineVisibility::Hidden
        );
        assert_eq!(
            ChartPieLeaderLineVisibility::from_native(NATIVE_VISIBLE).unwrap(),
            ChartPieLeaderLineVisibility::Visible
        );
        assert!(ChartPieLeaderLineVisibility::from_native(1).is_err());
        assert!(ChartPieLeaderLineVisibility::from_native(3).is_err());
    }

    #[test]
    fn leader_line_visibility_patch_is_lossless_and_resets_exactly() {
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

        assert_eq!(
            read_series_non_style_leader_line_visibility(&original).unwrap(),
            ChartPieLeaderLineVisibility::Visible
        );
        let hidden = patch_series_non_style_leader_line_visibility(
            &original,
            ChartPieLeaderLineVisibility::Hidden,
        )
        .unwrap();
        assert_eq!(
            read_series_non_style_leader_line_visibility(&hidden).unwrap(),
            ChartPieLeaderLineVisibility::Hidden
        );
        let hidden_outer = parse_wire_fields(&hidden).unwrap();
        assert!(
            hidden_outer
                .iter()
                .any(|field| field.number == UNMAPPED_OUTER_FIELD)
        );
        let hidden_extension = generated_chart_series_non_style_extension(&hidden)
            .unwrap()
            .unwrap();
        let hidden_generated = parse_wire_fields(hidden_extension).unwrap();
        assert!(
            hidden_generated
                .iter()
                .any(|field| field.number == UNMAPPED_GENERATED_FIELD)
        );

        let restored = patch_series_non_style_leader_line_visibility(
            &hidden,
            ChartPieLeaderLineVisibility::Visible,
        )
        .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn malformed_leader_line_visibility_is_rejected() {
        for value in [1, 3] {
            let mut generated = Vec::new();
            append_varint_field(&mut generated, PIE_LEADER_LINE_VISIBILITY_FIELD, value).unwrap();
            let mut data = canonical_empty_chart_series_non_style_data().unwrap();
            append_length_delimited_field(
                &mut data,
                GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
                &generated,
            )
            .unwrap();
            assert!(read_series_non_style_leader_line_visibility(&data).is_err());
        }

        let mut generated = Vec::new();
        append_varint_field(
            &mut generated,
            PIE_LEADER_LINE_VISIBILITY_FIELD,
            NATIVE_HIDDEN,
        )
        .unwrap();
        append_varint_field(
            &mut generated,
            PIE_LEADER_LINE_VISIBILITY_FIELD,
            NATIVE_VISIBLE,
        )
        .unwrap();
        let mut data = canonical_empty_chart_series_non_style_data().unwrap();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        assert!(read_series_non_style_leader_line_visibility(&data).is_err());

        let malformed =
            patch_length_delimited_field(&[], PIE_LEADER_LINE_VISIBILITY_FIELD, false, Some(&[0]))
                .unwrap();
        let mut data = canonical_empty_chart_series_non_style_data().unwrap();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_SERIES_NON_STYLE_EXTENSION_FIELD,
            &malformed,
        )
        .unwrap();
        assert!(read_series_non_style_leader_line_visibility(&data).is_err());
    }
}
