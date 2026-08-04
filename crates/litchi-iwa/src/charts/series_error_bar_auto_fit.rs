//! Lossless per-series error-bar Auto-Fit CRUD for native charts.
//!
//! A missing spacing override lets iWork prevent overlapping bars. The native
//! fixed-spacing override disables that behavior.

use crate::charts::series_style::{
    GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD, chart_series_style_slots,
    generated_chart_series_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};
use prost::Message;

const ERROR_BAR_SPACING_FIELD: u32 = 98;
const NATIVE_FIXED_SPACING: i32 = 1;

/// Whether iWork automatically repositions one series' error bars to avoid overlap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesErrorBarAutoFit {
    /// Let iWork reposition error bars that would overlap.
    #[default]
    Enabled,
    /// Keep fixed error-bar spacing even when bars overlap.
    Disabled,
}

impl ChartSeriesErrorBarAutoFit {
    /// Whether automatic overlap prevention is enabled.
    pub const fn is_enabled(self) -> bool {
        matches!(self, Self::Enabled)
    }
}

pub(crate) fn chart_series_error_bar_auto_fits(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartSeriesErrorBarAutoFit>> {
    let slots = effective_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        series_count,
    )?;
    slots
        .iter()
        .map(|slot| slot.read(package, read_auto_fit))
        .collect()
}

pub(crate) fn set_chart_series_error_bar_auto_fits(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartSeriesErrorBarAutoFit],
) -> Result<()> {
    let slots = effective_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )?;
    let current = slots
        .iter()
        .map(|slot| slot.read(package, read_auto_fit))
        .collect::<Result<Vec<_>>>()?;
    if current == expected {
        return Ok(());
    }
    for (slot, (current, replacement)) in slots.iter().zip(current.iter().zip(expected)) {
        if current == replacement {
            continue;
        }
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        slot.update(package, |data| patch_auto_fit(data, *replacement))?;
    }
    if chart_series_error_bar_auto_fits(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} error-bar Auto-Fit update failed validation"
        )));
    }
    Ok(())
}

fn effective_slots(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<crate::charts::series_style::ChartSeriesStyleSlot>> {
    let mut slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slots.len() < series_count {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} has {} series styles for {series_count} series",
            slots.len()
        )));
    }
    slots.truncate(series_count);
    Ok(slots)
}

fn read_auto_fit(data: &[u8]) -> Result<ChartSeriesErrorBarAutoFit> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(ChartSeriesErrorBarAutoFit::Enabled);
    };
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    match strict_optional_i32(extension, ERROR_BAR_SPACING_FIELD)? {
        None | Some(0) => Ok(ChartSeriesErrorBarAutoFit::Enabled),
        Some(NATIVE_FIXED_SPACING) => Ok(ChartSeriesErrorBarAutoFit::Disabled),
        Some(value) => Err(Error::InvalidFormat(format!(
            "unsupported native chart error-bar spacing {value}"
        ))),
    }
}

fn patch_auto_fit(data: &[u8], expected: ChartSeriesErrorBarAutoFit) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_style_extension(data)?;
    if existing_extension.is_none() && expected.is_enabled() {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let present = strict_optional_i32(extension, ERROR_BAR_SPACING_FIELD)?.is_some();
    let extension = patch_varint_field(
        extension,
        ERROR_BAR_SPACING_FIELD,
        present,
        matches!(expected, ChartSeriesErrorBarAutoFit::Disabled)
            .then_some(NATIVE_FIXED_SPACING as u64),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_auto_fit(&patched)? != expected {
        return Err(Error::InvalidFormat(
            "chart series error-bar Auto-Fit wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn strict_optional_i32(data: &[u8], field_number: u32) -> Result<Option<i32>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart series style field {field_number} occurs more than once"
        )));
    }
    if field.wire_type != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart series style field {field_number} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "chart series style field {field_number} is invalid: {error}"
                ))
            })?;
    let decoded = value as i32;
    let encoded = decoded as i64 as u64;
    if consumed != field.end - field.key_end
        || litchi_iwa_common::varint::encoded_len(value) != consumed
        || (encoded != value && value > i32::MAX as u64)
    {
        return Err(Error::InvalidFormat(format!(
            "chart series style field {field_number} is not a canonical int32"
        )));
    }
    Ok(Some(decoded))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};

    fn empty_series_style() -> Vec<u8> {
        tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec()
    }

    #[test]
    fn enabled_is_minimal_and_disabled_resets_exactly() {
        let original = empty_series_style();
        assert_eq!(
            read_auto_fit(&original).unwrap(),
            ChartSeriesErrorBarAutoFit::Enabled
        );
        assert_eq!(
            patch_auto_fit(&original, ChartSeriesErrorBarAutoFit::Enabled).unwrap(),
            original
        );
        let disabled = patch_auto_fit(&original, ChartSeriesErrorBarAutoFit::Disabled).unwrap();
        assert_eq!(
            read_auto_fit(&disabled).unwrap(),
            ChartSeriesErrorBarAutoFit::Disabled
        );
        assert_eq!(
            patch_auto_fit(&disabled, ChartSeriesErrorBarAutoFit::Enabled).unwrap(),
            original
        );
    }

    #[test]
    fn patch_preserves_neighboring_style_fields() {
        const UNKNOWN_FIELD: u32 = 9_001;
        let mut extension = Vec::new();
        append_varint_field(&mut extension, 88, 6).unwrap();
        append_varint_field(&mut extension, UNKNOWN_FIELD, 73).unwrap();
        let mut original = empty_series_style();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        let patched = patch_auto_fit(&original, ChartSeriesErrorBarAutoFit::Disabled).unwrap();
        let extension = generated_chart_series_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_i32(extension, ERROR_BAR_SPACING_FIELD).unwrap(),
            Some(NATIVE_FIXED_SPACING)
        );
        assert_eq!(
            strict_optional_i32(extension, UNKNOWN_FIELD).unwrap(),
            Some(73)
        );
    }

    #[test]
    fn malformed_or_unknown_spacing_is_rejected() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, ERROR_BAR_SPACING_FIELD, 2).unwrap();
        let mut data = empty_series_style();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        assert!(read_auto_fit(&data).is_err());
    }
}
