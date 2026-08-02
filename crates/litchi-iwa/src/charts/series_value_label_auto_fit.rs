//! Lossless per-series value-label Auto-Fit CRUD for native charts.
//!
//! The inspector describes Auto-Fit as preventing overlapping labels. Native
//! archives store the inverse `show labels in front` switch in each effective
//! series style: a missing or false switch means Auto-Fit is enabled.

use prost::Message;

use crate::charts::series_style::{
    GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD, chart_series_style_slots,
    generated_chart_series_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const SHOW_LABELS_IN_FRONT_FIELD: u32 = 100;

/// Whether iWork automatically repositions one series' labels to avoid overlap.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ChartSeriesValueLabelAutoFit {
    /// Let iWork reposition value labels that would overlap.
    #[default]
    Enabled,
    /// Keep labels in front of their data points even when they overlap.
    Disabled,
}

impl ChartSeriesValueLabelAutoFit {
    /// Whether automatic overlap prevention is enabled.
    pub const fn is_enabled(self) -> bool {
        matches!(self, Self::Enabled)
    }

    const fn show_labels_in_front(self) -> bool {
        matches!(self, Self::Disabled)
    }
}

/// Read Auto-Fit settings in native series order.
pub(crate) fn chart_series_value_label_auto_fits(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    series_count: usize,
) -> Result<Vec<ChartSeriesValueLabelAutoFit>> {
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

/// Set Auto-Fit settings in native series order.
pub(crate) fn set_chart_series_value_label_auto_fits(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    expected: &[ChartSeriesValueLabelAutoFit],
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
    if chart_series_value_label_auto_fits(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        expected.len(),
    )? != expected
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} value-label Auto-Fit update failed validation"
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

fn read_auto_fit(data: &[u8]) -> Result<ChartSeriesValueLabelAutoFit> {
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(ChartSeriesValueLabelAutoFit::Enabled);
    };
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    Ok(
        if strict_optional_bool(extension, SHOW_LABELS_IN_FRONT_FIELD)?.unwrap_or(false) {
            ChartSeriesValueLabelAutoFit::Disabled
        } else {
            ChartSeriesValueLabelAutoFit::Enabled
        },
    )
}

fn patch_auto_fit(data: &[u8], expected: ChartSeriesValueLabelAutoFit) -> Result<Vec<u8>> {
    let existing_extension = generated_chart_series_style_extension(data)?;
    if existing_extension.is_none() && expected.is_enabled() {
        return Ok(data.to_vec());
    }
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let field_present = strict_optional_bool(extension, SHOW_LABELS_IN_FRONT_FIELD)?.is_some();
    let replacement = expected.show_labels_in_front().then_some(1);
    let extension = patch_varint_field(
        extension,
        SHOW_LABELS_IN_FRONT_FIELD,
        field_present,
        replacement,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    if read_auto_fit(&patched)? != expected {
        return Err(Error::InvalidFormat(
            "chart series value-label Auto-Fit wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn strict_optional_bool(data: &[u8], field_number: u32) -> Result<Option<bool>> {
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
    let (value, consumed) = crate::varint::decode_varint_from_bytes(
        &data[field.key_end..field.end],
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "chart series style field {field_number} is invalid: {error}"
        ))
    })?;
    if consumed != 1 || consumed != field.end - field.key_end || value > 1 {
        return Err(Error::InvalidFormat(format!(
            "chart series style field {field_number} is not a canonical boolean"
        )));
    }
    Ok(Some(value == 1))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};
    use prost::Message;

    fn empty_series_style() -> Vec<u8> {
        tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec()
    }

    #[test]
    fn enabled_is_the_minimal_native_default() {
        let original = empty_series_style();
        assert_eq!(
            read_auto_fit(&original).unwrap(),
            ChartSeriesValueLabelAutoFit::Enabled
        );
        assert_eq!(
            patch_auto_fit(&original, ChartSeriesValueLabelAutoFit::Enabled).unwrap(),
            original
        );
    }

    #[test]
    fn disabled_round_trips_and_resets_exactly() {
        let original = empty_series_style();
        let disabled = patch_auto_fit(&original, ChartSeriesValueLabelAutoFit::Disabled).unwrap();
        assert_eq!(
            read_auto_fit(&disabled).unwrap(),
            ChartSeriesValueLabelAutoFit::Disabled
        );
        assert_eq!(
            patch_auto_fit(&disabled, ChartSeriesValueLabelAutoFit::Enabled).unwrap(),
            original
        );
    }

    #[test]
    fn patch_preserves_neighboring_and_unknown_fields() {
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

        let patched = patch_auto_fit(&original, ChartSeriesValueLabelAutoFit::Disabled).unwrap();
        let extension = generated_chart_series_style_extension(&patched)
            .unwrap()
            .unwrap();
        assert_eq!(
            strict_optional_bool(extension, SHOW_LABELS_IN_FRONT_FIELD).unwrap(),
            Some(true)
        );
        let unknown = parse_wire_fields(extension)
            .unwrap()
            .into_iter()
            .find(|field| field.number == UNKNOWN_FIELD)
            .unwrap();
        assert_eq!(&extension[unknown.key_end..unknown.end], &[73]);
    }

    #[test]
    fn malformed_native_boolean_is_rejected() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, SHOW_LABELS_IN_FRONT_FIELD, 2).unwrap();
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
