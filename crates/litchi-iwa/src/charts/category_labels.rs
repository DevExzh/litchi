//! Lossless category-label frequency CRUD for native chart axes.
//!
//! Pages, Numbers, and Keynote expose the same category-label menu. Visibility
//! lives in the category-axis non-style, while interval and last-label
//! behavior live in the generated category-axis style extension.

use prost::Message;

pub use litchi_iwa_common::chart::category_labels::{Frequency, Interval, Layout};

use crate::charts::Axis;
use crate::charts::axis::{chart_axis_labels_visible, set_chart_axis_labels_visible};
use crate::charts::axis_style::{
    GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD, axis_style_slot, generated_axis_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{parse_wire_fields, patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const CATEGORY_LABEL_INTERVAL_FIELD: u32 = 5;
const CATEGORY_SHOW_LAST_LABEL_FIELD: u32 = 26;
const AUTO_FIT_INTERVAL_RAW: u64 = 0;
const SHOW_ALL_INTERVAL_RAW: u64 = 1;

/// Read the effective category-label layout for one native chart.
pub(crate) fn chart_category_label_layout(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Layout> {
    let visible = chart_axis_labels_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
    )?;
    let (stored_frequency, show_last_category) = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
    )?
    .read(package, read_category_label_style)?;
    Ok(Layout::new(
        if visible {
            stored_frequency
        } else {
            Frequency::None
        },
        show_last_category,
    ))
}

/// Set the complete category-label layout for one native chart.
pub(crate) fn set_chart_category_label_layout(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    layout: Layout,
) -> Result<()> {
    if chart_category_label_layout(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? == layout
    {
        return Ok(());
    }

    let style_slot = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
    )?;
    let (stored_frequency, stored_show_last) =
        style_slot.read(package, read_category_label_style)?;
    let requested_frequency = match layout.frequency() {
        Frequency::None => None,
        frequency => Some(frequency),
    };
    if stored_show_last != layout.show_last_category()
        || requested_frequency.is_some_and(|frequency| frequency != stored_frequency)
    {
        style_slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        style_slot.update(package, |data| {
            patch_category_label_style(data, requested_frequency, layout.show_last_category())
        })?;
    }

    set_chart_axis_labels_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        Axis::Category,
        layout.frequency() != Frequency::None,
    )?;
    if chart_category_label_layout(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != layout
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} category-label layout update failed validation"
        )));
    }
    Ok(())
}

fn read_category_label_style(data: &[u8]) -> Result<(Frequency, bool)> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok((Frequency::AutoFit, Layout::default().show_last_category()));
    };
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let frequency = match strict_optional_varint(extension, CATEGORY_LABEL_INTERVAL_FIELD)? {
        None | Some(AUTO_FIT_INTERVAL_RAW) => Frequency::AutoFit,
        Some(SHOW_ALL_INTERVAL_RAW) => Frequency::All,
        Some(raw) => Frequency::from_native(native_i32(raw)?),
    };
    let show_last_category =
        match strict_optional_varint(extension, CATEGORY_SHOW_LAST_LABEL_FIELD)? {
            None | Some(1) => true,
            Some(0) => false,
            Some(raw) => {
                return Err(Error::InvalidFormat(format!(
                    "native chart show-last-category switch must be 0 or 1, found {raw}"
                )));
            },
        };
    Ok((frequency, show_last_category))
}

fn patch_category_label_style(
    data: &[u8],
    frequency: Option<Frequency>,
    show_last_category: bool,
) -> Result<Vec<u8>> {
    let existing_extension = generated_axis_style_extension(data)?;
    let extension = existing_extension.unwrap_or_default();
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;

    let interval_present =
        strict_optional_varint(extension, CATEGORY_LABEL_INTERVAL_FIELD)?.is_some();
    let show_last_present =
        strict_optional_varint(extension, CATEGORY_SHOW_LAST_LABEL_FIELD)?.is_some();
    let mut patched_extension = if let Some(frequency) = frequency {
        let replacement = match frequency {
            Frequency::None => {
                return Err(Error::InvalidFormat(
                    "hidden category labels do not have a stored interval".to_owned(),
                ));
            },
            Frequency::AutoFit => None,
            Frequency::All => Some(SHOW_ALL_INTERVAL_RAW),
            Frequency::Every(interval) => Some(u64::from(interval.value())),
            Frequency::Unsupported(value) => Some(i64::from(value) as u64),
            _ => {
                return Err(Error::InvalidFormat(
                    "unsupported chart category-label frequency".to_owned(),
                ));
            },
        };
        patch_varint_field(
            extension,
            CATEGORY_LABEL_INTERVAL_FIELD,
            interval_present,
            replacement,
        )?
    } else {
        extension.to_vec()
    };

    let current_show_last_present =
        strict_optional_varint(&patched_extension, CATEGORY_SHOW_LAST_LABEL_FIELD)?.is_some();
    let show_last_replacement = if show_last_category {
        show_last_present.then_some(1)
    } else {
        Some(0)
    };
    patched_extension = patch_varint_field(
        &patched_extension,
        CATEGORY_SHOW_LAST_LABEL_FIELD,
        current_show_last_present,
        show_last_replacement,
    )?;

    if existing_extension.is_none() && patched_extension.is_empty() {
        return Ok(data.to_vec());
    }
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
        existing_extension.is_some(),
        (!patched_extension.is_empty()).then_some(patched_extension.as_slice()),
    )?;
    let (actual_frequency, actual_show_last) = read_category_label_style(&patched)?;
    if frequency.is_some_and(|expected| actual_frequency != expected)
        || actual_show_last != show_last_category
    {
        return Err(Error::InvalidFormat(
            "category-label style wire patch failed validation".to_owned(),
        ));
    }
    Ok(patched)
}

fn native_i32(raw: u64) -> Result<i32> {
    const NEGATIVE_INT32_START: u64 = 0xffff_ffff_8000_0000;
    if raw <= i32::MAX as u64 {
        return Ok(raw as i32);
    }
    if raw >= NEGATIVE_INT32_START {
        return Ok(raw as i64 as i32);
    }
    Err(Error::InvalidFormat(format!(
        "native chart category-label interval {raw} is outside the int32 range"
    )))
}

fn strict_optional_varint(data: &[u8], field_number: u32) -> Result<Option<u64>> {
    let fields = parse_wire_fields(data)?;
    let mut matches = fields.iter().filter(|field| field.number() == field_number);
    let Some(field) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "singular chart category-label field {field_number} occurs more than once"
        )));
    }
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "chart category-label field {field_number} is not a varint"
        )));
    }
    let (value, consumed) = litchi_iwa_common::varint::decode_varint_from_bytes(
        &data[field.payload_start()..field.end()],
    )
    .map_err(|error| Error::InvalidFormat(format!("invalid category-label value: {error}")))?;
    if field.payload_start() + consumed != field.end() {
        return Err(Error::InvalidFormat(
            "chart category-label varint has trailing bytes".to_owned(),
        ));
    }
    Ok(Some(value))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::wire::{append_length_delimited_field, append_varint_field};
    use litchi_iwa_common::chart::category_labels::Interval;

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn category_label_styles_round_trip_and_restore_exactly() {
        let original = style_with_unknown_fields();
        assert_eq!(
            read_category_label_style(&original).unwrap(),
            (Frequency::AutoFit, true)
        );
        let customized = patch_category_label_style(
            &original,
            Some(Frequency::Every(Interval::new(3).unwrap())),
            false,
        )
        .unwrap();
        assert_eq!(
            read_category_label_style(&customized).unwrap(),
            (Frequency::Every(Interval::new(3).unwrap()), false)
        );
        let restored =
            patch_category_label_style(&customized, Some(Frequency::AutoFit), true).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn unknown_native_intervals_round_trip_without_normalization() {
        let original = style_with_unknown_fields();
        let unknown =
            patch_category_label_style(&original, Some(Frequency::Unsupported(-7)), true).unwrap();
        assert_eq!(
            read_category_label_style(&unknown).unwrap(),
            (Frequency::Unsupported(-7), true)
        );

        let restored =
            patch_category_label_style(&unknown, Some(Frequency::AutoFit), true).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn malformed_category_label_styles_are_rejected() {
        let original = style_with_unknown_fields();
        let extension = generated_axis_style_extension(&original).unwrap().unwrap();

        let mut duplicate = patch_varint_field(
            extension,
            CATEGORY_LABEL_INTERVAL_FIELD,
            false,
            Some(SHOW_ALL_INTERVAL_RAW),
        )
        .unwrap();
        append_varint_field(
            &mut duplicate,
            CATEGORY_LABEL_INTERVAL_FIELD,
            SHOW_ALL_INTERVAL_RAW,
        )
        .unwrap();
        let duplicate = replace_extension(&original, duplicate);
        assert!(read_category_label_style(&duplicate).is_err());

        let invalid_interval = patch_varint_field(
            extension,
            CATEGORY_LABEL_INTERVAL_FIELD,
            false,
            Some(u64::from(i32::MAX as u32) + 1),
        )
        .unwrap();
        let invalid_interval = replace_extension(&original, invalid_interval);
        assert!(read_category_label_style(&invalid_interval).is_err());

        let invalid_boolean =
            patch_varint_field(extension, CATEGORY_SHOW_LAST_LABEL_FIELD, true, Some(2)).unwrap();
        let invalid_boolean = replace_extension(&original, invalid_boolean);
        assert!(read_category_label_style(&invalid_boolean).is_err());
    }

    fn style_with_unknown_fields() -> Vec<u8> {
        let mut generated = tsch::generated::ChartAxisStyleArchive {
            tschchartaxiscategoryshowlastlabel: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 91).unwrap();
        let mut outer = tsch::ChartAxisStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut outer,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        append_varint_field(&mut outer, UNKNOWN_OUTER_FIELD, 73).unwrap();
        outer
    }

    fn replace_extension(original: &[u8], extension: Vec<u8>) -> Vec<u8> {
        patch_length_delimited_field(
            original,
            GENERATED_CHART_AXIS_STYLE_EXTENSION_FIELD,
            true,
            Some(extension.as_slice()),
        )
        .unwrap()
    }
}
