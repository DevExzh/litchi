//! Lossless category-label frequency CRUD for native chart axes.
//!
//! Pages, Numbers, and Keynote expose the same category-label menu. Visibility
//! lives in the category-axis non-style, while interval and last-label
//! behavior live in the generated category-axis style extension.

use prost::Message;

use crate::charts::ChartAxis;
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
const MINIMUM_CUSTOM_INTERVAL: u32 = 2;
const MAXIMUM_CUSTOM_INTERVAL: u32 = i32::MAX as u32;

/// A validated interval used by iWork's `Custom Category Intervals` option.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartCategoryLabelInterval(u32);

impl ChartCategoryLabelInterval {
    /// Construct a custom interval.
    ///
    /// Native values `0` and `1` are reserved for Auto-Fit and Show All, so
    /// custom intervals begin at two and must fit the signed protobuf field.
    pub fn new(interval: u32) -> Result<Self> {
        if !(MINIMUM_CUSTOM_INTERVAL..=MAXIMUM_CUSTOM_INTERVAL).contains(&interval) {
            return Err(Error::InvalidFormat(format!(
                "chart category-label interval must be in {MINIMUM_CUSTOM_INTERVAL}..={MAXIMUM_CUSTOM_INTERVAL}"
            )));
        }
        Ok(Self(interval))
    }

    /// Return the number of categories between displayed labels.
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl TryFrom<u32> for ChartCategoryLabelInterval {
    type Error = Error;

    fn try_from(value: u32) -> Result<Self> {
        Self::new(value)
    }
}

/// How frequently a native chart displays category labels.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum ChartCategoryLabelFrequency {
    /// Hide ordinary category labels.
    None,
    /// Let iWork choose a readable interval for the available chart width.
    #[default]
    AutoFit,
    /// Display every category label.
    All,
    /// Display labels at one explicit category interval.
    Every(ChartCategoryLabelInterval),
}

/// Complete native category-label menu state.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartCategoryLabelLayout {
    frequency: ChartCategoryLabelFrequency,
    show_last_category: bool,
}

impl ChartCategoryLabelLayout {
    /// Construct category-label layout settings.
    pub const fn new(frequency: ChartCategoryLabelFrequency, show_last_category: bool) -> Self {
        Self {
            frequency,
            show_last_category,
        }
    }

    /// Return how frequently category labels are displayed.
    pub const fn frequency(self) -> ChartCategoryLabelFrequency {
        self.frequency
    }

    /// Return whether iWork forces the final category label to appear.
    pub const fn show_last_category(self) -> bool {
        self.show_last_category
    }
}

impl Default for ChartCategoryLabelLayout {
    fn default() -> Self {
        Self::new(ChartCategoryLabelFrequency::AutoFit, true)
    }
}

/// Read the effective category-label layout for one native chart.
pub(crate) fn chart_category_label_layout(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartCategoryLabelLayout> {
    let visible = chart_axis_labels_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Category,
    )?;
    let (stored_frequency, show_last_category) = axis_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Category,
    )?
    .read(package, read_category_label_style)?;
    Ok(ChartCategoryLabelLayout::new(
        if visible {
            stored_frequency
        } else {
            ChartCategoryLabelFrequency::None
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
    layout: ChartCategoryLabelLayout,
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
        ChartAxis::Category,
    )?;
    let (stored_frequency, stored_show_last) =
        style_slot.read(package, read_category_label_style)?;
    let requested_frequency = match layout.frequency {
        ChartCategoryLabelFrequency::None => None,
        frequency => Some(frequency),
    };
    if stored_show_last != layout.show_last_category
        || requested_frequency.is_some_and(|frequency| frequency != stored_frequency)
    {
        style_slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
        style_slot.update(package, |data| {
            patch_category_label_style(data, requested_frequency, layout.show_last_category)
        })?;
    }

    set_chart_axis_labels_visible(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        ChartAxis::Category,
        layout.frequency != ChartCategoryLabelFrequency::None,
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

fn read_category_label_style(data: &[u8]) -> Result<(ChartCategoryLabelFrequency, bool)> {
    let Some(extension) = generated_axis_style_extension(data)? else {
        return Ok((
            ChartCategoryLabelFrequency::AutoFit,
            ChartCategoryLabelLayout::default().show_last_category,
        ));
    };
    tsch::generated::ChartAxisStyleArchive::decode(extension)?;
    let frequency = match strict_optional_varint(extension, CATEGORY_LABEL_INTERVAL_FIELD)? {
        None | Some(AUTO_FIT_INTERVAL_RAW) => ChartCategoryLabelFrequency::AutoFit,
        Some(SHOW_ALL_INTERVAL_RAW) => ChartCategoryLabelFrequency::All,
        Some(raw) => {
            let interval = u32::try_from(raw).map_err(|_| {
                Error::InvalidFormat(format!(
                    "native chart category-label interval {raw} exceeds u32"
                ))
            })?;
            ChartCategoryLabelFrequency::Every(ChartCategoryLabelInterval::new(interval).map_err(
                |_| {
                    Error::InvalidFormat(format!(
                        "unsupported native chart category-label interval {raw}"
                    ))
                },
            )?)
        },
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
    frequency: Option<ChartCategoryLabelFrequency>,
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
            ChartCategoryLabelFrequency::None => {
                return Err(Error::InvalidFormat(
                    "hidden category labels do not have a stored interval".to_owned(),
                ));
            },
            ChartCategoryLabelFrequency::AutoFit => None,
            ChartCategoryLabelFrequency::All => Some(SHOW_ALL_INTERVAL_RAW),
            ChartCategoryLabelFrequency::Every(interval) => Some(u64::from(interval.get())),
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
    let (value, consumed) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.payload_start()..field.end()])
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid category-label value: {error}"))
            })?;
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

    const UNKNOWN_OUTER_FIELD: u32 = 4_096;
    const UNKNOWN_GENERATED_FIELD: u32 = 4_097;

    #[test]
    fn custom_intervals_are_strict() {
        assert!(ChartCategoryLabelInterval::new(0).is_err());
        assert!(ChartCategoryLabelInterval::new(1).is_err());
        assert_eq!(ChartCategoryLabelInterval::new(2).unwrap().get(), 2);
        assert_eq!(
            ChartCategoryLabelInterval::new(MAXIMUM_CUSTOM_INTERVAL)
                .unwrap()
                .get(),
            MAXIMUM_CUSTOM_INTERVAL
        );
        assert!(ChartCategoryLabelInterval::new(MAXIMUM_CUSTOM_INTERVAL + 1).is_err());
    }

    #[test]
    fn category_label_styles_round_trip_and_restore_exactly() {
        let original = style_with_unknown_fields();
        assert_eq!(
            read_category_label_style(&original).unwrap(),
            (ChartCategoryLabelFrequency::AutoFit, true)
        );
        let customized = patch_category_label_style(
            &original,
            Some(ChartCategoryLabelFrequency::Every(
                ChartCategoryLabelInterval::new(3).unwrap(),
            )),
            false,
        )
        .unwrap();
        assert_eq!(
            read_category_label_style(&customized).unwrap(),
            (
                ChartCategoryLabelFrequency::Every(ChartCategoryLabelInterval::new(3).unwrap()),
                false
            )
        );
        let restored = patch_category_label_style(
            &customized,
            Some(ChartCategoryLabelFrequency::AutoFit),
            true,
        )
        .unwrap();
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
            Some(u64::from(MAXIMUM_CUSTOM_INTERVAL) + 1),
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
