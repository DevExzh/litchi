//! Lossless native chart-shadow storage and mutation.
//!
//! The chart inspector distinguishes shadows drawn for each series from one
//! shadow drawn around a grouped series set. iWork stores the drop-shadow
//! appearance and that grouping switch as separate generated chart-style
//! fields. This module combines them into one strict public value while
//! preserving every unrelated protobuf byte.

use prost::Message;

use crate::charts::series_style::{
    GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD, chart_series_style_slots,
    generated_chart_series_style_extension,
};
use crate::charts::style::{
    GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot, generated_chart_style_extension,
};
use crate::protobuf::tsch;
use litchi_iwa_common::shape::shadow::{Angle, Opacity};
use crate::shapes::{
    Appearance, BlurRadius, Drop, Offset, RgbaColor, Shadow, shadow_from_native, shadow_to_native,
};
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

/// `tschchartinfodefaultcombinelayers` in `TSCH.Generated.ChartStyleArchive`.
const CHART_COMBINE_LAYERS_FIELD: u32 = 13;
/// `tschchartseriesdefaultshadow` in `TSCH.Generated.ChartSeriesStyleArchive`.
const CHART_SERIES_SHADOW_FIELD: u32 = 38;
/// Native chart-shadow scope and appearance.
///
/// Charts support only drop shadows. Contact and curved shadows belong to the
/// ordinary shape inspector and are deliberately excluded here.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ChartShadow {
    /// Render no chart shadow.
    None,
    /// Draw the shadow independently for every series.
    IndividualSeries(Drop),
    /// Draw one shadow around each grouped series set.
    Grouped(Drop),
}

impl ChartShadow {
    /// Shadow shown for a newly inserted native chart.
    pub const NATIVE_DEFAULT: Self = Self::IndividualSeries(Drop::new(
        Appearance::new(
            RgbaColor::black(),
            BlurRadius::TEN_POINTS,
            Offset::SIX_POINTS,
            Opacity::THREE_QUARTERS,
        ),
        Angle::FORTY_FIVE_DEGREES,
    ));

    /// Construct the shadow shown for a newly inserted native chart.
    pub const fn native_default() -> Self {
        Self::NATIVE_DEFAULT
    }

    /// Return the enabled drop-shadow appearance, if any.
    pub const fn drop_shadow(self) -> Option<Drop> {
        match self {
            Self::None => None,
            Self::IndividualSeries(shadow) | Self::Grouped(shadow) => Some(shadow),
        }
    }

    /// Whether iWork combines series layers before drawing the shadow.
    pub const fn groups_series(self) -> bool {
        matches!(self, Self::Grouped(_))
    }
}

impl Default for ChartShadow {
    fn default() -> Self {
        Self::NATIVE_DEFAULT
    }
}

/// Read the effective shadow of one native chart.
pub(crate) fn chart_shadow(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartShadow> {
    let chart_slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let combined = chart_slot.read(package, read_chart_shadow_scope)?;
    let series_slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let mut shadows = series_slots
        .iter()
        .map(|slot| slot.read(package, read_chart_series_shadow));
    let shadow = shadows
        .next()
        .transpose()?
        .ok_or_else(|| Error::InvalidFormat("chart has no readable series shadows".to_owned()))?;
    for candidate in shadows {
        if candidate? != shadow {
            return Err(Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} has nonuniform series shadows"
            )));
        }
    }
    chart_shadow_from_native(shadow, combined)
}

/// Set the shadow of one native chart.
pub(crate) fn set_chart_shadow(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    shadow: ChartShadow,
) -> Result<()> {
    if chart_shadow(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? == shadow
    {
        return Ok(());
    }
    let chart_slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let series_slots = chart_series_style_slots(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    chart_slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    for slot in &series_slots {
        slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    }
    chart_slot.update(package, |data| patch_chart_shadow_scope(data, shadow))?;
    let series_shadow = shadow
        .drop_shadow()
        .map(Shadow::Drop)
        .unwrap_or(Shadow::Disabled);
    for slot in &series_slots {
        slot.update(package, |data| {
            patch_chart_series_shadow(data, series_shadow)
        })?;
    }
    if chart_shadow(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )? != shadow
    {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} shadow update failed validation"
        )));
    }
    Ok(())
}

fn read_chart_shadow_scope(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    Ok(generated.tschchartinfodefaultcombinelayers.unwrap_or(false))
}

fn read_chart_series_shadow(data: &[u8]) -> Result<Shadow> {
    let native_default = ChartShadow::native_default().drop_shadow().ok_or_else(|| {
        Error::InvalidFormat("native chart-shadow default is disabled".to_owned())
    })?;
    let Some(extension) = generated_chart_series_style_extension(data)? else {
        return Ok(Shadow::Drop(native_default));
    };
    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    generated
        .tschchartseriesdefaultshadow
        .as_ref()
        .map(shadow_from_native)
        .transpose()
        .map(|shadow| shadow.unwrap_or(Shadow::Drop(native_default)))
}

fn chart_shadow_from_native(shadow: Shadow, combined: bool) -> Result<ChartShadow> {
    match shadow {
        Shadow::Disabled if !combined => Ok(ChartShadow::None),
        Shadow::Drop(shadow) => Ok(if combined {
            ChartShadow::Grouped(shadow)
        } else {
            ChartShadow::IndividualSeries(shadow)
        }),
        Shadow::Disabled => Err(Error::InvalidFormat(
            "disabled chart shadow unexpectedly combines series layers".to_owned(),
        )),
        Shadow::Contact(_) | Shadow::Curved(_) => Err(Error::InvalidFormat(
            "native chart uses a non-drop shadow".to_owned(),
        )),
    }
}

fn patch_chart_shadow_scope(data: &[u8], shadow: ChartShadow) -> Result<Vec<u8>> {
    let native_default = ChartShadow::native_default();
    let combined = shadow.groups_series();
    let Some(extension) = generated_chart_style_extension(data)? else {
        if !combined {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultcombinelayers: Some(true),
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_shadow_scope(&patched, combined)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartStyleArchive::decode(extension)?;
    let combined_present = generated.tschchartinfodefaultcombinelayers.is_some();
    let native_combined = if shadow == native_default {
        None
    } else {
        (combined_present || combined).then_some(u64::from(combined))
    };
    let extension = patch_varint_field(
        extension,
        CHART_COMBINE_LAYERS_FIELD,
        combined_present,
        native_combined,
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_shadow_scope(&patched, combined)?;
    Ok(patched)
}

fn patch_chart_series_shadow(data: &[u8], shadow: Shadow) -> Result<Vec<u8>> {
    let native_default = ChartShadow::native_default().drop_shadow().ok_or_else(|| {
        Error::InvalidFormat("native chart-shadow default is disabled".to_owned())
    })?;
    let native = match shadow {
        Shadow::Drop(drop) if drop == native_default => None,
        Shadow::Disabled | Shadow::Drop(_) => Some(shadow_to_native(shadow)),
        Shadow::Contact(_) | Shadow::Curved(_) => {
            return Err(Error::InvalidFormat(
                "chart series cannot use a non-drop shadow".to_owned(),
            ));
        },
    };

    let Some(extension) = generated_chart_series_style_extension(data)? else {
        if native.is_none() {
            return Ok(data.to_vec());
        }
        let generated = tsch::generated::ChartSeriesStyleArchive {
            tschchartseriesdefaultshadow: native,
            ..Default::default()
        };
        let encoded = generated.encode_to_vec();
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            false,
            Some(encoded.as_slice()),
        )?;
        validate_patched_chart_series_shadow(&patched, shadow)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartSeriesStyleArchive::decode(extension)?;
    let shadow_present = generated.tschchartseriesdefaultshadow.is_some();
    let native = native.map(|shadow| shadow.encode_to_vec());
    let extension = patch_length_delimited_field(
        extension,
        CHART_SERIES_SHADOW_FIELD,
        shadow_present,
        native.as_deref(),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_chart_series_shadow(&patched, shadow)?;
    Ok(patched)
}

fn validate_patched_chart_shadow_scope(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_shadow_scope(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart shadow-scope wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn validate_patched_chart_series_shadow(data: &[u8], expected: Shadow) -> Result<()> {
    if read_chart_series_shadow(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart series-shadow wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::RgbColorSpace;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn chart_shadow_defaults_natively_and_creates_an_extension_when_needed() {
        let chart_style = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let series_style = tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let native_default = ChartShadow::native_default().drop_shadow().unwrap();

        assert!(!read_chart_shadow_scope(&chart_style).unwrap());
        assert_eq!(
            read_chart_series_shadow(&series_style).unwrap(),
            Shadow::Drop(native_default)
        );
        assert_eq!(
            patch_chart_shadow_scope(&chart_style, ChartShadow::native_default()).unwrap(),
            chart_style
        );
        assert_eq!(
            patch_chart_series_shadow(&series_style, Shadow::Drop(native_default)).unwrap(),
            series_style
        );

        let none = patch_chart_series_shadow(&series_style, Shadow::Disabled).unwrap();
        assert_eq!(read_chart_series_shadow(&none).unwrap(), Shadow::Disabled);
        let grouped =
            patch_chart_shadow_scope(&chart_style, ChartShadow::Grouped(custom_shadow())).unwrap();
        let customized =
            patch_chart_series_shadow(&series_style, Shadow::Drop(custom_shadow())).unwrap();
        assert!(read_chart_shadow_scope(&grouped).unwrap());
        assert_eq!(
            read_chart_series_shadow(&customized).unwrap(),
            Shadow::Drop(custom_shadow())
        );
    }

    #[test]
    fn chart_shadow_patch_retains_other_style_fields_and_unmapped_data() {
        let original_shadow = custom_shadow();
        let replacement_shadow = Drop::new(
            original_shadow
                .appearance()
                .with_opacity(Opacity::new(0.42).unwrap()),
            Angle::from_degrees(135.0).unwrap(),
        );
        let original_chart = chart_style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultcombinelayers: Some(true),
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(25.0),
            ..Default::default()
        });
        let original_series =
            series_style_with_unknown_fields(tsch::generated::ChartSeriesStyleArchive {
                tschchartseriesdefaultshadow: Some(shadow_to_native(Shadow::Drop(original_shadow))),
                tschchartseriesdefaultopacity: Some(0.8),
                ..Default::default()
            });

        let patched_chart = patch_chart_shadow_scope(
            &original_chart,
            ChartShadow::IndividualSeries(replacement_shadow),
        )
        .unwrap();
        let patched_series =
            patch_chart_series_shadow(&original_series, Shadow::Drop(replacement_shadow)).unwrap();
        assert!(!read_chart_shadow_scope(&patched_chart).unwrap());
        assert_eq!(
            read_chart_series_shadow(&patched_series).unwrap(),
            Shadow::Drop(replacement_shadow)
        );
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&patched_chart)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultinterbargap, Some(25.0));
        let generated_series = tsch::generated::ChartSeriesStyleArchive::decode(
            generated_chart_series_style_extension(&patched_series)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated_series.tschchartseriesdefaultopacity, Some(0.8));
        assert_chart_unknown_fields_retained(&original_chart, &patched_chart);
        assert_series_unknown_fields_retained(&original_series, &patched_series);

        let restored_chart =
            patch_chart_shadow_scope(&patched_chart, ChartShadow::Grouped(original_shadow))
                .unwrap();
        let restored_series =
            patch_chart_series_shadow(&patched_series, Shadow::Drop(original_shadow)).unwrap();
        assert_eq!(restored_chart, original_chart);
        assert_eq!(restored_series, original_series);
    }

    #[test]
    fn resetting_chart_shadow_retains_other_style_fields() {
        let original_chart = chart_style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultcombinelayers: Some(true),
            tschchartinfodefaultshowborder: Some(true),
            ..Default::default()
        });
        let original_series =
            series_style_with_unknown_fields(tsch::generated::ChartSeriesStyleArchive {
                tschchartseriesdefaultshadow: Some(shadow_to_native(Shadow::Drop(custom_shadow()))),
                tschchartseriesdefaultopacity: Some(0.8),
                ..Default::default()
            });
        let default_shadow = ChartShadow::native_default().drop_shadow().unwrap();

        let reset_chart =
            patch_chart_shadow_scope(&original_chart, ChartShadow::native_default()).unwrap();
        let reset_series =
            patch_chart_series_shadow(&original_series, Shadow::Drop(default_shadow)).unwrap();
        assert!(!read_chart_shadow_scope(&reset_chart).unwrap());
        assert_eq!(
            read_chart_series_shadow(&reset_series).unwrap(),
            Shadow::Drop(default_shadow)
        );
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&reset_chart)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultcombinelayers, None);
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        let generated_series = tsch::generated::ChartSeriesStyleArchive::decode(
            generated_chart_series_style_extension(&reset_series)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(generated_series.tschchartseriesdefaultshadow, None);
        assert_eq!(generated_series.tschchartseriesdefaultopacity, Some(0.8));
        assert_chart_unknown_fields_retained(&original_chart, &reset_chart);
        assert_series_unknown_fields_retained(&original_series, &reset_series);
    }

    #[test]
    fn malformed_native_chart_shadows_are_rejected() {
        let contact = series_style_with_unknown_fields(tsch::generated::ChartSeriesStyleArchive {
            tschchartseriesdefaultshadow: Some(shadow_to_native(Shadow::Contact(
                crate::shapes::Contact::new(
                    custom_shadow().appearance(),
                    crate::shapes::Perspective::LEVEL,
                ),
            ))),
            ..Default::default()
        });
        assert!(
            chart_shadow_from_native(read_chart_series_shadow(&contact).unwrap(), false).is_err()
        );

        assert!(
            chart_shadow_from_native(Shadow::Disabled, true).is_err(),
            "disabled grouped shadows must remain invalid"
        );
    }

    fn custom_shadow() -> Drop {
        Drop::new(
            Appearance::new(
                RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb).unwrap(),
                BlurRadius::from_points(15).unwrap(),
                Offset::from_points(8.0).unwrap(),
                Opacity::new(0.6).unwrap(),
            ),
            Angle::from_degrees(60.0).unwrap(),
        )
    }

    fn chart_style_with_unknown_fields(generated: tsch::generated::ChartStyleArchive) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut data = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(&mut data, GENERATED_CHART_STYLE_EXTENSION_FIELD, &extension)
            .unwrap();
        append_varint_field(&mut data, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        data
    }

    fn series_style_with_unknown_fields(
        generated: tsch::generated::ChartSeriesStyleArchive,
    ) -> Vec<u8> {
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let mut data = tsch::ChartSeriesStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_SERIES_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut data, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();
        data
    }

    fn assert_chart_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        assert_unknown_fields_retained(original, patched, generated_chart_style_extension);
    }

    fn assert_series_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        assert_unknown_fields_retained(original, patched, generated_chart_series_style_extension);
    }

    fn assert_unknown_fields_retained<'a>(
        original: &'a [u8],
        patched: &'a [u8],
        extension: impl Fn(&'a [u8]) -> Result<Option<&'a [u8]>>,
    ) {
        assert_eq!(
            raw_field(patched, UNMAPPED_OUTER_FIELD),
            raw_field(original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                extension(patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                extension(original).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            )
        );
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
