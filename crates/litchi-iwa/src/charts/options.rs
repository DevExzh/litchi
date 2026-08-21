//! Lossless native chart-option storage and mutation.
//!
//! iWork chart title and legend controls live in the generated extension of a
//! chart's `TSCH.ChartNonStyleArchive`. This module keeps both the outer
//! non-style payload and the generated extension lossless while updating only
//! their requested native option fields.

use prost::Message;

use crate::charts::non_style::{
    GENERATED_CHART_NON_STYLE_EXTENSION_FIELD, chart_non_style_slot,
    generated_chart_non_style_extension,
};
use crate::protobuf::tsch;
use crate::wire::{patch_length_delimited_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};
use litchi_iwa_protos::keynote_chart_title_codec::{
    ChartTitleWrite, DecodeOptions, decode_chart_title, decode_visible_chart_title,
    rewrite_chart_title,
};

/// `tschchartinfodefaultshowlegend` in `TSCH.Generated.ChartNonStyleArchive`.
const CHART_LEGEND_VISIBLE_FIELD: u32 = 20;
/// Private hard ceiling for one native chart title's UTF-8 bytes.
const MAX_CHART_TITLE_BYTES: usize = 64 * 1024 * 1024;
/// Read one chart title from its native chart non-style extension.
pub(crate) fn chart_title(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<Option<String>> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_non_style_title)
}

/// Set one chart title, enabling the native title switch when necessary.
pub(crate) fn set_chart_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    title: &str,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_title)?.as_deref() == Some(title) {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_non_style_title(data, Some(title))
    })?;
    if slot.read(package, read_chart_non_style_title)?.as_deref() != Some(title) {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} title update failed validation"
        )));
    }
    Ok(())
}

/// Remove one visible chart title.
///
/// Returns whether the title was visible. A chart non-style shared by more than
/// one chart is rejected rather than silently changing another chart.
pub(crate) fn remove_chart_title(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_title)?.is_none() {
        return Ok(false);
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_non_style_title(data, None))?;
    if slot.read(package, read_chart_non_style_title)?.is_some() {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} title removal failed validation"
        )));
    }
    Ok(true)
}

/// Read whether one chart shows its native legend.
pub(crate) fn chart_legend_visible(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<bool> {
    chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_non_style_legend_visible)
}

/// Set whether one chart shows its native legend.
pub(crate) fn set_chart_legend_visible(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    visible: bool,
) -> Result<()> {
    let slot = chart_non_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    if slot.read(package, read_chart_non_style_legend_visible)? == visible {
        return Ok(());
    }
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| {
        patch_chart_non_style_legend_visibility(data, visible)
    })?;
    if slot.read(package, read_chart_non_style_legend_visible)? != visible {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend update failed validation"
        )));
    }
    Ok(())
}

/// Decode a `TSCH.ChartNonStyleArchive` and return its visible native title.
pub(crate) fn read_chart_non_style_title(data: &[u8]) -> Result<Option<String>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(None);
    };
    decode_visible_chart_title(extension, chart_title_decode_options(extension, None))
        .map(|title| title.map(str::to_owned))
        .map_err(chart_title_codec_error)
}

/// Decode a `TSCH.ChartNonStyleArchive` and return its native legend switch.
pub(crate) fn read_chart_non_style_legend_visible(data: &[u8]) -> Result<bool> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        return Ok(false);
    };
    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    Ok(generated.tschchartinfodefaultshowlegend.unwrap_or(false))
}

fn patch_chart_non_style_title(data: &[u8], title: Option<&str>) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        let Some(title) = title else {
            return Ok(data.to_vec());
        };
        let extension = rewrite_chart_title(
            &[],
            ChartTitleWrite::new(Some(true), Some(title)),
            chart_title_decode_options(&[], Some(title)),
        )
        .map_err(chart_title_codec_error)?;
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_title(&patched, Some(title))?;
        return Ok(patched);
    };

    if title.is_none()
        && decode_chart_title(extension, chart_title_decode_options(extension, None))
            .map_err(chart_title_codec_error)?
            .title_visible()
            != Some(true)
    {
        return Ok(data.to_vec());
    }

    let extension = rewrite_chart_title(
        extension,
        ChartTitleWrite::new(Some(title.is_some()), title),
        chart_title_decode_options(extension, title),
    )
    .map_err(chart_title_codec_error)?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_title(&patched, title)?;
    Ok(patched)
}

/// Build a finite policy for the strict generated title projection.
///
/// The codec validates the whole selected extension, while this module owns
/// the outer length-delimited field. A replacement title may be larger than
/// the source extension, so the rewrite budget is based on both payloads;
/// the source byte ceiling remains exact for the first decode and is widened
/// by the codec itself only for its readback of the candidate output.
fn chart_title_decode_options(source: &[u8], replacement: Option<&str>) -> DecodeOptions {
    let title_bytes = replacement.map_or(0, str::len);
    let output_bytes = source
        .len()
        .saturating_add(title_bytes)
        .saturating_add(32)
        .max(1);
    let source_bytes = source.len().max(1);
    DecodeOptions::new(
        source_bytes,
        output_bytes.saturating_mul(8).max(1),
        output_bytes.saturating_mul(16).max(1),
        8,
    )
    .with_max_output_bytes(output_bytes)
    .with_max_title_bytes(chart_title_limit(source.len(), title_bytes))
}

fn chart_title_limit(source_bytes: usize, replacement_bytes: usize) -> usize {
    source_bytes
        .max(replacement_bytes)
        .clamp(1, MAX_CHART_TITLE_BYTES)
}

fn chart_title_codec_error(
    error: litchi_iwa_protos::keynote_chart_title_codec::DecodeError,
) -> Error {
    Error::InvalidFormat(format!("invalid chart title generated extension: {error}"))
}

fn patch_chart_non_style_legend_visibility(data: &[u8], visible: bool) -> Result<Vec<u8>> {
    let Some(extension) = generated_chart_non_style_extension(data)? else {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(visible),
            ..Default::default()
        };
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            false,
            Some(generated.encode_to_vec().as_slice()),
        )?;
        validate_patched_legend_visibility(&patched, visible)?;
        return Ok(patched);
    };

    let generated = tsch::generated::ChartNonStyleArchive::decode(extension)?;
    let visible_present = generated.tschchartinfodefaultshowlegend.is_some();
    let extension = patch_varint_field(
        extension,
        CHART_LEGEND_VISIBLE_FIELD,
        visible_present,
        Some(u64::from(visible)),
    )?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
        true,
        Some(extension.as_slice()),
    )?;
    validate_patched_legend_visibility(&patched, visible)?;
    Ok(patched)
}

fn validate_patched_title(data: &[u8], expected: Option<&str>) -> Result<()> {
    if read_chart_non_style_title(data)?.as_deref() != expected {
        return Err(Error::InvalidFormat(
            "chart non-style title wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

fn validate_patched_legend_visibility(data: &[u8], expected: bool) -> Result<()> {
    if read_chart_non_style_legend_visible(data)? != expected {
        return Err(Error::InvalidFormat(
            "chart non-style legend wire patch failed validation".to_owned(),
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
    fn title_patch_retains_unmapped_chart_non_style_fields() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(false),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let titled = patch_chart_non_style_title(&original, Some("Revenue by region")).unwrap();
        assert_eq!(
            read_chart_non_style_title(&titled).unwrap(),
            Some("Revenue by region".to_owned())
        );
        assert_eq!(
            raw_field(&titled, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&titled)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );

        let removed = patch_chart_non_style_title(&titled, None).unwrap();
        assert_eq!(read_chart_non_style_title(&removed).unwrap(), None);
        assert_eq!(removed, original);
    }

    #[test]
    fn title_patch_preserves_hidden_stale_text_until_explicitly_replaced() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowtitle: Some(false),
            tschchartinfodefaulttitle: Some("stale title".to_owned()),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        assert_eq!(read_chart_non_style_title(&original).unwrap(), None);
        assert_eq!(
            patch_chart_non_style_title(&original, None).unwrap(),
            original
        );

        let visible = patch_chart_non_style_title(&original, Some("fresh title")).unwrap();
        assert_eq!(
            read_chart_non_style_title(&visible).unwrap(),
            Some("fresh title".to_owned())
        );
        assert_eq!(
            raw_field(&visible, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&visible)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );
    }

    #[test]
    fn title_patch_preserves_absent_visibility_with_stale_text() {
        let mut extension = Vec::new();
        append_length_delimited_field(&mut extension, 23, b"stale title").unwrap();
        let original = outer_chart_non_style(&extension);

        assert_eq!(read_chart_non_style_title(&original).unwrap(), None);
        assert_eq!(
            patch_chart_non_style_title(&original, None).unwrap(),
            original
        );

        let visible = patch_chart_non_style_title(&original, Some("fresh title")).unwrap();
        assert_eq!(
            read_chart_non_style_title(&visible).unwrap(),
            Some("fresh title".to_owned())
        );
    }

    #[test]
    fn title_patch_preserves_visible_empty_default_when_text_is_absent() {
        let mut extension = Vec::new();
        append_varint_field(&mut extension, 21, 1).unwrap();
        let original = outer_chart_non_style(&extension);

        assert_eq!(
            read_chart_non_style_title(&original).unwrap(),
            Some(String::new())
        );
        let titled = patch_chart_non_style_title(&original, Some("fresh title")).unwrap();
        assert_eq!(
            read_chart_non_style_title(&titled).unwrap(),
            Some("fresh title".to_owned())
        );

        let removed = patch_chart_non_style_title(&original, None).unwrap();
        assert_eq!(read_chart_non_style_title(&removed).unwrap(), None);
    }

    #[test]
    fn title_codec_rejects_duplicate_selected_fields_before_rewrite() {
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut extension = Vec::new();
        append_varint_field(&mut extension, 21, 1).unwrap();
        append_varint_field(&mut extension, 21, 0).unwrap();
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();

        let error = read_chart_non_style_title(&original).expect_err("duplicate title switch");
        assert!(error.to_string().contains("duplicate singular field"));
        assert!(patch_chart_non_style_title(&original, Some("new")).is_err());
    }

    #[test]
    fn title_decode_policy_caps_source_and_replacement_at_64_mib() {
        assert_eq!(
            chart_title_limit(0, MAX_CHART_TITLE_BYTES),
            MAX_CHART_TITLE_BYTES
        );
        assert_eq!(
            chart_title_limit(0, MAX_CHART_TITLE_BYTES + 1),
            MAX_CHART_TITLE_BYTES
        );
        assert_eq!(
            chart_title_limit(MAX_CHART_TITLE_BYTES + 1, 0),
            MAX_CHART_TITLE_BYTES
        );
    }

    #[test]
    fn title_patch_rejects_replacement_over_64_mib_before_output_reservation() {
        let oversized = "x".repeat(MAX_CHART_TITLE_BYTES + 1);
        let error = patch_chart_non_style_title(&[], Some(&oversized))
            .expect_err("oversized chart title must fail closed");
        assert!(error.to_string().contains("maximum is 67108864"));
    }

    #[test]
    fn legend_patch_retains_title_and_unmapped_chart_non_style_fields() {
        let generated = tsch::generated::ChartNonStyleArchive {
            tschchartinfodefaultshowlegend: Some(true),
            tschchartinfodefaultshowtitle: Some(true),
            tschchartinfodefaulttitle: Some("Revenue by region".to_owned()),
            ..Default::default()
        };
        let mut extension = generated.encode_to_vec();
        append_varint_field(&mut extension, UNMAPPED_GENERATED_FIELD, UNMAPPED_VALUE).unwrap();
        let base = tsch::ChartNonStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        };
        let mut original = base.encode_to_vec();
        append_length_delimited_field(
            &mut original,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            &extension,
        )
        .unwrap();
        append_varint_field(&mut original, UNMAPPED_OUTER_FIELD, UNMAPPED_VALUE).unwrap();

        let hidden = patch_chart_non_style_legend_visibility(&original, false).unwrap();
        assert!(!read_chart_non_style_legend_visible(&hidden).unwrap());
        assert_eq!(
            read_chart_non_style_title(&hidden).unwrap(),
            Some("Revenue by region".to_owned())
        );
        assert_eq!(
            raw_field(&hidden, UNMAPPED_OUTER_FIELD),
            raw_field(&original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_non_style_extension(&hidden)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            ),
            raw_field(
                generated_chart_non_style_extension(&original)
                    .unwrap()
                    .unwrap(),
                UNMAPPED_GENERATED_FIELD
            )
        );

        let visible = patch_chart_non_style_legend_visibility(&hidden, true).unwrap();
        assert!(read_chart_non_style_legend_visible(&visible).unwrap());
        assert_eq!(visible, original);
    }

    fn raw_field(data: &[u8], number: u32) -> Vec<Vec<u8>> {
        parse_wire_fields(data)
            .unwrap()
            .into_iter()
            .filter(|field| field.number() == number)
            .map(|field| data[field.start()..field.end()].to_vec())
            .collect()
    }

    fn outer_chart_non_style(extension: &[u8]) -> Vec<u8> {
        let mut data = tsch::ChartNonStyleArchive::default().encode_to_vec();
        append_length_delimited_field(
            &mut data,
            GENERATED_CHART_NON_STYLE_EXTENSION_FIELD,
            extension,
        )
        .unwrap();
        data
    }
}
