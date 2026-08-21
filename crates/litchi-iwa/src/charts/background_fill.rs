//! Lossless native chart-background fill storage and mutation.
//!
//! The chart inspector's background modes map to one `TSD.FillArchive` in the
//! generated chart-style extension. This module handles ordinary color,
//! gradient, and image fills, including media insertion and component
//! data-reference accounting, while preserving unrelated protobuf bytes.

use prost::Message;

use litchi_iwa_common::media::Type as MediaType;
use litchi_iwa_common::shape::fill::{Angle, Gradient};

use crate::charts::style::{
    ChartStyleSlot, GENERATED_CHART_STYLE_EXTENSION_FIELD, chart_style_slot,
    generated_chart_style_extension, replace_known_wire_fields_preserving_unknown,
};
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::media::MediaAssetId;
use crate::package_metadata::component_identifier_for_entry;
#[cfg(test)]
use crate::protobuf::tsch;
use crate::shapes::{
    DrawableSize, RgbColorSpace, RgbaColor, ShapeFill, ShapeImageDataIdentifier, ShapeImageFill,
    ShapeImageFillTechnique, fill_from_native, fill_to_native, image_data_identifier,
    remove_orphaned_image_asset, validate_image_asset,
};
use crate::wire::patch_length_delimited_field;
use crate::{Error, IWorkMediaEditor, IWorkPackage, Result};

mod codec {
    use super::CHART_BACKGROUND_FILL_FIELD;
    use crate::wire::{parse_wire_fields, patch_length_delimited_field};
    use crate::{Error, Result};

    /// Borrowed projection of the selected generated chart-style field.
    ///
    /// The surrounding generated chart-style extension remains caller-owned;
    /// unknown fields and their original wire representation are never
    /// materialized or normalized here.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    pub(super) struct BackgroundFillProjection<'source> {
        payload: Option<&'source [u8]>,
    }

    impl<'source> BackgroundFillProjection<'source> {
        pub(super) const fn payload(self) -> Option<&'source [u8]> {
            self.payload
        }
    }

    /// Project only `tschchartinfodefaultgridbackgroundfill` from one
    /// `TSCH.Generated.ChartStyleArchive` extension payload.
    pub(super) fn project<'source>(
        source: &'source [u8],
    ) -> Result<BackgroundFillProjection<'source>> {
        let fields = parse_wire_fields(source)?;
        let mut payload = None;
        for field in fields {
            if field.number() != CHART_BACKGROUND_FILL_FIELD {
                continue;
            }
            if payload.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "chart style background-fill field {CHART_BACKGROUND_FILL_FIELD} occurs more than once"
                )));
            }
            if field.wire_type() != 2 {
                return Err(Error::InvalidFormat(format!(
                    "chart style background-fill field {CHART_BACKGROUND_FILL_FIELD} is not length-delimited"
                )));
            }
            field.validate_canonical_framing(source)?;
            payload = Some(field.checked_payload(source)?);
        }
        Ok(BackgroundFillProjection { payload })
    }

    /// Rewrite only field 14, retaining every other field byte-for-byte.
    pub(super) fn rewrite(source: &[u8], replacement: Option<&[u8]>) -> Result<Vec<u8>> {
        let current = project(source)?;
        let output = patch_length_delimited_field(
            source,
            CHART_BACKGROUND_FILL_FIELD,
            current.payload.is_some(),
            replacement,
        )?;
        let projected = project(&output)?;
        if projected.payload != replacement {
            return Err(Error::InvalidFormat(
                "chart style background-fill wire rewrite failed validation".to_owned(),
            ));
        }
        Ok(output)
    }

    #[cfg(test)]
    mod tests {
        use super::*;
        use crate::wire::{append_length_delimited_field, append_varint_field};

        const UNKNOWN_BEFORE_FIELD: u32 = 4_096;
        const UNKNOWN_AFTER_FIELD: u32 = 4_097;

        #[test]
        fn projection_distinguishes_absent_and_present_empty_payloads() {
            let mut source = Vec::new();
            append_varint_field(&mut source, UNKNOWN_BEFORE_FIELD, 7).unwrap();

            assert_eq!(project(&source).unwrap().payload(), None);

            let present = rewrite(&source, Some(&[])).unwrap();
            assert_eq!(project(&present).unwrap().payload(), Some(&[] as &[u8]));

            let removed = rewrite(&present, None).unwrap();
            assert_eq!(removed, source);
            assert_eq!(project(&removed).unwrap().payload(), None);
        }

        #[test]
        fn rewrite_preserves_unknown_fields_and_selected_presence() {
            let mut source = Vec::new();
            append_varint_field(&mut source, UNKNOWN_BEFORE_FIELD, 11).unwrap();
            append_length_delimited_field(&mut source, CHART_BACKGROUND_FILL_FIELD, b"old")
                .unwrap();
            append_varint_field(&mut source, UNKNOWN_AFTER_FIELD, 13).unwrap();

            let replacement = rewrite(&source, Some(b"new")).unwrap();
            assert_eq!(
                project(&replacement).unwrap().payload(),
                Some(b"new".as_slice())
            );
            assert_eq!(
                raw_fields(&replacement, UNKNOWN_BEFORE_FIELD),
                raw_fields(&source, UNKNOWN_BEFORE_FIELD)
            );
            assert_eq!(
                raw_fields(&replacement, UNKNOWN_AFTER_FIELD),
                raw_fields(&source, UNKNOWN_AFTER_FIELD)
            );

            let cleared = rewrite(&replacement, None).unwrap();
            assert_eq!(project(&cleared).unwrap().payload(), None);
            assert_eq!(
                raw_fields(&cleared, UNKNOWN_BEFORE_FIELD),
                raw_fields(&source, UNKNOWN_BEFORE_FIELD)
            );
            assert_eq!(
                raw_fields(&cleared, UNKNOWN_AFTER_FIELD),
                raw_fields(&source, UNKNOWN_AFTER_FIELD)
            );
        }

        #[test]
        fn projection_rejects_duplicate_or_wrong_wire_selected_fields() {
            let mut wrong_wire = Vec::new();
            append_varint_field(&mut wrong_wire, CHART_BACKGROUND_FILL_FIELD, 1).unwrap();
            assert!(project(&wrong_wire).is_err());

            let mut duplicate = Vec::new();
            append_length_delimited_field(&mut duplicate, CHART_BACKGROUND_FILL_FIELD, b"one")
                .unwrap();
            append_length_delimited_field(&mut duplicate, CHART_BACKGROUND_FILL_FIELD, b"two")
                .unwrap();
            assert!(project(&duplicate).is_err());
            assert!(rewrite(&duplicate, Some(b"new")).is_err());
        }

        fn raw_fields(data: &[u8], number: u32) -> Vec<Vec<u8>> {
            parse_wire_fields(data)
                .unwrap()
                .into_iter()
                .filter(|field| field.number() == number)
                .map(|field| field.raw(data).unwrap().to_vec())
                .collect()
        }
    }
}

/// `tschchartinfodefaultgridbackgroundfill` in `TSCH.Generated.ChartStyleArchive`.
const CHART_BACKGROUND_FILL_FIELD: u32 = 14;
/// Gray channel shown as 67% by a newly inserted native chart.
const DEFAULT_BACKGROUND_GRAY_CHANNEL: f32 = 2.0 / 3.0;
/// Native inspector angle for the default top-to-bottom background gradient.
const DEFAULT_BACKGROUND_ANGLE_DEGREES: f32 = 0.0;

/// The gray-to-white gradient used when no direct background fill is stored.
pub(crate) fn native_default_chart_background_fill() -> Result<ShapeFill> {
    let gray = RgbaColor::new(
        DEFAULT_BACKGROUND_GRAY_CHANNEL,
        DEFAULT_BACKGROUND_GRAY_CHANNEL,
        DEFAULT_BACKGROUND_GRAY_CHANNEL,
        1.0,
        RgbColorSpace::Srgb,
    )?;
    let angle = Angle::from_degrees(DEFAULT_BACKGROUND_ANGLE_DEGREES)?;
    Ok(ShapeFill::Gradient(Gradient::linear(
        gray,
        RgbaColor::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb)?,
        angle,
    )))
}

/// Read the effective background fill of one native chart.
pub(crate) fn chart_background_fill(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ShapeFill> {
    chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?
    .read(package, read_chart_background_fill)
}

/// Set the background fill of one native chart.
pub(crate) fn set_chart_background_fill(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    fill: &ShapeFill,
) -> Result<()> {
    let slot = chart_style_slot(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = slot.read(package, read_chart_background_fill)?;
    if &current == fill {
        return Ok(());
    }
    validate_image_asset(package, fill)?;
    let old_data_identifier = image_data_identifier(&current)
        .map(MediaAssetId::try_from)
        .transpose()?;
    let new_data_identifier = image_data_identifier(fill)
        .map(MediaAssetId::try_from)
        .transpose()?;
    slot.ensure_exclusive(package, drawable_object_id, drawable_label)?;
    slot.update(package, |data| patch_chart_background_fill(data, fill))?;
    adjust_chart_style_data_reference(package, &slot, old_data_identifier, new_data_identifier)?;
    if slot.read(package, read_chart_background_fill)? != *fill {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} background fill update failed validation"
        )));
    }
    remove_orphaned_image_asset(package, old_data_identifier.map(MediaAssetId::get))?;
    Ok(())
}

/// Embed image bytes and use them as a simple or tinted chart background.
pub(crate) fn set_chart_background_image_fill_data(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    preferred_filename: &str,
    data: &[u8],
    technique: ShapeImageFillTechnique,
    fill_size: DrawableSize,
    tint: Option<RgbaColor>,
) -> Result<ShapeImageFill> {
    let mut media = IWorkMediaEditor::from_package(package.clone())?;
    let asset = media.insert_unreferenced(preferred_filename, data)?;
    if asset.media_type != MediaType::Image {
        return Err(Error::Bundle(format!(
            "Chart background image fills require image data, not {}",
            asset.media_type.name()
        )));
    }
    let mut staged = media.into_package();
    let mut image = ShapeImageFill::embedded(
        ShapeImageDataIdentifier::new(asset.data_identifier.get())?,
        technique,
        fill_size,
    )?;
    if let Some(tint) = tint {
        image = image.with_tint(tint);
    }
    set_chart_background_fill(
        &mut staged,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
        &ShapeFill::Image(image.clone()),
    )?;
    let package_path = asset.package_path.as_deref().ok_or_else(|| {
        Error::InvalidFormat("Chart background image-fill asset is not materialized".to_owned())
    })?;
    if staged.entry(package_path) != Some(data) {
        return Err(Error::InvalidFormat(
            "Chart background image-fill insertion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(image)
}

fn read_chart_background_fill(data: &[u8]) -> Result<ShapeFill> {
    let Some(extension) = generated_chart_style_extension(data)? else {
        return native_default_chart_background_fill();
    };
    let projection = codec::project(extension)?;
    let Some(payload) = projection.payload() else {
        return native_default_chart_background_fill();
    };
    let fill = crate::protobuf::tsd::FillArchive::decode(payload)?;
    fill_from_native(&fill)
}

fn patch_chart_background_fill(data: &[u8], fill: &ShapeFill) -> Result<Vec<u8>> {
    let default = native_default_chart_background_fill()?;
    let Some(extension) = generated_chart_style_extension(data)? else {
        if fill == &default {
            return Ok(data.to_vec());
        }
        let encoded = fill_to_native(fill).encode_to_vec();
        let extension = codec::rewrite(&[], Some(encoded.as_slice()))?;
        let patched = patch_length_delimited_field(
            data,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )?;
        validate_patched_chart_background_fill(&patched, fill)?;
        return Ok(patched);
    };

    let native = (fill != &default).then(|| fill_to_native(fill).encode_to_vec());
    let native = native
        .map(|replacement| {
            codec::project(extension).and_then(|projection| {
                projection.payload().map_or_else(
                    || Ok(replacement.clone()),
                    |existing| {
                        replace_known_wire_fields_preserving_unknown(
                            existing,
                            &replacement,
                            &[1, 2, 3],
                        )
                    },
                )
            })
        })
        .transpose()?;
    let extension = codec::rewrite(extension, native.as_deref())?;
    let patched = patch_length_delimited_field(
        data,
        GENERATED_CHART_STYLE_EXTENSION_FIELD,
        true,
        (!extension.is_empty()).then_some(extension.as_slice()),
    )?;
    validate_patched_chart_background_fill(&patched, fill)?;
    Ok(patched)
}

fn adjust_chart_style_data_reference(
    package: &mut IWorkPackage,
    slot: &ChartStyleSlot,
    old_data_identifier: Option<MediaAssetId>,
    new_data_identifier: Option<MediaAssetId>,
) -> Result<()> {
    if old_data_identifier == new_data_identifier {
        return Ok(());
    }
    let component_id = component_identifier_for_entry(package, slot.archive_name())?
        .ok_or_else(|| Error::InvalidFormat("Chart style has no owning component".to_owned()))?;
    if let Some(identifier) = old_data_identifier {
        remove_component_data_reference(package, component_id, identifier.get(), slot.object_id())?;
    }
    if let Some(identifier) = new_data_identifier {
        add_component_data_reference(package, component_id, identifier.get(), slot.object_id())?;
    }
    Ok(())
}

fn validate_patched_chart_background_fill(data: &[u8], expected: &ShapeFill) -> Result<()> {
    if &read_chart_background_fill(data)? != expected {
        return Err(Error::InvalidFormat(
            "Chart background fill wire patch failed validation".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tss;
    use crate::shapes::Pattern;
    use crate::wire::{append_length_delimited_field, append_varint_field, parse_wire_fields};

    const UNMAPPED_OUTER_FIELD: u32 = 4_096;
    const UNMAPPED_GENERATED_FIELD: u32 = 4_097;
    const UNMAPPED_SELECTED_FILL_FIELD: u32 = 4_098;
    const UNMAPPED_VALUE: u64 = 42;

    #[test]
    fn background_fill_defaults_natively_and_creates_an_extension_when_needed() {
        let original = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let replacement = solid_fill();

        assert_eq!(
            read_chart_background_fill(&original).unwrap(),
            native_default_chart_background_fill().unwrap()
        );
        assert_eq!(
            patch_chart_background_fill(
                &original,
                &native_default_chart_background_fill().unwrap()
            )
            .unwrap(),
            original
        );

        let patched = patch_chart_background_fill(&original, &replacement).unwrap();
        assert_eq!(read_chart_background_fill(&patched).unwrap(), replacement);
        let empty = patch_chart_background_fill(&patched, &ShapeFill::None).unwrap();
        assert_eq!(read_chart_background_fill(&empty).unwrap(), ShapeFill::None);
    }

    #[test]
    fn background_fill_patch_retains_other_style_fields_and_unmapped_data() {
        let original_fill = solid_fill();
        let replacement = ShapeFill::Gradient(Gradient::linear(
            color(0.8, 0.2, 0.1),
            color(0.1, 0.4, 0.9),
            Angle::from_degrees(90.0).unwrap(),
        ));
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultgridbackgroundfill: Some(fill_to_native(&original_fill)),
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultinterbargap: Some(25.0),
            tschchartinfodefaultroundedcornerradius: Some(0.2),
            ..Default::default()
        });

        let patched = patch_chart_background_fill(&original, &replacement).unwrap();
        assert_eq!(read_chart_background_fill(&patched).unwrap(), replacement);
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&patched).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert_eq!(generated.tschchartinfodefaultinterbargap, Some(25.0));
        assert_eq!(generated.tschchartinfodefaultroundedcornerradius, Some(0.2));
        assert_unknown_fields_retained(&original, &patched);

        let restored = patch_chart_background_fill(&patched, &original_fill).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn resetting_background_fill_retains_other_style_fields() {
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultgridbackgroundfill: Some(fill_to_native(&solid_fill())),
            tschchartinfodefaultshowborder: Some(true),
            tschchartinfodefaultborderstroke: Some(crate::shapes::stroke_to_native(
                crate::shapes::Stroke::new(
                    RgbaColor::black(),
                    crate::shapes::Width::ONE,
                    Pattern::Solid,
                ),
            )),
            ..Default::default()
        });
        let native_default = native_default_chart_background_fill().unwrap();

        let reset = patch_chart_background_fill(&original, &native_default).unwrap();
        assert_eq!(read_chart_background_fill(&reset).unwrap(), native_default);
        let generated = tsch::generated::ChartStyleArchive::decode(
            generated_chart_style_extension(&reset).unwrap().unwrap(),
        )
        .unwrap();
        assert_eq!(generated.tschchartinfodefaultgridbackgroundfill, None);
        assert_eq!(generated.tschchartinfodefaultshowborder, Some(true));
        assert!(generated.tschchartinfodefaultborderstroke.is_some());
        assert_unknown_fields_retained(&original, &reset);
    }

    #[test]
    fn resetting_only_background_fill_removes_empty_generated_extension() {
        let base = tsch::ChartStyleArchive {
            super_: Some(tss::StyleArchive::default()),
        }
        .encode_to_vec();
        let extension = tsch::generated::ChartStyleArchive {
            tschchartinfodefaultgridbackgroundfill: Some(fill_to_native(&solid_fill())),
            ..Default::default()
        }
        .encode_to_vec();
        let original = patch_length_delimited_field(
            &base,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            false,
            Some(extension.as_slice()),
        )
        .unwrap();

        let reset = patch_chart_background_fill(
            &original,
            &native_default_chart_background_fill().unwrap(),
        )
        .unwrap();
        assert_eq!(reset, base);
        assert_eq!(generated_chart_style_extension(&reset).unwrap(), None);
    }

    #[test]
    fn background_fill_patch_preserves_unknown_selected_fill_payload_bytes() {
        let original_fill = solid_fill();
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultgridbackgroundfill: Some(fill_to_native(&original_fill)),
            ..Default::default()
        });
        let extension = generated_chart_style_extension(&original).unwrap().unwrap();
        let mut selected_payload = fill_to_native(&original_fill).encode_to_vec();
        append_varint_field(
            &mut selected_payload,
            UNMAPPED_SELECTED_FILL_FIELD,
            UNMAPPED_VALUE,
        )
        .unwrap();
        let extension = codec::rewrite(extension, Some(selected_payload.as_slice())).unwrap();
        let original = patch_length_delimited_field(
            &original,
            GENERATED_CHART_STYLE_EXTENSION_FIELD,
            true,
            Some(extension.as_slice()),
        )
        .unwrap();
        let replacement = ShapeFill::Gradient(Gradient::linear(
            color(0.8, 0.2, 0.1),
            color(0.1, 0.4, 0.9),
            Angle::from_degrees(90.0).unwrap(),
        ));

        let patched = patch_chart_background_fill(&original, &replacement).unwrap();
        let original_selected =
            codec::project(generated_chart_style_extension(&original).unwrap().unwrap())
                .unwrap()
                .payload()
                .unwrap();
        let patched_selected =
            codec::project(generated_chart_style_extension(&patched).unwrap().unwrap())
                .unwrap()
                .payload()
                .unwrap();
        assert_eq!(
            raw_field(patched_selected, UNMAPPED_SELECTED_FILL_FIELD),
            raw_field(original_selected, UNMAPPED_SELECTED_FILL_FIELD)
        );

        let restored = patch_chart_background_fill(&patched, &original_fill).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn chart_style_extension_rejects_noncanonical_outer_framing() {
        let original = style_with_unknown_fields(tsch::generated::ChartStyleArchive::default());
        let field = parse_wire_fields(&original)
            .unwrap()
            .into_iter()
            .find(|field| field.number() == GENERATED_CHART_STYLE_EXTENSION_FIELD)
            .unwrap();
        let mut overlong_length = original[..field.key_end()].to_vec();
        let mut length = original[field.key_end()..field.payload_start()].to_vec();
        let last = length.pop().unwrap();
        length.push(last | 0x80);
        length.push(0);
        overlong_length.extend_from_slice(&length);
        overlong_length.extend_from_slice(&original[field.payload_start()..field.end()]);
        overlong_length.extend_from_slice(&original[field.end()..]);

        let error = generated_chart_style_extension(&overlong_length).unwrap_err();
        assert!(error.to_string().contains("noncanonical length prefix"));
    }

    #[test]
    fn malformed_native_background_fills_are_rejected() {
        let malformed = style_with_unknown_fields(tsch::generated::ChartStyleArchive {
            tschchartinfodefaultgridbackgroundfill: Some(crate::protobuf::tsd::FillArchive {
                color: Some(crate::protobuf::tsp::Color::default()),
                gradient: Some(crate::protobuf::tsd::GradientArchive::default()),
                ..Default::default()
            }),
            ..Default::default()
        });
        assert!(read_chart_background_fill(&malformed).is_err());
    }

    fn solid_fill() -> ShapeFill {
        ShapeFill::Solid(color(0.1, 0.3, 0.8))
    }

    fn color(red: f32, green: f32, blue: f32) -> RgbaColor {
        RgbaColor::new(red, green, blue, 1.0, RgbColorSpace::Srgb).unwrap()
    }

    fn style_with_unknown_fields(generated: tsch::generated::ChartStyleArchive) -> Vec<u8> {
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

    fn assert_unknown_fields_retained(original: &[u8], patched: &[u8]) {
        assert_eq!(
            raw_field(patched, UNMAPPED_OUTER_FIELD),
            raw_field(original, UNMAPPED_OUTER_FIELD)
        );
        assert_eq!(
            raw_field(
                generated_chart_style_extension(patched).unwrap().unwrap(),
                UNMAPPED_GENERATED_FIELD,
            ),
            raw_field(
                generated_chart_style_extension(original).unwrap().unwrap(),
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
