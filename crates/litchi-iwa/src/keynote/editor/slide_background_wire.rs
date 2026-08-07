//! Wire-preserving Keynote slide-style variation lifecycle.

use litchi_keynote::background::Background;

use super::slide_background_color::native_color_space;
use super::slide_background_gradient_wire::gradient_to_fill;
use super::slide_style_graph::{
    SLIDE_MESSAGE_TYPE, SLIDE_STYLE_MESSAGE_TYPE, is_collapsible_background_variation,
    patch_slide_style_reference, reference, style_is_exclusive,
};
use super::slide_style_metadata::update_package_metadata;
use super::slide_style_registry::patch_stylesheet;
use super::*;

pub(super) fn set_slide_background(
    editor: &mut KeynoteEditor,
    slide_index: usize,
    background: Background,
    inherited_fill_payload: &[u8],
) -> Result<()> {
    let slides = editor.slides()?;
    let slide_info = slides.get(slide_index).ok_or_else(|| {
        Error::ParseError(format!(
            "Keynote slide index {slide_index} is out of range for {} slides",
            slides.len()
        ))
    })?;
    let graph = ObjectGraph::read(editor.package())?;
    let slide: kn::SlideArchive =
        graph.decode_type(slide_info.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
    let old_style_id = slide.style.identifier;
    let old_style: kn::SlideStyleArchive = graph.decode_type(
        old_style_id,
        SLIDE_STYLE_MESSAGE_TYPE,
        "KN.SlideStyleArchive",
    )?;
    let stylesheet_id = old_style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Keynote slide style {old_style_id} has no stylesheet reference"
            ))
        })?;
    let slide_archive = graph.archive_name(slide_info.slide_id)?.to_owned();
    let stylesheet_archive = graph.archive_name(stylesheet_id)?.to_owned();
    if graph.archive_name(old_style_id)? != stylesheet_archive {
        return Err(Error::InvalidFormat(format!(
            "Keynote slide style {old_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let old_style_raw = graph.message_data_type(
        old_style_id,
        SLIDE_STYLE_MESSAGE_TYPE,
        "KN.SlideStyleArchive",
    )?;
    let disposable = is_collapsible_background_variation(&old_style, old_style_raw)?
        && style_is_exclusive(&graph, old_style_id)?;
    let parent_style_id = if disposable {
        old_style.super_.parent.as_ref().unwrap().identifier
    } else {
        old_style_id
    };
    let new_style_id = next_object_identifier(editor.package())?;
    let fill_payload = encode_background_fill(&background, inherited_fill_payload)?;
    let new_style = new_style_object(new_style_id, parent_style_id, stylesheet_id, &fill_payload)?;

    let mut staged = editor.package().clone();
    patch_slide_style_reference(
        &mut staged,
        &slide_archive,
        slide_info.slide_id,
        old_style_id,
        new_style_id,
    )?;
    patch_stylesheet(
        &mut staged,
        &stylesheet_archive,
        stylesheet_id,
        disposable.then_some(old_style_id),
        Some((parent_style_id, new_style_id, new_style)),
    )?;
    update_package_metadata(
        &mut staged,
        &slide_archive,
        &stylesheet_archive,
        disposable.then_some(old_style_id),
        Some(new_style_id),
    )?;
    set_package_last_object_identifier(&mut staged, new_style_id)?;

    let bytes = staged.to_bytes()?;
    let verified = KeynoteEditor::from_bytes(&bytes)?;
    if verified.slide_background(slide_index)? != background {
        return Err(Error::InvalidFormat(
            "Keynote slide-background update failed validation".to_owned(),
        ));
    }
    editor.text = IWorkTextEditor::from_package(staged);
    Ok(())
}

fn encode_background_fill(
    background: &Background,
    inherited_fill_payload: &[u8],
) -> Result<Vec<u8>> {
    match background {
        Background::None => Ok(tsd::FillArchive::default().encode_to_vec()),
        Background::Opaque(payload) => {
            tsd::FillArchive::decode(payload.as_bytes())?;
            Ok(payload.as_bytes().to_vec())
        },
        Background::Gradient(gradient) => Ok(gradient_to_fill(gradient)),
        Background::Solid(color) => {
            let existing = tsd::FillArchive::decode(inherited_fill_payload)?;
            let mut data = if existing.gradient.is_none() && existing.image.is_none() {
                if existing.color.is_some() {
                    inherited_fill_payload.to_vec()
                } else {
                    patch_length_delimited_field(
                        inherited_fill_payload,
                        1,
                        false,
                        Some(
                            &tsp::Color {
                                model: tsp::color::ColorModel::Rgb as i32,
                                ..Default::default()
                            }
                            .encode_to_vec(),
                        ),
                    )?
                }
            } else {
                tsd::FillArchive {
                    color: Some(tsp::Color {
                        model: tsp::color::ColorModel::Rgb as i32,
                        r: Some(color.red()),
                        g: Some(color.green()),
                        b: Some(color.blue()),
                        rgbspace: Some(native_color_space(color.color_space())),
                        a: Some(color.alpha()),
                        ..Default::default()
                    }),
                    ..Default::default()
                }
                .encode_to_vec()
            };
            data = patch_nested_varint_field(
                &data,
                &[1, 1],
                true,
                Some(tsp::color::ColorModel::Rgb as u64),
            )?;
            for (field, value) in [
                (3, color.red()),
                (4, color.green()),
                (5, color.blue()),
                (6, color.alpha()),
            ] {
                data = patch_nested_fixed32_field(&data, &[1, field], true, Some(value.to_bits()))?;
            }
            data = patch_nested_varint_field(
                &data,
                &[1, 12],
                true,
                Some(native_color_space(color.color_space()) as u64),
            )?;
            let verified = tsd::FillArchive::decode(data.as_slice())?;
            if verified.gradient.is_some() || verified.image.is_some() {
                return Err(Error::InvalidFormat(
                    "Keynote solid background retained an incompatible fill".to_owned(),
                ));
            }
            Ok(data)
        },
        _ => Err(Error::InvalidFormat(
            "unsupported Keynote slide-background semantic variant".to_owned(),
        )),
    }
}

fn new_style_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    fill_payload: &[u8],
) -> Result<ArchiveObject> {
    let style = kn::SlideStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        slide_properties: Some(kn::SlideStylePropertiesArchive {
            fill: Some(tsd::FillArchive::default()),
            ..Default::default()
        }),
    };
    let data = patch_nested_length_delimited_field(
        &style.encode_to_vec(),
        &[11, 1],
        true,
        Some(fill_payload),
    )?;
    kn::SlideStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SLIDE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    Ok(object)
}
