//! Binary OfficeArt/PPT record conversion helpers.

use super::super::escher::{UserShapeData, shape_type as escher_shape_type};
use super::super::hyperlink::HyperlinkCollection;
use super::super::shape_style::{
    ArrowStyle, LineCapStyle, LineJoinStyle, LineStyle, LineStyleConfig,
};
use super::super::text_format::Paragraph;
use super::model::{
    ShapeProperties, ShapeType, TextAlignment, WritableShape, WritableSlide, WriteError, Writer,
};
use std::collections::{BTreeMap, HashMap};

#[derive(Default)]
pub(super) struct ConvertedLineProperties {
    pub(super) color: Option<u32>,
    pub(super) width: Option<i32>,
    pub(super) opacity: Option<u32>,
    pub(super) style: Option<u32>,
    pub(super) dash_style: Option<u32>,
    pub(super) start_arrow: Option<u32>,
    pub(super) end_arrow: Option<u32>,
    pub(super) start_arrow_width: Option<u32>,
    pub(super) start_arrow_length: Option<u32>,
    pub(super) end_arrow_width: Option<u32>,
    pub(super) end_arrow_length: Option<u32>,
    pub(super) join_style: Option<u32>,
    pub(super) end_cap_style: Option<u32>,
}

pub(super) fn convert_line_properties(line: Option<&LineStyleConfig>) -> ConvertedLineProperties {
    let Some(line) = line.filter(|line| line.enabled && line.width > 0) else {
        return ConvertedLineProperties::default();
    };
    let has_start_arrow = line.start_arrow != ArrowStyle::None;
    let has_end_arrow = line.end_arrow != ArrowStyle::None;
    ConvertedLineProperties {
        color: Some(line.color.to_rgbx()),
        width: Some(line.width as i32),
        opacity: (line.opacity < 100).then(|| (u32::from(line.opacity) * 65536) / 100),
        style: (line.style != LineStyle::Simple).then_some(line.style as u32),
        dash_style: (line.dash != super::super::shape_style::LineDashStyle::Solid)
            .then_some(line.dash as u32),
        start_arrow: has_start_arrow.then_some(line.start_arrow as u32),
        end_arrow: has_end_arrow.then_some(line.end_arrow as u32),
        start_arrow_width: has_start_arrow.then_some(line.start_arrow_width as u32),
        start_arrow_length: has_start_arrow.then_some(line.start_arrow_length as u32),
        end_arrow_width: has_end_arrow.then_some(line.end_arrow_width as u32),
        end_arrow_length: has_end_arrow.then_some(line.end_arrow_length as u32),
        join_style: (line.join != LineJoinStyle::Miter).then_some(line.join as u32),
        end_cap_style: (line.cap != LineCapStyle::Round).then_some(line.cap as u32),
    }
}
pub(super) fn build_writer_sound_collection(
    slides: &[WritableSlide],
    sound_resources: &BTreeMap<u32, crate::animation::SoundType>,
) -> Result<(Vec<u8>, HashMap<u32, u32>), WriteError> {
    let mut builder = super::super::sound_collection::SoundCollectionBuilder::new(
        super::super::sound_collection::SoundCollectionLimits::default(),
    );

    // Register explicit resources before resolving raw action/legacy references,
    // so an interaction can share a non-built-in embedded animation sound.
    for (sound_id, sound_type) in sound_resources {
        builder
            .register(*sound_id, sound_type)
            .map_err(sound_collection_error)?;
    }
    for slide in slides {
        for shape in &slide.shapes {
            let Some(animation) = &shape.animation_info else {
                continue;
            };
            if let Some(sound) = &animation.sound {
                builder
                    .register(sound.sound_ref, &sound.sound_type)
                    .map_err(sound_collection_error)?;
            }
            if let Some(builds) = &animation.build_list {
                for build in &builds.builds {
                    if let Some(sound) = &build.sound {
                        builder
                            .register(sound.sound_ref, &sound.sound_type)
                            .map_err(sound_collection_error)?;
                    }
                }
            }
        }
    }

    for slide in slides {
        for shape in &slide.shapes {
            if let Some(animation) = &shape.animation_info {
                if let Some(atom) = &animation.legacy_atom {
                    builder
                        .register_reference(atom.sound_id_ref)
                        .map_err(sound_collection_error)?;
                }
                if let Some(sound) = &animation.sound {
                    builder
                        .register_reference(sound.sound_ref)
                        .map_err(sound_collection_error)?;
                }
                if let Some(builds) = &animation.build_list {
                    for build in &builds.builds {
                        if let Some(sound) = &build.sound {
                            builder
                                .register_reference(sound.sound_ref)
                                .map_err(sound_collection_error)?;
                        }
                    }
                }
            }
            for interaction in &shape.properties.interactions {
                builder
                    .register_reference(interaction.sound_id)
                    .map_err(sound_collection_error)?;
            }
            for interaction in &shape.properties.text_interactions {
                builder
                    .register_reference(interaction.interaction.sound_id)
                    .map_err(sound_collection_error)?;
            }
        }
    }

    builder.build().map_err(sound_collection_error)
}

pub(super) fn sound_collection_error(error: std::io::Error) -> WriteError {
    WriteError::InvalidData(error.to_string())
}

pub(super) fn append_child_to_built_container(
    container: &mut Vec<u8>,
    child: &[u8],
) -> Result<(), WriteError> {
    if container.len() < 8 {
        return Err(WriteError::InvalidData(
            "PPT container is missing its record header".to_string(),
        ));
    }
    let stored_len =
        u32::from_le_bytes([container[4], container[5], container[6], container[7]]) as usize;
    if stored_len != container.len() - 8 {
        return Err(WriteError::InvalidData(
            "PPT container length does not match its payload".to_string(),
        ));
    }
    let new_len = stored_len
        .checked_add(child.len())
        .and_then(|len| u32::try_from(len).ok())
        .ok_or_else(|| WriteError::InvalidData("PPT container is too large".to_string()))?;
    container.extend_from_slice(child);
    container[4..8].copy_from_slice(&new_len.to_le_bytes());
    Ok(())
}

/// Convert ShapeType to Escher MSOSPT value
fn shape_type_to_escher(shape_type: ShapeType) -> u16 {
    match shape_type {
        ShapeType::Rectangle => escher_shape_type::RECTANGLE,
        ShapeType::TextBox => escher_shape_type::TEXT_BOX,
        ShapeType::Placeholder => escher_shape_type::RECTANGLE,
        ShapeType::Line => escher_shape_type::LINE,
        ShapeType::Ellipse => escher_shape_type::ELLIPSE,
        ShapeType::RoundRectangle => escher_shape_type::ROUND_RECTANGLE,
        ShapeType::Diamond => escher_shape_type::DIAMOND,
        ShapeType::Triangle => 5, // TRIANGLE
        ShapeType::Arrow => 13,   // ARROW
        ShapeType::Star => 12,    // STAR
        ShapeType::Heart => 74,   // HEART
        ShapeType::Picture => 75, // FRAME (PictureFrame) per POI HSLFPictureShape
        ShapeType::Freeform => escher_shape_type::NOT_PRIMITIVE,
    }
}

/// Convert WritableShape to UserShapeData for Escher serialization
#[cfg(test)]
pub(super) fn convert_shape_to_escher(
    shape: &WritableShape,
    hyperlinks: &HyperlinkCollection,
) -> UserShapeData {
    convert_shape_to_escher_with_sound_mapping(shape, hyperlinks, &HashMap::new())
}

pub(super) fn convert_shape_to_escher_with_sound_mapping(
    shape: &WritableShape,
    hyperlinks: &HyperlinkCollection,
    sound_id_mapping: &HashMap<u32, u32>,
) -> UserShapeData {
    let props = &shape.properties;

    // Extract fill properties from FillStyle
    let (fill_color, fill_type, fill_opacity, fill_back_color, fill_angle, fill_blip_index) = props
        .fill
        .as_ref()
        .map_or((None, None, None, None, None, None), |fill| {
            if !fill.enabled {
                return (None, None, None, None, None, None);
            }

            let color = Some(fill.color.to_rgbx());
            let fill_type = Some(fill.fill_type as u32);

            // Opacity: convert 0-100 to 0-65536
            let opacity = if fill.opacity < 100 {
                Some(((fill.opacity as u32) * 65536) / 100)
            } else {
                None
            };

            // Back color for gradients
            let back_color = fill.back_color.as_ref().map(|c| c.to_rgbx());

            // Gradient angle (degrees * 65536)
            // Per Apache POI HSLFFill.java: "Zero degrees represents a vertical vector from bottom to top"
            // Standard: 0° = horizontal right, 90° = vertical up
            // PPT format: 0° = vertical up, so we need: PPT_angle = 90 - user_angle
            let angle = fill.gradient_angle.map(|a| ((90 - a) as i32) * 65536);

            (
                color,
                fill_type,
                opacity,
                back_color,
                angle,
                fill.picture_index.map(u32::from),
            )
        });

    let line = convert_line_properties(props.line.as_ref());

    // Extract shadow properties from ShadowStyle
    let (has_shadow, shadow_color, shadow_offset_x, shadow_offset_y, shadow_opacity, shadow_type) =
        props
            .shadow
            .as_ref()
            .map_or((false, None, None, None, None, None), |shadow| {
                if !shadow.enabled {
                    (false, None, None, None, None, None)
                } else {
                    (
                        true,
                        Some(shadow.color.to_rgbx()),
                        Some(shadow.offset_x),
                        Some(shadow.offset_y),
                        Some(((shadow.opacity as u32) * 65536) / 100),
                        Some(shadow.shadow_type as u32),
                    )
                }
            });

    // Get text content - prefer paragraphs with formatting
    let mut paragraphs = props.paragraphs.clone().or_else(|| {
        if props.alignment == TextAlignment::Left {
            None
        } else {
            props
                .text
                .as_ref()
                .map(|text| vec![Paragraph::new(text.clone()).align(props.alignment.into())])
        }
    });
    let smart_tag_runs = paragraphs.as_mut().and_then(|paragraphs| {
        if !paragraphs
            .iter()
            .flat_map(|paragraph| &paragraph.runs)
            .any(|run| !run.smart_tag_indices.is_empty())
        {
            return None;
        }
        let mut mappings = Vec::new();
        for (ordinal, run) in paragraphs
            .iter_mut()
            .flat_map(|paragraph| &mut paragraph.runs)
            .enumerate()
        {
            run.style.pp9_run_id = Some(
                u8::try_from(ordinal % 16).expect("a modulo-16 run identifier always fits u8"),
            );
            mappings.push(
                run.smart_tag_indices
                    .iter()
                    .map(|index| index.as_u32())
                    .collect(),
            );
        }
        Some(mappings)
    });
    let text = if paragraphs.is_some() {
        None // Don't use plain text if paragraphs are available
    } else {
        props.text.clone()
    };

    let mut interactions = props.interactions.clone();
    for interaction in &mut interactions {
        remap_sound_reference(&mut interaction.sound_id, sound_id_mapping);
    }
    let mut text_interactions = props.text_interactions.clone();
    for interaction in &mut text_interactions {
        remap_sound_reference(&mut interaction.interaction.sound_id, sound_id_mapping);
    }
    let mut animation_info = shape.animation_info.clone();
    if let Some(animation) = &mut animation_info {
        if let Some(atom) = &mut animation.legacy_atom {
            remap_sound_reference(&mut atom.sound_id_ref, sound_id_mapping);
        }
        if let Some(sound) = &mut animation.sound {
            remap_sound_reference(&mut sound.sound_ref, sound_id_mapping);
        }
        if let Some(builds) = &mut animation.build_list {
            for build in &mut builds.builds {
                if let Some(sound) = &mut build.sound {
                    remap_sound_reference(&mut sound.sound_ref, sound_id_mapping);
                }
            }
        }
    }

    UserShapeData {
        shape_type: shape_type_to_escher(props.shape_type),
        x: props.x,
        y: props.y,
        width: props.width,
        height: props.height,
        fill_color,
        fill_type,
        fill_opacity,
        fill_back_color,
        fill_angle,
        fill_blip_index,
        line_color: line.color,
        line_width: line.width,
        line_opacity: line.opacity,
        line_style: line.style,
        line_dash_style: line.dash_style,
        line_start_arrow: line.start_arrow,
        line_end_arrow: line.end_arrow,
        line_start_arrow_width: line.start_arrow_width,
        line_start_arrow_length: line.start_arrow_length,
        line_end_arrow_width: line.end_arrow_width,
        line_end_arrow_length: line.end_arrow_length,
        line_join_style: line.join_style,
        line_end_cap_style: line.end_cap_style,
        text,
        paragraphs,
        smart_tag_runs,
        text_type: 4,           // OTHER for regular shapes
        placeholder_type: None, // Not a placeholder for regular shapes
        has_shadow,
        flip_h: props.flip_h,
        flip_v: props.flip_v,
        rotation: shape_rotation_to_fixed(props.rotation),
        adjust_values: props.adjust_values.clone(),
        hyperlink_id: props.hyperlink_id,
        hyperlink_action: get_hyperlink_info(props.hyperlink_id, hyperlinks).0,
        hyperlink_jump: get_hyperlink_info(props.hyperlink_id, hyperlinks).1,
        hyperlink_type: get_hyperlink_info(props.hyperlink_id, hyperlinks).2,
        interactions,
        text_interactions,
        picture_index: props.picture_index.map(u32::from),
        freeform_geometry: props.freeform_geometry.clone(),
        animation_info,
        shadow_color,
        shadow_offset_x,
        shadow_offset_y,
        shadow_opacity,
        shadow_type,
    }
}

fn remap_sound_reference(sound_id: &mut u32, mapping: &HashMap<u32, u32>) {
    if let Some(mapped) = mapping.get(sound_id) {
        *sound_id = *mapped;
    }
}

fn shape_rotation_to_fixed(degrees: f32) -> Option<i32> {
    if !degrees.is_finite() || degrees.abs() <= 0.001 {
        return None;
    }
    Some(((f64::from(degrees) % 360.0) * 65536.0).round() as i32)
}

/// Get hyperlink interactive info values based on hyperlink target
/// Returns (action, jump, hyperlink_type)
pub(super) fn get_hyperlink_info(
    hyperlink_id: Option<u32>,
    hyperlinks: &HyperlinkCollection,
) -> (u8, u8, u8) {
    let Some(interaction) = hyperlink_id.and_then(|id| interaction_for_hyperlink(id, hyperlinks))
    else {
        return (4, 0, 8);
    };
    let payload = interaction.atom().to_payload();
    (payload[8], payload[10], payload[12])
}

pub(super) fn interaction_for_hyperlink(
    hyperlink_id: u32,
    hyperlinks: &HyperlinkCollection,
) -> Option<crate::Interaction> {
    use super::super::hyperlink::HyperlinkTarget;
    use crate::{
        Interaction, InteractionAction, InteractionJump, InteractionLinkTarget, InteractionTrigger,
    };

    let hyperlink = hyperlinks.get(hyperlink_id)?;
    let (action, jump, target) = match &hyperlink.target {
        HyperlinkTarget::Url(_) => (
            InteractionAction::Hyperlink,
            InteractionJump::None,
            InteractionLinkTarget::Url,
        ),
        HyperlinkTarget::File(_) => (
            InteractionAction::Hyperlink,
            InteractionJump::None,
            InteractionLinkTarget::OtherFile,
        ),
        HyperlinkTarget::Slide(_) => (
            InteractionAction::Hyperlink,
            InteractionJump::None,
            InteractionLinkTarget::SlideNumber,
        ),
        HyperlinkTarget::NextSlide => (
            InteractionAction::Jump,
            InteractionJump::NextSlide,
            InteractionLinkTarget::NextSlide,
        ),
        HyperlinkTarget::PrevSlide => (
            InteractionAction::Jump,
            InteractionJump::PreviousSlide,
            InteractionLinkTarget::PreviousSlide,
        ),
        HyperlinkTarget::FirstSlide => (
            InteractionAction::Jump,
            InteractionJump::FirstSlide,
            InteractionLinkTarget::FirstSlide,
        ),
        HyperlinkTarget::LastSlide => (
            InteractionAction::Jump,
            InteractionJump::LastSlide,
            InteractionLinkTarget::LastSlide,
        ),
        HyperlinkTarget::EndShow => (
            InteractionAction::Jump,
            InteractionJump::EndShow,
            InteractionLinkTarget::Nil,
        ),
        HyperlinkTarget::CustomShow(_) => (
            InteractionAction::CustomShow,
            InteractionJump::None,
            InteractionLinkTarget::CustomShow,
        ),
    };
    let mut interaction = Interaction::new(InteractionTrigger::Click, action, target);
    interaction.hyperlink_id = hyperlink_id;
    interaction.jump = jump;
    Some(interaction)
}

pub(super) fn shape_text_unit_count(properties: &ShapeProperties) -> Result<u32, WriteError> {
    let units = if let Some(paragraphs) = &properties.paragraphs {
        let mut units = 0usize;
        for (index, paragraph) in paragraphs.iter().enumerate() {
            for run in &paragraph.runs {
                units = units
                    .checked_add(run.text.encode_utf16().count())
                    .ok_or_else(|| {
                        WriteError::InvalidData(
                            "Shape text UTF-16 length overflows usize".to_string(),
                        )
                    })?;
            }
            if index + 1 < paragraphs.len() {
                units = units.checked_add(1).ok_or_else(|| {
                    WriteError::InvalidData("Shape text UTF-16 length overflows usize".to_string())
                })?;
            }
        }
        units
    } else if let Some(text) = &properties.text {
        text.encode_utf16().count()
    } else {
        return Err(WriteError::InvalidData(
            "Shape has no text for a text interaction".to_string(),
        ));
    };
    u32::try_from(units)
        .map_err(|_| WriteError::InvalidData("Shape text exceeds the PPT size limit".to_string()))
}

/// PPT file writer
///
/// Provides methods to create and modify PPT files with full support for:
/// - Shapes with fill, line, and shadow styling
/// - Rich text formatting (bold, italic, colors, sizes)
/// - Pictures/images
/// - Hyperlinks
/// - Speaker notes

impl Writer {
    pub(super) fn build_docinfo_list(
        &self,
        vba_persist_id: Option<u32>,
    ) -> Result<Vec<u8>, WriteError> {
        let ppt11 = super::super::smart_tags::build_document_binary_tag(&self.smart_tags)?;
        Ok(
            super::super::records::create_docinfo_list_container_with_binary_tags(
                self.slide_view_info.as_ref(),
                self.notes_view_info.as_ref(),
                super::super::env_data::VBAInfoAtom {
                    persist_id_ref: vba_persist_id.unwrap_or(0),
                    has_macros: vba_persist_id.is_some(),
                    runtime_version: 2,
                },
                ppt11.as_deref(),
            )?,
        )
    }
}
