//! `OfficeArt` shape-container record family.

use zerocopy::IntoBytes;

use litchi_core::unit::emu_i32_to_ppt_master_i16_round;
use litchi_odraw::shape::Flags;
use litchi_odraw::write::{COMPLEX, Property, Sp, record_type};

use super::super::{ChildAnchor, Error, UserShapeData, header_version, prop_id};
use super::client_data::{
    append_client_data_payload, append_client_data_record_payload,
    build_client_data_with_animation, build_client_data_with_placeholder,
    legacy_hyperlink_interaction,
};
use super::properties::build_shape_properties;
use super::text::{
    build_client_textbox_formatted_with_interactions, build_client_textbox_with_interactions,
};
use super::validation::validate_user_shape;
use super::wire::EscherBuilder;

#[derive(Clone, Copy)]
enum ShapeAnchor {
    /// Compact eight-byte PPT host anchor for ordinary top-level shapes.
    Ppt,
    /// Sixteen-byte host anchor used for a root group's direct member.
    Host(ChildAnchor),
    /// Sixteen-byte group-coordinate anchor for a nested member.
    Child(ChildAnchor),
}

/// Creates a standalone PPT shape container with a host `ClientAnchor`.
pub(crate) fn create_user_shape_container(
    shape_id: u32,
    shape: &UserShapeData,
) -> Result<Vec<u8>, Error> {
    create_shape_container(shape_id, shape, ShapeAnchor::Ppt)
}

/// Creates a group-member shape container with a typed `ChildAnchor`.
pub(crate) fn create_child_shape_container(
    shape_id: u32,
    shape: &UserShapeData,
    child_anchor: ChildAnchor,
) -> Result<Vec<u8>, Error> {
    create_shape_container(shape_id, shape, ShapeAnchor::Child(child_anchor))
}

/// Creates a root-group member with the fixed-width host anchor used by PPT
/// nested group records.
pub(super) fn create_root_group_shape_container(
    shape_id: u32,
    shape: &UserShapeData,
    anchor: ChildAnchor,
) -> Result<Vec<u8>, Error> {
    create_shape_container(shape_id, shape, ShapeAnchor::Host(anchor))
}

fn create_shape_container(
    shape_id: u32,
    shape: &UserShapeData,
    anchor_kind: ShapeAnchor,
) -> Result<Vec<u8>, Error> {
    validate_user_shape(shape)?;
    let mut container = EscherBuilder::new(header_version::CONTAINER, 0, record_type::SP_CONTAINER);

    let mut flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
    if matches!(anchor_kind, ShapeAnchor::Child(_)) {
        flags |= Flags::CHILD;
    }
    if shape.flip_h {
        flags |= Flags::FLIP_H;
    }
    if shape.flip_v {
        flags |= Flags::FLIP_V;
    }

    let mut sp = EscherBuilder::new(header_version::SP, shape.shape_type, record_type::SP);
    sp.add_data(Sp::with_flags(shape_id, flags).as_bytes());
    container.add_data(&sp.build()?);

    let mut properties: Vec<(Property, Option<Vec<u8>>)> = build_shape_properties(shape)
        .into_iter()
        .map(|property| (property, None))
        .collect();
    if let Some(geometry) = &shape.freeform_geometry {
        let rect = geometry.coordinate_space();
        let (vertices, segments) = geometry.encode_arrays()?;
        #[allow(
            clippy::cast_possible_truncation,
            reason = "`encode_arrays` validates the vertex and segment counts to fit in `u16`"
        )]
        properties.extend([
            (
                Property::new(prop_id::GEOM_LEFT, rect.left.cast_unsigned()),
                None,
            ),
            (
                Property::new(prop_id::GEOM_TOP, rect.top.cast_unsigned()),
                None,
            ),
            (
                Property::new(prop_id::GEOM_RIGHT, rect.right.cast_unsigned()),
                None,
            ),
            (
                Property::new(prop_id::GEOM_BOTTOM, rect.bottom.cast_unsigned()),
                None,
            ),
            (
                Property::new(prop_id::SHAPE_PATH, geometry.path_type() as u32),
                None,
            ),
            (
                Property::new(prop_id::VERTICES | COMPLEX, vertices.len() as u32),
                Some(vertices),
            ),
            (
                Property::new(prop_id::SEGMENT_INFO | COMPLEX, segments.len() as u32),
                Some(segments),
            ),
        ]);
    }
    properties.sort_by_key(|(property, _)| property.raw_id() & 0x3FFF);
    #[allow(
        clippy::cast_possible_truncation,
        reason = "a shape property table holds at most a few dozen `OfficeArt` properties"
    )]
    let mut opt = EscherBuilder::new(
        header_version::OPT,
        properties.len() as u16,
        record_type::OPT,
    );
    for (property, _) in &properties {
        opt.add_data(property.as_bytes());
    }
    for (_, complex_data) in &properties {
        if let Some(data) = complex_data {
            opt.add_data(data);
        }
    }
    container.add_data(&opt.build()?);

    match anchor_kind {
        ShapeAnchor::Child(anchor_data) => {
            let mut anchor =
                EscherBuilder::new(header_version::SIMPLE, 0, record_type::CHILD_ANCHOR);
            anchor.add_data(anchor_data.as_bytes());
            container.add_data(&anchor.build()?);
        },
        ShapeAnchor::Host(anchor_data) => {
            let mut anchor =
                EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
            anchor.add_data(anchor_data.as_bytes());
            container.add_data(&anchor.build()?);
        },
        ShapeAnchor::Ppt => {
            // PPT top-level shapes use the compact eight-byte host anchor.
            let mut anchor =
                EscherBuilder::new(header_version::SIMPLE, 0, record_type::CLIENT_ANCHOR);
            let x1 = emu_i32_to_ppt_master_i16_round(shape.x);
            let y1 = emu_i32_to_ppt_master_i16_round(shape.y);
            let x2 = emu_i32_to_ppt_master_i16_round(shape.x + shape.width);
            let y2 = emu_i32_to_ppt_master_i16_round(shape.y + shape.height);
            anchor.add_data(&y1.to_le_bytes());
            anchor.add_data(&x1.to_le_bytes());
            anchor.add_data(&x2.to_le_bytes());
            anchor.add_data(&y2.to_le_bytes());
            container.add_data(&anchor.build()?);
        },
    }

    let mut client_data = if let Some(ref animation_info) = shape.animation_info {
        Some(build_client_data_with_animation(animation_info)?)
    } else {
        None
    };
    let legacy_click = shape
        .hyperlink_id
        .map(|hyperlink_id| {
            legacy_hyperlink_interaction(
                hyperlink_id,
                shape.hyperlink_action,
                shape.hyperlink_jump,
                shape.hyperlink_type,
            )
        })
        .transpose()?;
    for trigger in [
        crate::InteractionTrigger::Click,
        crate::InteractionTrigger::MouseOver,
    ] {
        let mut matching = shape
            .interactions
            .iter()
            .filter(|interaction| interaction.trigger == trigger);
        let declared_interaction = matching.next();
        if matching.next().is_some() {
            return Err(std::io::Error::other(
                "shape contains duplicate interactive triggers",
            ));
        }
        let interaction_or_legacy = declared_interaction.or_else(|| {
            (trigger == crate::InteractionTrigger::Click)
                .then_some(legacy_click.as_ref())
                .flatten()
        });
        if let Some(interaction) = interaction_or_legacy {
            let bytes = interaction
                .to_bytes_with_limits(crate::InteractionLimits::default())
                .map_err(|error| std::io::Error::other(error.to_string()))?;
            append_client_data_payload(&mut client_data, &bytes)?;
        }
    }
    if let Some(placeholder_type) = shape.placeholder_type {
        let placeholder = build_client_data_with_placeholder(placeholder_type)?;
        append_client_data_record_payload(&mut client_data, &placeholder)?;
    }
    if let Some(runs) = &shape.smart_tag_runs
        && let Some(programmable_tags) =
            crate::writer::smart_tags::build_shape_text_extensions(runs)
                .map_err(|error| std::io::Error::other(error.to_string()))?
    {
        append_client_data_payload(&mut client_data, &programmable_tags)?;
    }
    if let Some(client_data_bytes) = client_data {
        crate::ClientData::parse(&client_data_bytes)
            .map_err(|error| std::io::Error::other(error.to_string()))?;
        container.add_data(&client_data_bytes);
    }

    if let Some(paragraphs) = &shape.paragraphs {
        if !paragraphs.is_empty() {
            let textbox = build_client_textbox_formatted_with_interactions(
                paragraphs,
                shape.text_type,
                &shape.text_interactions,
            )?;
            container.add_data(&textbox);
        } else if !shape.text_interactions.is_empty() {
            return Err(std::io::Error::other(
                "shape has text interactions but no corresponding text",
            ));
        }
    } else if let Some(text) = &shape.text {
        let textbox = build_client_textbox_with_interactions(
            text,
            shape.text_type,
            &shape.text_interactions,
        )?;
        container.add_data(&textbox);
    } else if !shape.text_interactions.is_empty() {
        return Err(std::io::Error::other(
            "shape has text interactions but no corresponding text",
        ));
    }

    container.build()
}
