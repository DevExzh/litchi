//! Escher facade and byte-level regression tests.

use super::*;
use crate::shapes::geometry::{
    GeometryRect, ShapePathType, extract_geometry_rect, extract_segment_info, extract_shape_path,
    extract_vertices,
};
use crate::writer::shapes::shape_type;
use crate::writer::text_format::{Paragraph, TextRun};
use litchi_odraw::prop::Props;
use litchi_odraw::shape::Flags;
use litchi_odraw::write::{Header, Property, Sp};
use litchi_odraw::{Container, Record, RecordKind};
use zerocopy::IntoBytes;

#[test]
fn test_escher_header() {
    let header = EscherHeader::new(0x0F, 5, record_type::DG_CONTAINER, 100);
    assert_eq!(header.version, 0x0F);
    assert_eq!(header.instance, 5);
    assert_eq!(header.record_type, record_type::DG_CONTAINER);
    assert_eq!(header.length, 100);
}

#[test]
fn test_escher_record_header() {
    let header = Header::new(0x0F, 0, record_type::DG_CONTAINER, 100);
    assert_eq!(header.version(), 0x0F);
    assert_eq!(header.kind().raw(), record_type::DG_CONTAINER);
    assert_eq!(header.len(), 100);
}

#[test]
fn test_escher_record_header_as_bytes() {
    let header = Header::new(0x0F, 1, record_type::SP_CONTAINER, 50);
    let bytes = header.as_bytes();
    assert_eq!(bytes.len(), 8);

    // Verify byte content directly
    // ver_inst = (0x0F & 0x0F) | ((1 & 0x0FFF) << 4) = 0x000F | 0x0010 = 0x001F
    let ver_inst = u16::from_le_bytes([bytes[0], bytes[1]]);
    assert_eq!(ver_inst, 0x001F);

    let rec_type = u16::from_le_bytes([bytes[2], bytes[3]]);
    assert_eq!(rec_type, record_type::SP_CONTAINER);

    let length = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    assert_eq!(length, 50);
}

#[test]
fn test_escher_builder_basic() {
    let mut builder = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
    builder.add_data(&[1, 2, 3, 4]);

    let result = builder.build();
    assert!(result.is_ok());
    let data = result.unwrap();
    assert!(data.len() >= 12); // 8 bytes header + 4 bytes data
}

#[test]
fn test_escher_builder_empty() {
    let builder = EscherBuilder::new(header_version::CONTAINER, 0, record_type::DG_CONTAINER);
    let result = builder.build();
    assert!(result.is_ok());
    let data = result.unwrap();
    assert_eq!(data.len(), 8); // Just header
}

#[test]
fn test_escher_dg_data() {
    let dg_data = EscherDgData::new(10, 1);
    let bytes = dg_data.as_bytes();
    assert_eq!(bytes.len(), 8);

    // Verify byte content directly - fields are csp and spid_cur
    let csp = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let spid_cur = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    assert_eq!(csp, 10);
    // spid_cur = (drawing_id << 10) + shape_count = (1 << 10) + 10 = 1024 + 10 = 1034
    assert_eq!(spid_cur, 1034);
}

#[test]
fn test_escher_sp_data() {
    let sp_data = Sp::with_flags(0x0401, Flags::HAVE_ANCHOR | Flags::HAVE_SPT);
    let bytes = sp_data.as_bytes();
    assert_eq!(bytes.len(), 8);

    // Verify byte content directly - field is spid not sp_id
    let spid = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let flags = u32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    assert_eq!(spid, 0x0401);
    assert_eq!(flags, (Flags::HAVE_ANCHOR | Flags::HAVE_SPT).bits());
}

#[test]
fn test_escher_spgr_data() {
    // Construct using struct literal since there's no new() method
    let data = EscherSpgrData {
        left: 0,
        top: 0,
        right: 1000,
        bottom: 1000,
    };
    let bytes = data.as_bytes();
    assert_eq!(bytes.len(), 16);

    // Verify byte content directly
    let left = i32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let top = i32::from_le_bytes([bytes[4], bytes[5], bytes[6], bytes[7]]);
    let right = i32::from_le_bytes([bytes[8], bytes[9], bytes[10], bytes[11]]);
    let bottom = i32::from_le_bytes([bytes[12], bytes[13], bytes[14], bytes[15]]);
    assert_eq!(left, 0);
    assert_eq!(top, 0);
    assert_eq!(right, 1000);
    assert_eq!(bottom, 1000);
}

#[test]
fn test_escher_property() {
    let prop = Property::new(0x0181, 0x00FF0000);
    let bytes = prop.as_bytes();
    assert_eq!(bytes.len(), 6);

    // Verify byte content directly
    let prop_id = u16::from_le_bytes([bytes[0], bytes[1]]);
    let value = u32::from_le_bytes([bytes[2], bytes[3], bytes[4], bytes[5]]);
    assert_eq!(prop_id, 0x0181);
    assert_eq!(value, 0x00FF0000);
}

#[test]
fn test_shape_flags() {
    let flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
    let value: u32 = flags.bits();
    assert_eq!(value, 0x0A00);

    let flags2 = Flags::FLIP_H | Flags::FLIP_V;
    let value2: u32 = flags2.bits();
    assert_eq!(value2, 0x00C0);
}

#[test]
fn test_user_shape_data_default() {
    let shape = UserShapeData::default();
    assert_eq!(shape.shape_type, shape_type::RECTANGLE);
    assert_eq!(shape.x, 0);
    assert_eq!(shape.y, 0);
    assert_eq!(shape.width, 914400); // 1 inch in EMUs
    assert_eq!(shape.height, 914400);
    assert!(!shape.has_shadow);
    assert!(!shape.flip_h);
    assert!(!shape.flip_v);
}

#[test]
fn test_create_dgg_container() {
    let container = create_dgg_container(5, &[3, 4, 5]);
    assert!(container.is_ok());
    let data = container.unwrap();
    assert!(!data.is_empty());
    assert!(data.len() > 20);
}

#[test]
fn test_create_dgg_container_empty() {
    let container = create_dgg_container(0, &[]);
    assert!(container.is_ok());
    let data = container.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn test_create_dgg_container_many_slides() {
    let slide_counts = vec![1u32; 100];
    let container = create_dgg_container(5, &slide_counts);
    assert!(container.is_ok());
}

#[test]
fn test_create_dg_container_with_shapes_empty() {
    let container = create_dg_container_with_shapes(1, &[]);
    assert!(container.is_ok());
    let data = container.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn test_create_dg_container_with_shapes_single() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        x: 100000,
        y: 100000,
        width: 500000,
        height: 300000,
        text: Some("Test".to_string()),
        ..Default::default()
    };
    let container = create_dg_container_with_shapes(1, &[shape]);
    assert!(container.is_ok());
    let data = container.unwrap();
    assert!(!data.is_empty());
    assert!(data.len() > 50);
}

#[test]
fn test_create_dg_container_with_shapes_multiple() {
    let shapes = vec![
        UserShapeData {
            shape_type: shape_type::RECTANGLE,
            x: 0,
            y: 0,
            width: 100000,
            height: 100000,
            ..Default::default()
        },
        UserShapeData {
            shape_type: shape_type::ELLIPSE,
            x: 200000,
            y: 200000,
            width: 100000,
            height: 100000,
            ..Default::default()
        },
        UserShapeData {
            shape_type: shape_type::LINE,
            x: 0,
            y: 300000,
            width: 300000,
            height: 0,
            ..Default::default()
        },
    ];
    let container = create_dg_container_with_shapes(1, &shapes);
    assert!(container.is_ok());
}

#[test]
fn test_build_shape_properties_rectangle() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        fill_color: Some(0x00FF0000),
        line_color: Some(0x00000000),
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    assert!(!props.is_empty());
    // Should have fill and line properties
    assert!(props.len() >= 4);
    assert!(props.iter().any(|property| {
        property.raw_id() == prop_id::NO_FILL_HIT_TEST
            && property.value() == ppt_prop_value::FILL_STYLE_ENABLED
    }));
}

#[test]
fn test_build_shape_properties_no_fill() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        fill_color: None,
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    // Should have scheme fill with no-fill flag
    let has_no_fill = props.iter().any(|property| {
        property.raw_id() == prop_id::NO_FILL_HIT_TEST
            && property.value() == ppt_prop_value::FILL_STYLE_DISABLED
    });
    assert!(has_no_fill);
}

#[test]
fn test_build_shape_properties_with_shadow() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        has_shadow: true,
        shadow_color: Some(0x00808080),
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    assert!(
        props
            .iter()
            .any(|property| { property.raw_id() == prop_id::SHADOW_TYPE && property.value() == 0 })
    );
    assert!(props.iter().any(|property| {
        property.raw_id() == prop_id::SHADOW_BOOL
            && property.value() == ppt_prop_value::SHADOW_STYLE_ENABLED
    }));
}

#[test]
fn test_build_shape_properties_without_shadow_disables_inheritance() {
    let props = build_shape_properties(&UserShapeData {
        shape_type: shape_type::RECTANGLE,
        ..Default::default()
    });

    assert!(props.iter().any(|property| {
        property.raw_id() == prop_id::SHADOW_BOOL
            && property.value() == ppt_prop_value::SHADOW_STYLE_DISABLED
    }));
}

#[test]
fn test_build_shape_properties_picture() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        picture_index: Some(1),
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    // Should have BLIP property
    let has_blip = props.iter().any(|p| p.raw_id() == 0x4104);
    assert!(has_blip);
}

#[test]
fn test_picture_shape_preserves_rotation_property() {
    let shape = UserShapeData {
        shape_type: 75,
        picture_index: Some(1),
        rotation: Some(-90 * 65536),
        ..Default::default()
    };
    let properties = build_shape_properties(&shape);

    assert!(properties.iter().any(|property| {
        property.raw_id() == prop_id::ROTATION && property.value() == (-90i32 * 65536) as u32
    }));
}

#[test]
fn test_shape_properties_preserve_all_ten_adjustments() {
    let shape = UserShapeData {
        adjust_values: (0..10).map(|index| index * -100).collect(),
        ..Default::default()
    };
    let properties = build_shape_properties(&shape);
    let adjustments: Vec<(u16, u32)> = properties
        .iter()
        .filter_map(|property| {
            let id = { property.raw_id() };
            (prop_id::ADJUST_VALUE..=0x0150)
                .contains(&id)
                .then_some((id, { property.value() }))
        })
        .collect();

    assert_eq!(adjustments.len(), 10);
    assert_eq!(adjustments[0], (0x0147, 0));
    assert_eq!(adjustments[9], (0x0150, (-900i32) as u32));
}

#[test]
fn test_user_shape_rejects_more_than_ten_adjustments() {
    let shape = UserShapeData {
        adjust_values: vec![0; 11],
        ..Default::default()
    };

    assert!(create_user_shape_container(1, &shape).is_err());
}

#[test]
fn test_build_shape_properties_with_arrows() {
    let shape = UserShapeData {
        shape_type: shape_type::LINE,
        line_color: Some(0x00000000),
        line_end_arrow: Some(1), // Triangle arrow
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    // Should have arrow properties
    let has_arrow = props.iter().any(|p| p.raw_id() == prop_id::LINE_END_ARROW);
    assert!(has_arrow);
}

#[test]
fn test_shape_properties_preserve_extended_line_style() {
    let shape = UserShapeData {
        line_color: Some(0x0000_00FF),
        line_width: Some(25400),
        line_opacity: Some(32768),
        line_style: Some(4),
        line_start_arrow: Some(1),
        line_end_arrow: Some(5),
        line_start_arrow_width: Some(0),
        line_start_arrow_length: Some(2),
        line_end_arrow_width: Some(2),
        line_end_arrow_length: Some(0),
        line_join_style: Some(2),
        line_end_cap_style: Some(2),
        ..Default::default()
    };
    let properties = build_shape_properties(&shape);
    let value = |id| {
        properties
            .iter()
            .find(|property| property.raw_id() == id)
            .map(|property| property.value())
    };

    assert_eq!(value(prop_id::LINE_OPACITY), Some(32768));
    assert_eq!(value(prop_id::LINE_STYLE), Some(4));
    assert_eq!(value(prop_id::LINE_START_ARROW_WIDTH), Some(0));
    assert_eq!(value(prop_id::LINE_START_ARROW_LENGTH), Some(2));
    assert_eq!(value(prop_id::LINE_END_ARROW_WIDTH), Some(2));
    assert_eq!(value(prop_id::LINE_END_ARROW_LENGTH), Some(0));
    assert_eq!(value(prop_id::LINE_JOIN_STYLE), Some(2));
    assert_eq!(value(prop_id::LINE_END_CAP_STYLE), Some(2));
}

#[test]
fn test_build_shape_properties_gradient_fill() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        fill_color: Some(0x00FF0000),
        fill_type: Some(4), // Shade/gradient
        fill_back_color: Some(0x0000FF00),
        fill_angle: Some(0),
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    // Should have fill type and back color
    let has_fill_type = props.iter().any(|p| p.raw_id() == prop_id::FILL_TYPE);
    let has_back_color = props.iter().any(|p| p.raw_id() == prop_id::FILL_BACK_COLOR);
    let has_fill_angle = props.iter().any(|p| p.raw_id() == 0x018B && p.value() == 0);
    assert!(has_fill_type);
    assert!(has_back_color);
    assert!(has_fill_angle);
}

#[test]
fn test_shape_properties_preserve_fill_blip_reference() {
    let shape = UserShapeData {
        fill_color: Some(0),
        fill_type: Some(3),
        fill_blip_index: Some(2),
        ..Default::default()
    };
    let properties = build_shape_properties(&shape);

    assert!(
        properties
            .iter()
            .any(|property| { property.raw_id() == prop_id::FILL_BLIP && property.value() == 2 })
    );
    assert_eq!(prop_id::FILL_BLIP, 0x4186);
}

#[test]
fn test_client_textbox_plain_ascii() {
    let textbox = build_client_textbox("Hello World", 4);
    assert!(textbox.is_ok());
    let data = textbox.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn test_client_textbox_unicode() {
    let textbox = build_client_textbox("Hello 世界 🌍", 4);
    assert!(textbox.is_ok());
    let data = textbox.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn client_textbox_plain_style_counts_utf16_code_units() {
    let data = build_client_textbox("😀", 4).unwrap();
    let wrapper = crate::EscherTextboxWrapper::new(data[8..].to_vec()).unwrap();
    let style = wrapper.find_style_text_prop_atom().unwrap();

    assert_eq!(u32::from_le_bytes(style.data[0..4].try_into().unwrap()), 3);
    assert_eq!(
        u32::from_le_bytes(style.data[10..14].try_into().unwrap()),
        3
    );
    assert_eq!(wrapper.text(), "😀");
}

#[test]
fn client_textbox_writes_adjacent_trigger_matched_text_interaction_pairs() {
    use crate::consts::RecordType;
    use crate::{
        Interaction, InteractionAction, InteractionLinkTarget, InteractionTrigger, TextInteraction,
        TextRange,
    };

    let interactions = [
        TextInteraction::new(
            TextRange::new(0, 1).unwrap(),
            Interaction::new(
                InteractionTrigger::Click,
                InteractionAction::NoAction,
                InteractionLinkTarget::Nil,
            ),
        )
        .unwrap(),
        TextInteraction::new(
            TextRange::new(1, 3).unwrap(),
            Interaction::new(
                InteractionTrigger::MouseOver,
                InteractionAction::NoAction,
                InteractionLinkTarget::Nil,
            ),
        )
        .unwrap(),
    ];
    let data = build_client_textbox_with_interactions("A😀", 4, &interactions).unwrap();
    let wrapper = crate::EscherTextboxWrapper::new(data[8..].to_vec()).unwrap();
    let records = wrapper.child_records();
    let tail = &records[records.len() - 4..];

    assert_eq!(
        tail.iter()
            .map(|record| record.record_type)
            .collect::<Vec<_>>(),
        [
            RecordType::InteractiveInfo,
            RecordType::TextInteractiveInfoAtom,
            RecordType::InteractiveInfo,
            RecordType::TextInteractiveInfoAtom,
        ]
    );
    assert_eq!(tail[1].instance, 0);
    assert_eq!(tail[3].instance, 1);
    assert_eq!(wrapper.text_interactions(), interactions);
}

#[test]
fn test_client_textbox_empty() {
    let textbox = build_client_textbox("", 4);
    assert!(textbox.is_ok());
}

#[test]
fn test_client_textbox_formatted() {
    let paragraphs = vec![
        Paragraph::new("First paragraph"),
        Paragraph::with_runs(vec![
            TextRun::new("Bold text").bold(),
            TextRun::new(" and "),
            TextRun::new("italic").italic(),
        ]),
    ];
    let textbox = build_client_textbox_formatted(&paragraphs, 1);
    assert!(textbox.is_ok());
    let data = textbox.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn formatted_textbox_round_trips_non_bmp_multi_paragraph_runs() {
    let paragraphs = vec![
        Paragraph::with_runs(vec![
            TextRun::new("😀")
                .bold()
                .shadow()
                .size(24)
                .color_rgb(10, 20, 30)
                .font(2)
                .baseline_position(30),
        ]),
        Paragraph::with_runs(vec![
            TextRun::new("x")
                .italic()
                .embossed()
                .size(14)
                .color_scheme(4)
                .font(3)
                .baseline_position(-25),
        ]),
    ];
    let data = build_client_textbox_formatted(&paragraphs, 1).unwrap();
    let wrapper = crate::EscherTextboxWrapper::new(data[8..].to_vec()).unwrap();

    assert_eq!(wrapper.text(), "😀\rx");
    assert_eq!(wrapper.runs().len(), 2);
    assert_eq!(wrapper.runs()[0].text, "😀\r");
    assert!(wrapper.runs()[0].formatting.bold);
    assert!(wrapper.runs()[0].formatting.shadow);
    assert_eq!(wrapper.runs()[0].formatting.baseline_position, Some(30));
    assert_eq!(wrapper.runs()[0].formatting.font_size, Some(24));
    assert_eq!(wrapper.runs()[0].formatting.font_color, Some(0x000A_141E));
    assert_eq!(wrapper.runs()[0].formatting.font_index, Some(2));
    assert_eq!(wrapper.runs()[1].text, "x");
    assert!(wrapper.runs()[1].formatting.italic);
    assert!(wrapper.runs()[1].formatting.embossed);
    assert_eq!(wrapper.runs()[1].formatting.baseline_position, Some(-25));
    assert_eq!(wrapper.runs()[1].formatting.font_size, Some(14));
    assert_eq!(wrapper.runs()[1].formatting.font_color, None);
    assert_eq!(wrapper.runs()[1].formatting.font_scheme_color, Some(4));
    assert_eq!(wrapper.runs()[1].formatting.font_index, Some(3));
}

#[test]
fn test_build_client_data_with_hyperlink() {
    let client_data = build_client_data_with_hyperlink(1, 4, 0, 8);
    assert!(client_data.is_ok());
    let data = client_data.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn typed_interactions_coexist_in_client_data_grammar_order() {
    let click = crate::Interaction::new(
        crate::InteractionTrigger::Click,
        crate::InteractionAction::Macro,
        crate::InteractionLinkTarget::Nil,
    )
    .with_macro_name("Run")
    .unwrap();
    let hover = crate::Interaction::new(
        crate::InteractionTrigger::MouseOver,
        crate::InteractionAction::Ole,
        crate::InteractionLinkTarget::Nil,
    );
    let shape = UserShapeData {
        interactions: vec![click, hover],
        animation_info: Some(crate::animation::AnimationInfo::new()),
        placeholder_type: Some(6),
        ..Default::default()
    };

    let bytes = create_user_shape_container(45, &shape).unwrap();
    let (root, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    let root = Container::try_new(root).expect("shape container");
    let record = root
        .find(RecordKind::ClientData)
        .unwrap()
        .expect("ClientData");
    let mut complete = Vec::with_capacity(8 + record.data().len());
    let version_instance = u16::from(record.version()) | (record.instance() << 4);
    complete.extend_from_slice(&version_instance.to_le_bytes());
    complete.extend_from_slice(&record.raw_kind().to_le_bytes());
    complete.extend_from_slice(&record.len().to_le_bytes());
    complete.extend_from_slice(record.data());

    let client_data = crate::ClientData::parse(&complete).unwrap();
    assert!(client_data.animation_info().is_some());
    assert!(client_data.mouse_click_interactive_info().is_some());
    assert!(client_data.mouse_over_interactive_info().is_some());
    assert!(client_data.placeholder().is_some());
}

#[test]
fn duplicate_typed_interaction_triggers_are_rejected() {
    let click = crate::Interaction::new(
        crate::InteractionTrigger::Click,
        crate::InteractionAction::NoAction,
        crate::InteractionLinkTarget::Nil,
    );
    let shape = UserShapeData {
        interactions: vec![click.clone(), click],
        ..Default::default()
    };

    assert!(create_user_shape_container(45, &shape).is_err());
}

#[test]
fn test_build_client_data_with_placeholder() {
    let client_data = build_client_data_with_placeholder(6); // NotesBody
    assert!(client_data.is_ok());
    let data = client_data.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn test_shape_type_constants() {
    // Verify all shape type constants are correctly defined
    assert_eq!(shape_type::NOT_PRIMITIVE, 0);
    assert_eq!(shape_type::RECTANGLE, 1);
    assert_eq!(shape_type::ROUND_RECTANGLE, 2);
    assert_eq!(shape_type::ELLIPSE, 3);
    assert_eq!(shape_type::DIAMOND, 4);
    assert_eq!(shape_type::ISOCELES_TRIANGLE, 5);
    assert_eq!(shape_type::RIGHT_TRIANGLE, 6);
    assert_eq!(shape_type::PARALLELOGRAM, 7);
    assert_eq!(shape_type::TRAPEZOID, 8);
    assert_eq!(shape_type::HEXAGON, 9);
    assert_eq!(shape_type::OCTAGON, 10);
    assert_eq!(shape_type::PLUS, 11);
    assert_eq!(shape_type::STAR, 12);
    assert_eq!(shape_type::ARROW, 13);
    assert_eq!(shape_type::THICK_ARROW, 14);
    assert_eq!(shape_type::LINE, 20);
    assert_eq!(shape_type::TEXT_BOX, 202);
}

#[test]
fn test_prop_id_constants() {
    // Verify property ID constants
    assert_eq!(prop_id::FILL_TYPE, 0x0180);
    assert_eq!(prop_id::FILL_COLOR, 0x0181);
    assert_eq!(prop_id::FILL_OPACITY, 0x0182);
    assert_eq!(prop_id::FILL_BACK_COLOR, 0x0183);
    assert_eq!(prop_id::LINE_COLOR, 0x01C0);
    assert_eq!(prop_id::LINE_WIDTH, 0x01CB);
    assert_eq!(prop_id::LINE_START_ARROW, 0x01D0);
    assert_eq!(prop_id::LINE_END_ARROW, 0x01D1);
    assert_eq!(prop_id::SHADOW_TYPE, 0x0200);
    assert_eq!(prop_id::SHADOW_COLOR, 0x0201);
    assert_eq!(prop_id::NO_FILL_HIT_TEST, 0x01BF);
    assert_eq!(prop_id::LINE_STYLE_BOOL, 0x01FF);
    assert_eq!(prop_id::BACKGROUND_SHAPE, 0x033F);
}

#[test]
fn test_default_and_background_boolean_property_groups() {
    let dgg_line_bool = DGG_DEFAULT_PROPERTIES[6];
    let dgg_line_bool_id = { dgg_line_bool.raw_id() };
    let dgg_line_bool_value = { dgg_line_bool.value() };
    assert_eq!(dgg_line_bool_id, 0x01FF);
    assert_eq!(dgg_line_bool_value, 0x0008_0008);

    let actual: Vec<(u16, u32)> = BG_SHAPE_PROPERTIES
        .iter()
        .map(|property| ({ property.raw_id() }, { property.value() }))
        .collect();
    assert_eq!(
        actual,
        [
            (0x0181, 0x0800_0000),
            (0x0183, 0x0800_0005),
            (0x0193, 0x0099_936E),
            (0x0194, 0x0076_B1BE),
            (0x01BF, 0x0012_0012),
            (0x01FF, 0x0008_0000),
            (0x0304, 0x0000_0009),
            (0x033F, 0x0001_0001),
        ]
    );
}

#[test]
fn test_record_type_constants() {
    assert_eq!(record_type::DGG_CONTAINER, 0xF000);
    assert_eq!(record_type::DGG, 0xF006);
    assert_eq!(record_type::DG_CONTAINER, 0xF002);
    assert_eq!(record_type::DG, 0xF008);
    assert_eq!(record_type::SPGR_CONTAINER, 0xF003);
    assert_eq!(record_type::SP_CONTAINER, 0xF004);
    assert_eq!(record_type::SP, 0xF00A);
    assert_eq!(record_type::SPGR, 0xF009);
    assert_eq!(record_type::OPT, 0xF00B);
    assert_eq!(record_type::CLIENT_ANCHOR, 0xF010);
    assert_eq!(record_type::CLIENT_DATA, 0xF011);
}

#[test]
fn test_header_version_constants() {
    assert_eq!(header_version::CONTAINER, 0x0F);
    // DGG doesn't exist in header_version - the DGG record type is different from header version
    assert_eq!(header_version::DG, 0x00);
    assert_eq!(header_version::SPGR, 0x01);
    assert_eq!(header_version::SP, 0x02);
}

#[test]
fn test_escher_header_as_bytes() {
    let header = EscherHeader::new(0x0F, 5, record_type::DG_CONTAINER, 100);
    let mut buf = Vec::new();
    header.write(&mut buf).unwrap();
    assert_eq!(buf.len(), 8);

    // Verify byte content directly
    // version in low 4 bits, instance in high 12 bits
    let ver_inst = u16::from_le_bytes([buf[0], buf[1]]);
    assert_eq!(ver_inst, 0x005F); // 0x0F | (5 << 4) = 0x0F | 0x50 = 0x5F

    let rec_type = u16::from_le_bytes([buf[2], buf[3]]);
    assert_eq!(rec_type, record_type::DG_CONTAINER);

    let length = u32::from_le_bytes([buf[4], buf[5], buf[6], buf[7]]);
    assert_eq!(length, 100);
}

#[test]
fn test_create_dg_container_with_flip_flags() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        x: 100000,
        y: 100000,
        width: 500000,
        height: 300000,
        flip_h: true,
        flip_v: true,
        ..Default::default()
    };
    let container = create_dg_container_with_shapes(1, &[shape]);
    assert!(container.is_ok());
}

#[test]
fn test_create_dg_container_with_dash_style() {
    let shape = UserShapeData {
        shape_type: shape_type::LINE,
        line_color: Some(0x00000000),
        line_dash_style: Some(1), // Dash
        ..Default::default()
    };
    let container = create_dg_container_with_shapes(1, &[shape]);
    assert!(container.is_ok());
}

#[test]
fn test_freeform_geometry_round_trips_through_opt_record() {
    let geometry = FreeformGeometry::new(
        GeometryRect::new(0, 0, 21600, 21600),
        ShapePathType::Complex,
        vec![(0, 0), (10800, 21600), (21600, 0)],
        vec![0x4000, 0x0001, 0x0001, 0x8000],
    );
    let shape = UserShapeData {
        shape_type: shape_type::NOT_PRIMITIVE,
        freeform_geometry: Some(geometry),
        ..Default::default()
    };

    let bytes = create_user_shape_container(45, &shape).unwrap();
    let (root, consumed) = Record::parse(&bytes, 0).expect("shape container record");
    assert_eq!(consumed, bytes.len());
    let root = Container::try_new(root).expect("shape container");
    let opt = root
        .find(RecordKind::Opt)
        .expect("valid shape container")
        .expect("OPT record");
    let properties = Props::parse(&opt).expect("valid OPT properties");

    assert_eq!(
        extract_geometry_rect(&properties),
        Some(GeometryRect::new(0, 0, 21600, 21600))
    );
    assert_eq!(
        extract_shape_path(&properties),
        Some(ShapePathType::Complex)
    );
    assert_eq!(
        extract_vertices(&properties)
            .expect("vertex array")
            .iter()
            .collect::<Vec<_>>(),
        [(0, 0), (10800, 21600), (21600, 0)]
    );
    let segment_words: Vec<u16> = extract_segment_info(&properties)
        .expect("segment array")
        .chunks_exact(2)
        .map(|word| u16::from_le_bytes([word[0], word[1]]))
        .collect();
    assert_eq!(segment_words, [0x4000, 0x0001, 0x0001, 0x8000]);
}

#[test]
fn test_shape_properties_with_line_width() {
    let shape = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        line_color: Some(0x00000000),
        line_width: Some(25400), // 2pt in EMUs
        ..Default::default()
    };
    let props = build_shape_properties(&shape);
    let has_width = props.iter().any(|p| p.raw_id() == prop_id::LINE_WIDTH);
    assert!(has_width);
}

#[test]
fn test_multiple_paragraphs_textbox() {
    let paragraphs = vec![
        Paragraph::new("First paragraph with some text"),
        Paragraph::new("Second paragraph").center(),
        Paragraph::new("Third paragraph").right(),
    ];
    let textbox = build_client_textbox_formatted(&paragraphs, 1);
    assert!(textbox.is_ok());
    let data = textbox.unwrap();
    assert!(!data.is_empty());
}

#[test]
fn test_text_with_multiple_runs() {
    let runs = vec![
        TextRun::new("Normal "),
        TextRun::new("bold ").bold(),
        TextRun::new("italic").italic(),
        TextRun::new(" "),
        TextRun::new("underline").underline(),
    ];
    let para = Paragraph::with_runs(runs);
    let textbox = build_client_textbox_formatted(&[para], 4);
    assert!(textbox.is_ok());
}
