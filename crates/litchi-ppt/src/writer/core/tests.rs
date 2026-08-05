//! Focused regression tests for the PPT writer core.

use super::super::escher::FreeformGeometry;
use super::codec::{convert_line_properties, convert_shape_to_escher, get_hyperlink_info};
use super::*;
use crate::shapes::geometry::{GeometryRect, ShapePathType};
use crate::writer::shape_style::{ArrowSize, ShapeColor};
use std::io::Cursor;

#[test]
fn test_create_writer() {
    let writer = Writer::new();
    assert_eq!(writer.slides.len(), 0);
    assert_eq!(writer.slide_width, 9144000);
    assert_eq!(writer.slide_height, 6858000);
}

#[test]
fn test_create_widescreen() {
    let writer = Writer::new_widescreen();
    assert_eq!(writer.slide_width, 9144000);
    assert_eq!(writer.slide_height, 5143500);
}

#[test]
fn test_add_slide() {
    let mut writer = Writer::new();
    let idx = writer.add_slide().unwrap();
    assert_eq!(idx, 0);
    assert_eq!(writer.slides.len(), 1);
}

#[test]
fn test_add_multiple_slides() {
    let mut writer = Writer::new();
    let idx1 = writer.add_slide().unwrap();
    let idx2 = writer.add_slide().unwrap();
    let idx3 = writer.add_slide().unwrap();
    assert_eq!(idx1, 0);
    assert_eq!(idx2, 1);
    assert_eq!(idx3, 2);
    assert_eq!(writer.slide_count(), 3);
}

#[test]
fn test_add_and_write_freeform_shape() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let geometry = FreeformGeometry::new(
        GeometryRect::new(0, 0, 21600, 21600),
        ShapePathType::Complex,
        vec![(0, 0), (10800, 21600), (21600, 0)],
        vec![0x4000, 0x0001, 0x0001, 0x8000],
    );

    writer
        .add_freeform(slide, 10, 20, 300, 200, geometry)
        .unwrap();
    assert_eq!(writer.slides[slide].shapes.len(), 1);
    assert_eq!(
        writer.slides[slide].shapes[0].properties.shape_type,
        ShapeType::Freeform
    );

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    assert!(!output.into_inner().is_empty());
}

#[test]
fn test_rejects_empty_freeform_geometry() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let geometry = FreeformGeometry::new(
        GeometryRect::new(0, 0, 21600, 21600),
        ShapePathType::Complex,
        Vec::new(),
        vec![0x8000],
    );

    assert!(
        writer
            .add_freeform(slide, 0, 0, 100, 100, geometry)
            .is_err()
    );
    assert!(writer.slides[slide].shapes.is_empty());
}

#[test]
fn test_generic_styled_shape_rejects_geometryless_freeform() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let result = writer.add_styled_shape(
        slide,
        ShapeType::Freeform,
        0,
        0,
        100,
        100,
        ShapeStyle::default(),
    );

    assert!(result.is_err());
    assert!(writer.slides[slide].shapes.is_empty());
}

#[test]
fn test_add_textbox() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 100, 50, "Test").unwrap();
    assert_eq!(writer.slides[0].shapes.len(), 1);
}

#[test]
fn test_plain_text_alignment_and_rotation_reach_escher_shape() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 20, 300, 100, "Centered")
        .unwrap();
    writer
        .set_last_shape_text_alignment(slide, TextAlignment::Center)
        .unwrap();
    writer.set_last_shape_rotation(slide, 450.5).unwrap();

    let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
    assert!(shape.text.is_none());
    let paragraphs = shape.paragraphs.as_ref().expect("formatted paragraph");
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].alignment, TextAlign::Center);
    assert_eq!(shape.rotation, Some((90 * 65536) + 32768));
}

#[test]
fn test_alignment_setter_updates_rich_paragraphs() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_rich_textbox(
            slide,
            0,
            0,
            200,
            100,
            vec![Paragraph::new("One"), Paragraph::new("Two")],
        )
        .unwrap();
    writer
        .set_last_shape_text_alignment(slide, TextAlignment::Justify)
        .unwrap();

    let paragraphs = writer.slides[slide].shapes[0]
        .properties
        .paragraphs
        .as_ref()
        .unwrap();
    assert!(
        paragraphs
            .iter()
            .all(|paragraph| paragraph.alignment == TextAlign::Justify)
    );
}

#[test]
fn test_rotation_setter_rejects_non_finite_values() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer.add_rectangle(slide, 0, 0, 100, 100).unwrap();

    assert!(writer.set_last_shape_rotation(slide, f32::NAN).is_err());
    assert_eq!(writer.slides[slide].shapes[0].properties.rotation, 0.0);
    assert!(
        writer
            .set_last_shape_text_alignment(slide, TextAlignment::Center)
            .is_err()
    );
}

#[test]
fn test_shape_adjustment_setter_preserves_sparse_positions() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_styled_shape(
            slide,
            ShapeType::Arrow,
            0,
            0,
            200,
            100,
            ShapeStyle::default(),
        )
        .unwrap();

    writer.set_last_shape_adjustment(slide, 3, -42).unwrap();
    assert_eq!(
        writer.slides[slide].shapes[0].properties.adjust_values,
        [0, 0, 0, -42]
    );
    let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
    assert_eq!(shape.adjust_values, [0, 0, 0, -42]);

    assert!(writer.set_last_shape_adjustment(slide, 10, 7).is_err());
    assert_eq!(
        writer.slides[slide].shapes[0].properties.adjust_values,
        [0, 0, 0, -42]
    );
}

#[test]
fn test_add_textbox_invalid_slide() {
    let mut writer = Writer::new();
    let result = writer.add_textbox(0, 10, 10, 100, 50, "Test");
    assert!(result.is_err());
}

#[test]
fn test_delete_slide() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    writer.add_slide().unwrap();
    writer.delete_slide(0).unwrap();
    assert_eq!(writer.slides.len(), 1);
}

#[test]
fn test_delete_invalid_slide() {
    let mut writer = Writer::new();
    let result = writer.delete_slide(0);
    assert!(result.is_err());
}

#[test]
fn test_add_styled_shape() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let style = ShapeStyle::solid_no_line(ShapeColor::RED);
    writer
        .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 50, style)
        .unwrap();
    assert_eq!(writer.slides[0].shapes.len(), 1);
}

#[test]
fn test_add_rectangle() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer.add_rectangle(slide, 10, 10, 100, 50).unwrap();
    assert_eq!(writer.slides[0].shapes.len(), 1);
}

#[test]
fn test_add_arrow_line() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer.add_arrow_line(slide, 0, 0, 100, 100).unwrap();
    assert_eq!(writer.slides[0].shapes.len(), 1);
}

#[test]
fn test_set_slide_notes() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .set_slide_notes(slide, "These are speaker notes")
        .unwrap();
    assert_eq!(
        writer.slides[0].notes,
        Some("These are speaker notes".to_string())
    );
}

#[test]
fn test_add_font() {
    let mut writer = Writer::new();
    // Writer::new() already adds Arial as default font at index 0
    let font = FontEntity::times_new_roman();
    let idx = writer.add_font(font);
    assert_eq!(idx, 1); // Second font at index 1
    assert_eq!(writer.font_count(), 2);
}

#[test]
fn test_add_multiple_fonts() {
    let mut writer = Writer::new();
    // Writer::new() already adds Arial as default font at index 0
    let idx1 = writer.add_font(FontEntity::arial()); // Returns 1
    let idx2 = writer.add_font(FontEntity::times_new_roman()); // Returns 2
    let idx3 = writer.add_font(FontEntity::new("Calibri")); // Returns 3
    assert_eq!(idx1, 1);
    assert_eq!(idx2, 2);
    assert_eq!(idx3, 3);
    assert_eq!(writer.font_count(), 4); // Arial (default) + 3 added
}

#[test]
fn test_set_property() {
    let mut writer = Writer::new();
    writer.set_property("Title", "My Presentation");
    writer.set_property("Author", "Test Author");
    assert_eq!(
        writer.properties.get("Title"),
        Some(&"My Presentation".to_string())
    );
    assert_eq!(
        writer.properties.get("Author"),
        Some(&"Test Author".to_string())
    );
}

#[test]
fn test_hyperlink_collection() {
    let mut writer = Writer::new();
    let link = Hyperlink::url("https://example.com").with_display_text("Example");
    let id = writer.add_hyperlink(link);
    assert_eq!(id, 1);
    assert_eq!(writer.hyperlink_count(), 1);
    assert!(writer.hyperlinks.get(1).is_some());
}

#[test]
fn maps_writer_hyperlinks_to_spec_link_targets() {
    let mut links = HyperlinkCollection::new();
    let slide = links.add(Hyperlink::slide(2));
    assert_eq!(get_hyperlink_info(Some(slide), &links), (4, 0, 7));

    let next = links.add(Hyperlink::next_slide());
    assert_eq!(get_hyperlink_info(Some(next), &links), (3, 1, 0));

    let custom = links.add(Hyperlink {
        id: 0,
        display_text: None,
        target: crate::writer::hyperlink::HyperlinkTarget::CustomShow("Demo".to_string()),
        target_frame: None,
    });
    assert_eq!(get_hyperlink_info(Some(custom), &links), (7, 0, 6));
}

#[test]
fn test_add_comment() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let comment = SlideComment::new("John Doe", "Great slide!", 100, 50);
    writer.add_comment(slide, comment).unwrap();
    assert_eq!(writer.slides[0].comments.len(), 1);
    assert_eq!(writer.slides[0].comments[0].author, "John Doe");
}

#[test]
fn test_add_multiple_comments() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_comment(slide, SlideComment::new("Alice", "First", 10, 10))
        .unwrap();
    writer
        .add_comment(slide, SlideComment::new("Bob", "Second", 20, 20))
        .unwrap();
    assert_eq!(writer.slides[0].comments.len(), 2);
}

#[test]
fn test_add_comment_invalid_slide() {
    let mut writer = Writer::new();
    let comment = SlideComment::new("John", "Test", 0, 0);
    let result = writer.add_comment(0, comment);
    assert!(result.is_err());
}

#[test]
fn test_shape_properties() {
    let props = ShapeProperties {
        shape_type: ShapeType::Rectangle,
        x: 100,
        y: 200,
        width: 300,
        height: 400,
        text: Some("Hello".to_string()),
        paragraphs: None,
        alignment: TextAlignment::Center,
        fill: None,
        line: None,
        shadow: None,
        rotation: 45.0,
        adjust_values: Vec::new(),
        flip_h: true,
        flip_v: false,
        hyperlink_id: None,
        interactions: Vec::new(),
        text_interactions: Vec::new(),
        picture_index: None,
        freeform_geometry: None,
    };
    assert_eq!(props.x, 100);
    assert_eq!(props.y, 200);
    assert_eq!(props.width, 300);
    assert_eq!(props.height, 400);
    assert!(props.flip_h);
    assert!(!props.flip_v);
}

#[test]
fn test_slide_count() {
    let mut writer = Writer::new();
    assert_eq!(writer.slide_count(), 0);
    writer.add_slide().unwrap();
    assert_eq!(writer.slide_count(), 1);
    writer.add_slide().unwrap();
    assert_eq!(writer.slide_count(), 2);
}

#[test]
fn test_default_writer() {
    let writer: Writer = Default::default();
    assert_eq!(writer.slide_count(), 0);
    assert_eq!(writer.slide_width, 9144000);
    assert_eq!(writer.slide_height, 6858000);
}

#[test]
fn test_ppt_write_error_display() {
    let io_err = WriteError::Io(std::io::Error::other("test error"));
    let err_str = format!("{}", io_err);
    assert!(err_str.contains("I/O error"));

    let data_err = WriteError::InvalidData("bad data".to_string());
    let err_str = format!("{}", data_err);
    assert!(err_str.contains("Invalid data"));
}

#[test]
fn test_text_alignment_conversions() {
    assert_eq!(TextAlignment::Left as u8, 0);
    assert_eq!(TextAlignment::Center as u8, 1);
    assert_eq!(TextAlignment::Right as u8, 2);
    assert_eq!(TextAlignment::Justify as u8, 3);
}

#[test]
fn test_slide_layout_types() {
    use super::super::spec::SlideLayoutType;
    assert_eq!(SlideLayoutType::TitleSlide as u32, 0);
    assert_eq!(SlideLayoutType::TitleBody as u32, 1);
    assert_eq!(SlideLayoutType::MasterTitle as u32, 2);
    assert_eq!(SlideLayoutType::TitleOnly as u32, 7);
    assert_eq!(SlideLayoutType::Blank as u32, 13);
}

#[test]
fn test_shape_type_variants() {
    let types = vec![
        ShapeType::Rectangle,
        ShapeType::TextBox,
        ShapeType::Placeholder,
        ShapeType::Line,
        ShapeType::Ellipse,
        ShapeType::RoundRectangle,
        ShapeType::Diamond,
        ShapeType::Triangle,
        ShapeType::Arrow,
        ShapeType::Star,
        ShapeType::Heart,
        ShapeType::Picture,
    ];
    for shape_type in types {
        let _ = format!("{:?}", shape_type);
    }
}

#[test]
fn test_write_to_memory() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 100, 100, 400, 200, "Hello, World!")
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
    assert!(!buffer.get_ref().is_empty());
}

#[test]
fn smart_tags_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let tokyo = writer
        .add_smart_tag(
            SmartTagDefinition::new("urn:example:geo", "place").with_property("city", "東京"),
        )
        .unwrap();
    let paris = writer
        .add_smart_tag(
            SmartTagDefinition::new("urn:example:geo", "place").with_property("city", "Paris"),
        )
        .unwrap();
    writer
        .add_rich_textbox(
            slide,
            10,
            10,
            300,
            100,
            vec![Paragraph::with_runs(vec![
                super::super::text_format::TextRun::new("Tokyo").with_smart_tag(tokyo),
                super::super::text_format::TextRun::new(" and "),
                super::super::text_format::TextRun::new("Paris").with_smart_tags([tokyo, paris]),
            ])],
        )
        .unwrap();
    let hyperlink = writer.add_hyperlink(Hyperlink::url("https://example.invalid/"));
    writer.set_last_shape_hyperlink(slide, hyperlink).unwrap();
    let encoded_shape =
        convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
    assert_eq!(
        encoded_shape.paragraphs.as_ref().unwrap()[0]
            .runs
            .iter()
            .map(|run| run.style.pp9_run_id)
            .collect::<Vec<_>>(),
        vec![Some(0), Some(1), Some(2)]
    );

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let presentation = package.presentation().unwrap();
    let store = presentation.smart_tags().unwrap().unwrap();
    assert_eq!(store.types.len(), 1);
    assert_eq!(store.tags.len(), 2);
    assert_eq!(store.tags[0].properties[0].value, "東京");
    let shape_tags = presentation.shape_programmable_tags().unwrap();
    assert_eq!(shape_tags.len(), 1);
    let tags = &shape_tags[0].programmable_tags;
    assert_eq!(tags.powerpoint9().unwrap().runs.len(), 3);
    assert_eq!(
        tags.powerpoint11().unwrap().runs[0].smart_tag_indices,
        vec![tokyo.as_u32()]
    );
    assert!(
        tags.powerpoint11().unwrap().runs[1]
            .smart_tag_indices
            .is_empty()
    );
    assert_eq!(
        tags.powerpoint11().unwrap().runs[2].smart_tag_indices,
        vec![tokyo.as_u32(), paris.as_u32()]
    );
    drop(presentation);
    drop(package);

    let path = std::env::temp_dir().join(format!(
        "litchi-ppt-smart-tags-{}-{}.ppt",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    let presentation = package.presentation().unwrap();
    assert_eq!(presentation.smart_tags().unwrap(), Some(store));
    assert_eq!(
        presentation.shape_programmable_tags().unwrap()[0]
            .programmable_tags
            .powerpoint11()
            .unwrap()
            .runs[2]
            .smart_tag_indices,
        vec![tokyo.as_u32(), paris.as_u32()]
    );
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rejects_missing_smart_tag_references() {
    let mut source = Writer::new();
    let dangling = source
        .add_smart_tag(SmartTagDefinition::new("urn:test", "dangling"))
        .unwrap();
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_rich_textbox(
            slide,
            0,
            0,
            100,
            50,
            vec![Paragraph::with_runs(vec![
                super::super::text_format::TextRun::new("invalid").with_smart_tag(dangling),
            ])],
        )
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());

    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let tag = writer
        .add_smart_tag(SmartTagDefinition::new("urn:test", "empty"))
        .unwrap();
    writer
        .add_rich_textbox(
            slide,
            0,
            0,
            100,
            50,
            vec![Paragraph::with_runs(vec![
                super::super::text_format::TextRun::new("").with_smart_tag(tag),
            ])],
        )
        .unwrap();
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_err());
}

#[test]
fn test_write_empty_presentation() {
    let mut writer = Writer::new();
    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
    assert!(!buffer.get_ref().is_empty());
}

#[test]
fn test_write_multiple_slides() {
    let mut writer = Writer::new();

    let slide1 = writer.add_slide().unwrap();
    writer
        .add_textbox(slide1, 100, 100, 400, 100, "Slide 1")
        .unwrap();

    let slide2 = writer.add_slide().unwrap();
    writer.add_ellipse(slide2, 100, 100, 200, 150).unwrap();

    let slide3 = writer.add_slide().unwrap();
    writer.add_line(slide3, 0, 0, 500, 500).unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
    assert!(!buffer.get_ref().is_empty());
}

#[test]
fn test_presentation_with_hyperlink() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    writer
        .add_textbox(slide, 100, 100, 300, 100, "Click here")
        .unwrap();

    let link = Hyperlink::url("https://example.com");
    let link_id = writer.add_hyperlink(link);
    writer.set_last_shape_hyperlink(slide, link_id).unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_presentation_with_comments() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 100, 100, 400, 200, "Content")
        .unwrap();

    let comment = SlideComment::new("Reviewer", "Please update this", 150, 150);
    writer.add_comment(slide, comment).unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_presentation_with_notes() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 100, 100, 400, 200, "Title")
        .unwrap();
    writer
        .set_slide_notes(slide, "These are speaker notes for this slide")
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_presentation_with_multiple_shapes() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    writer.add_rectangle(slide, 50, 50, 100, 100).unwrap();
    writer.add_line(slide, 50, 200, 300, 200).unwrap();
    writer
        .add_textbox(slide, 50, 300, 300, 100, "Text box content")
        .unwrap();

    assert_eq!(writer.slides[0].shapes.len(), 3);

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_custom_show_support() {
    let mut writer = Writer::new();
    writer.add_slide().unwrap();
    writer.add_slide().unwrap();
    writer.add_slide().unwrap();

    let custom_show = CustomShow::new("Important Slides", &[0, 2]);
    writer.add_custom_show(custom_show);

    assert_eq!(writer.custom_show_count(), 1);
}

#[test]
fn test_multiple_custom_shows() {
    let mut writer = Writer::new();
    for _ in 0..5 {
        writer.add_slide().unwrap();
    }

    writer.add_custom_show(CustomShow::new("First Show", &[0usize, 1]));
    writer.add_custom_show(CustomShow::new("Second Show", &[2usize, 3, 4]));

    assert_eq!(writer.custom_show_count(), 2);
}

#[test]
fn test_widescreen_write() {
    let mut writer = Writer::new_widescreen();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 100, 100, 800, 100, "Widescreen")
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_shape_with_styling() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();

    let style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(255, 0, 0))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            2.0,
        ));

    writer
        .add_styled_shape(slide, ShapeType::Rectangle, 100, 100, 200, 150, style)
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    let result = writer.write_to(&mut buffer);
    assert!(result.is_ok());
}

#[test]
fn test_extended_line_style_reaches_escher_shape() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let mut line = LineStyleConfig::with_color_and_width(ShapeColor::RED, 2.0);
    line.opacity = 50;
    line.style = LineStyle::Triple;
    line.cap = LineCapStyle::Flat;
    line.join = LineJoinStyle::Round;
    line.start_arrow = ArrowStyle::Triangle;
    line.start_arrow_width = ArrowSize::Small;
    line.start_arrow_length = ArrowSize::Large;
    line.end_arrow = ArrowStyle::Open;
    line.end_arrow_width = ArrowSize::Large;
    line.end_arrow_length = ArrowSize::Small;
    let style = ShapeStyle::new().with_line(line);

    writer
        .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 100, style)
        .unwrap();
    let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);

    assert_eq!(shape.line_opacity, Some(32768));
    assert_eq!(shape.line_style, Some(LineStyle::Triple as u32));
    assert_eq!(shape.line_end_cap_style, Some(LineCapStyle::Flat as u32));
    assert_eq!(shape.line_join_style, Some(LineJoinStyle::Round as u32));
    assert_eq!(shape.line_start_arrow_width, Some(ArrowSize::Small as u32));
    assert_eq!(shape.line_start_arrow_length, Some(ArrowSize::Large as u32));
    assert_eq!(shape.line_end_arrow_width, Some(ArrowSize::Large as u32));
    assert_eq!(shape.line_end_arrow_length, Some(ArrowSize::Small as u32));

    let default_line = convert_line_properties(Some(&LineStyleConfig::default_line()));
    assert_eq!(default_line.opacity, None);
    assert_eq!(default_line.style, None);
    assert_eq!(default_line.end_cap_style, None);
    assert_eq!(default_line.join_style, None);
}

#[test]
fn test_picture_fill_registers_and_serializes_blip_reference() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let blip_index = writer
        .add_picture_data_as(
            vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a],
            PictureKind::Png,
        )
        .unwrap();
    let style = ShapeStyle::new().with_fill(FillStyle::picture(blip_index));
    writer
        .add_styled_shape(slide, ShapeType::Rectangle, 10, 10, 100, 100, style)
        .unwrap();

    let shape = convert_shape_to_escher(&writer.slides[slide].shapes[0], &writer.hyperlinks);
    assert_eq!(shape.fill_type, Some(3));
    assert_eq!(shape.fill_blip_index, Some(1));
    assert_eq!(writer.picture_count(), 1);

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    assert!(!output.into_inner().is_empty());
}

#[test]
fn test_invalid_slide_picture_does_not_mutate_blip_store() {
    let mut writer = Writer::new();

    assert!(
        writer
            .add_picture(7, 0, 0, 100, 100, vec![0x89, b'P', b'N', b'G'])
            .is_err()
    );
    assert_eq!(writer.picture_count(), 0);
}

#[test]
fn test_invalid_operations() {
    let mut writer = Writer::new();

    // Try to add shape to non-existent slide
    let result = writer.add_rectangle(0, 0, 0, 100, 100);
    assert!(result.is_err());

    // Try to add textbox to non-existent slide
    let result = writer.add_textbox(5, 10, 10, 100, 50, "Test");
    assert!(result.is_err());

    // Try to set notes on non-existent slide
    let result = writer.set_slide_notes(0, "Notes");
    assert!(result.is_err());
}

#[test]
fn test_internal_slide_data() {
    let mut writer = Writer::new();
    let slide_idx = writer.add_slide().unwrap();

    // Verify slide was created with correct defaults
    let slide = &writer.slides[slide_idx];
    assert!(slide.shapes.is_empty());
    assert!(slide.notes.is_none());
    assert!(slide.comments.is_empty());
}

#[test]
fn test_slide_persist_tracking() {
    let mut writer = Writer::new();

    let idx1 = writer.add_slide().unwrap();
    let idx2 = writer.add_slide().unwrap();

    // Each slide gets a persist ID assigned during writing
    // (We can't check this directly, but we verify the structure)
    assert_eq!(idx1, 0);
    assert_eq!(idx2, 1);
}
