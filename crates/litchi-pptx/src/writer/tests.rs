#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::MutablePresentation;
use crate::shape::designer::{
    DrawingProperties, Limits, PROPERTIES_EXTENSION_URI, TAGS_EXTENSION_URI, Tag, Tags,
};

#[test]
fn shape_designer_properties_round_trip_for_text_and_preset_shapes() {
    let mut duplicate_tags = Tags::new();
    duplicate_tags
        .push(Tag::new("role&\"<", "hero").unwrap())
        .unwrap();
    duplicate_tags
        .push(Tag::new("role&\"<", "hero").unwrap())
        .unwrap();
    let explicit_false = DrawingProperties::new()
        .with_editable(Some(false))
        .with_tags(Some(duplicate_tags.clone()), Limits::default())
        .unwrap();

    let mut presentation = MutablePresentation::new();
    let slide = presentation.add_slide().unwrap();
    slide
        .add_text_box("Designer text", 0, 0, 100, 100)
        .set_designer_properties(explicit_false.clone())
        .unwrap();
    slide
        .add_rectangle_with_designer_properties(0, 0, 100, 100, None, DrawingProperties::new())
        .unwrap();
    slide.add_ellipse(0, 0, 100, 100, None);

    let xml = slide.generate_slide_xml().unwrap();
    assert_eq!(xml.matches(PROPERTIES_EXTENSION_URI).count(), 2);
    assert_eq!(xml.matches("<p202:designPr").count(), 2);
    assert_eq!(xml.matches("<p:nvPr/>").count(), 2); // group owner plus absent ellipse
    assert_eq!(
        xml.matches("xmlns:p202=\"http://schemas.microsoft.com/office/powerpoint/2020/02/main\"")
            .count(),
        2
    );

    let payloads = elements(&xml, "<p202:designPr", "</p202:designPr>");
    let text = crate::shape::designer::read_properties(payloads[0].as_bytes(), Limits::default())
        .unwrap()
        .value;
    let rectangle =
        crate::shape::designer::read_properties(payloads[1].as_bytes(), Limits::default())
            .unwrap()
            .value;
    assert_eq!(text, explicit_false);
    assert_eq!(text.editable(), Some(false));
    assert_eq!(text.tags().unwrap(), &duplicate_tags);
    assert_eq!(rectangle.editable(), None);
    assert_eq!(rectangle.tags(), None);

    let property_extension =
        format!("<p:nvPr><p:extLst><p:ext uri=\"{PROPERTIES_EXTENSION_URI}\"><p202:designPr");
    assert_eq!(xml.matches(&property_extension).count(), 2);
    let explicit_false_position = xml.find("edtDesignElem=\"false\"").unwrap();
    let embedded_tags_position = xml.find("<p202:designTagLst>").unwrap();
    assert!(explicit_false_position < embedded_tags_position);

    let tight = Limits::default().with_xml_bytes(1);
    let ellipse = slide.shapes.get_mut(2).unwrap();
    assert!(
        ellipse
            .set_designer_properties_with_limits(DrawingProperties::new(), tight)
            .is_err()
    );
    assert!(ellipse.designer_properties().is_none());
}

#[test]
fn designer_writer_clear_paths_and_limit_changes_are_semantic_noops() {
    let properties = DrawingProperties::new().with_editable(Some(false));
    let mut tags = Tags::new();
    tags.push(Tag::new("", "").unwrap()).unwrap();
    tags.push(Tag::new("duplicate", "first").unwrap()).unwrap();
    tags.push(Tag::new("duplicate", "second").unwrap()).unwrap();

    let mut presentation = MutablePresentation::new();
    let slide = presentation.add_slide().unwrap();
    slide
        .add_ellipse_with_designer_properties(0, 0, 100, 100, None, properties.clone())
        .unwrap();
    slide.set_designer_tags(tags.clone()).unwrap();

    let initial_slide = slide.generate_slide_xml().unwrap();
    assert!(initial_slide.contains("edtDesignElem=\"false\""));
    let initial_presentation = presentation.generate_presentation_xml().unwrap();
    assert!(initial_presentation.contains(&format!(
        "<p:sldId id=\"256\" r:id=\"rId2\"><p:extLst><p:ext uri=\"{TAGS_EXTENSION_URI}\"><p202:designTagLst"
    )));
    let duplicate_positions = [
        initial_presentation.find("name=\"\" val=\"\"").unwrap(),
        initial_presentation
            .find("name=\"duplicate\" val=\"first\"")
            .unwrap(),
        initial_presentation
            .find("name=\"duplicate\" val=\"second\"")
            .unwrap(),
    ];
    assert!(duplicate_positions[0] < duplicate_positions[1]);
    assert!(duplicate_positions[1] < duplicate_positions[2]);

    presentation.mark_clean();
    let larger = Limits::default().with_xml_bytes(16 * 1024 * 1024);
    let slide = presentation.slide_mut(0).unwrap();
    slide
        .shapes
        .first_mut()
        .unwrap()
        .set_designer_properties_with_limits(properties, larger)
        .unwrap();
    slide.set_designer_tags_with_limits(tags, larger).unwrap();
    assert!(!presentation.is_modified());

    let slide = presentation.slide_mut(0).unwrap();
    assert!(
        slide
            .shapes
            .first_mut()
            .unwrap()
            .clear_designer_properties()
    );
    assert!(slide.clear_designer_tags());
    let cleared_slide = slide.generate_slide_xml().unwrap();
    assert!(!cleared_slide.contains(PROPERTIES_EXTENSION_URI));
    let cleared_presentation = presentation.generate_presentation_xml().unwrap();
    assert!(!cleared_presentation.contains(TAGS_EXTENSION_URI));
}

#[test]
fn slide_designer_tags_follow_reorder_duplicate_delete_and_preserve_empty() {
    let mut duplicate_tags = Tags::new();
    duplicate_tags
        .push(Tag::new("kind", "same&\"").unwrap())
        .unwrap();
    duplicate_tags
        .push(Tag::new("kind", "same&\"").unwrap())
        .unwrap();

    let mut presentation = MutablePresentation::new();
    presentation
        .add_slide()
        .unwrap()
        .set_designer_tags(duplicate_tags.clone())
        .unwrap(); // 256
    presentation.add_slide().unwrap(); // 257, absent
    presentation
        .add_slide()
        .unwrap()
        .set_designer_tags(Tags::new())
        .unwrap(); // 258, explicitly empty

    presentation.move_slide(0, 2).unwrap();
    let moved = presentation.generate_presentation_xml().unwrap();
    assert!(moved.contains("<p:sldId id=\"257\" r:id=\"rId2\"/>"));
    assert!(moved.contains("<p:sldId id=\"258\" r:id=\"rId3\"><p:extLst>"));
    assert!(moved.contains("<p:sldId id=\"256\" r:id=\"rId4\"><p:extLst>"));
    assert_eq!(moved.matches(TAGS_EXTENSION_URI).count(), 2);

    let moved_payloads = elements(&moved, "<p202:designTagLst", "</p202:designTagLst>");
    let empty =
        crate::shape::designer::read_tags(moved_payloads[0].as_bytes(), Limits::default()).unwrap();
    let duplicate =
        crate::shape::designer::read_tags(moved_payloads[1].as_bytes(), Limits::default()).unwrap();
    assert!(empty.is_empty());
    assert_eq!(duplicate, duplicate_tags);

    presentation.duplicate_slide(2).unwrap(); // clone tags to new stable ID 259
    presentation.delete_slide(2).unwrap(); // delete original tagged slide
    let final_xml = presentation.generate_presentation_xml().unwrap();
    assert!(!final_xml.contains("<p:sldId id=\"256\""));
    assert!(final_xml.contains("<p:sldId id=\"259\" r:id=\"rId4\"><p:extLst>"));
    assert_eq!(final_xml.matches(TAGS_EXTENSION_URI).count(), 2);
    let final_payloads = elements(&final_xml, "<p202:designTagLst", "</p202:designTagLst>");
    assert!(
        crate::shape::designer::read_tags(final_payloads[0].as_bytes(), Limits::default())
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        crate::shape::designer::read_tags(final_payloads[1].as_bytes(), Limits::default()).unwrap(),
        duplicate_tags
    );
}

fn elements<'a>(xml: &'a str, start_marker: &str, end_marker: &str) -> Vec<&'a str> {
    let mut output = Vec::new();
    let mut offset = 0usize;
    while let Some(relative_start) = xml[offset..].find(start_marker) {
        let start = offset + relative_start;
        let start_close = start + xml[start..].find('>').expect("generated start tag") + 1;
        let end = if xml.as_bytes()[start_close - 2] == b'/' {
            start_close
        } else {
            start_close
                + xml[start_close..]
                    .find(end_marker)
                    .expect("generated end tag")
                + end_marker.len()
        };
        output.push(&xml[start..end]);
        offset = end;
    }
    output
}
