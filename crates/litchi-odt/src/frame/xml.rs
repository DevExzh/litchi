use super::Frame;
use crate::elements::element::{Element, ElementBase};
use litchi_core::Result;
use litchi_odf_common::drawing::authoring::validate_text_box;

/// Build the ODT `draw:frame` element for a package-linked image.
pub(crate) fn image_element(frame: &Frame, href: &str) -> Element {
    let mut frame_element = frame_shell(frame);
    let mut image = Element::new("draw:image");
    image.set_attribute("xlink:href", href);
    image.set_attribute("xlink:type", "simple");
    image.set_attribute("xlink:show", "embed");
    image.set_attribute("xlink:actuate", "onLoad");
    frame_element.add_child(image);
    frame_element
}

/// Build the ODT `draw:frame` element for a plain-text text box.
pub(crate) fn text_box_element(frame: &Frame, text: &str) -> Result<Element> {
    validate_text_box(text)?;
    let mut frame_element = frame_shell(frame);
    let mut text_box = Element::new("draw:text-box");
    if text.is_empty() {
        text_box.add_child(Element::new("text:p"));
    } else {
        for line in text.split('\n') {
            let mut paragraph = Element::new("text:p");
            if !line.is_empty() {
                paragraph.set_text(line);
            }
            text_box.add_child(paragraph);
        }
    }
    frame_element.add_child(text_box);
    Ok(frame_element)
}

fn frame_shell(frame: &Frame) -> Element {
    let mut element = Element::new("draw:frame");
    element.set_attribute("draw:name", frame.name());
    element.set_attribute("text:anchor-type", frame.anchor().as_str());
    element.set_attribute("svg:width", frame.width().as_str());
    element.set_attribute("svg:height", frame.height().as_str());
    element
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::elements::element::ElementBase;
    use crate::frame::Anchor;

    fn make_frame(anchor: Anchor) -> Frame {
        Frame::new(
            "Photo",
            super::super::Length::centimeters(10.0),
            super::super::Length::centimeters(4.0),
            anchor,
        )
        .unwrap()
    }

    #[test]
    fn image_frame_carries_identity_anchor_geometry_and_link() {
        let element = image_element(&make_frame(Anchor::AsChar), "Pictures/image1.png");
        assert_eq!(element.get_attribute("draw:name"), Some("Photo"));
        assert_eq!(element.get_attribute("text:anchor-type"), Some("as-char"));
        assert_eq!(element.get_attribute("svg:width"), Some("10cm"));
        let image = &element.get_children()[0];
        assert_eq!(image.tag_name(), "draw:image");
        assert_eq!(
            image.get_attribute("xlink:href"),
            Some("Pictures/image1.png")
        );
        let xml = element.to_xml_string();
        assert!(xml.contains("xlink:actuate=\"onLoad\""));
    }

    #[test]
    fn text_box_frame_splits_lines_and_escapes_text() {
        let element =
            text_box_element(&make_frame(Anchor::Paragraph), "a < b\nsecond & line").unwrap();
        let text_box = &element.get_children()[0];
        assert_eq!(text_box.tag_name(), "draw:text-box");
        assert_eq!(text_box.get_children().len(), 2);
        let xml = element.to_xml_string();
        assert!(xml.contains("a &lt; b"));
        assert!(xml.contains("second &amp; line"));
        assert!(text_box_element(&make_frame(Anchor::Page), "").is_ok());
    }
}
