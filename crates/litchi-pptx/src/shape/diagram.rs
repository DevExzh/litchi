//! Presentation-specific SmartArt semantics.
//!
//! The diagram model and its five part codecs live in `litchi-drawingml`.
//! This focused module supplies the PresentationML graphic-frame anchor and
//! gives the shared model contextual names (`Graphic`, `Node`, `Kind`, and
//! `Builder`) without copying the diagram payload.

pub use litchi_drawingml::diagram::{
    DiagramNode as Node, DiagramType as Kind, SmartArt as Graphic, SmartArtBuilder as Builder,
    generate_smartart_colors_xml as colors_xml, generate_smartart_data_xml as data_xml,
    generate_smartart_drawing_xml as drawing_xml, generate_smartart_layout_xml as layout_xml,
    generate_smartart_quickstyle_xml as quickstyle_xml,
};

/// Generate a PresentationML graphic frame for a diagram.
///
/// Relationship IDs are escaped before insertion. The data relationship is
/// expected to use the usual `rIdN` spelling; the non-numeric prefix is used
/// for the companion layout, quick-style, and color relationships.
pub fn graphic_frame(
    shape_id: u32,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    data_rel_id: &str,
) -> String {
    let mut xml = String::with_capacity(1024 + data_rel_id.len() * 4);
    let data_rel_id = escape_attribute(data_rel_id);
    let prefix_len = data_rel_id
        .trim_end_matches(|character: char| character.is_ascii_digit())
        .len();
    let base_rel_id = &data_rel_id[..prefix_len.max(1).min(data_rel_id.len())];

    xml.push_str(
        "<p:graphicFrame xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
    );
    xml.push_str("<p:nvGraphicFramePr><p:cNvPr id=\"");
    xml.push_str(&shape_id.to_string());
    xml.push_str("\" name=\"SmartArt ");
    xml.push_str(&shape_id.to_string());
    xml.push_str("\"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>");
    xml.push_str(
        "<p:xfrm><a:off xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" x=\"",
    );
    xml.push_str(&x.to_string());
    xml.push_str("\" y=\"");
    xml.push_str(&y.to_string());
    xml.push_str(
        "\"/><a:ext xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" cx=\"",
    );
    xml.push_str(&width.to_string());
    xml.push_str("\" cy=\"");
    xml.push_str(&height.to_string());
    xml.push_str("\"/></p:xfrm>");
    xml.push_str("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">");
    xml.push_str("<dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:dm=\"");
    xml.push_str(&data_rel_id);
    xml.push_str("\" r:lo=\"");
    xml.push_str(base_rel_id);
    xml.push_str("lo\" r:qs=\"");
    xml.push_str(base_rel_id);
    xml.push_str("qs\" r:cs=\"");
    xml.push_str(base_rel_id);
    xml.push_str("cs\"/></a:graphicData></a:graphic></p:graphicFrame>");
    xml
}

fn escape_attribute(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            character => escaped.push(character),
        }
    }
    escaped
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn contextual_model_and_frame_are_available() {
        assert_eq!(
            Kind::from_layout_uri("urn:microsoft.com/office/list"),
            Kind::List
        );
        let graphic = Builder::new(Kind::List).add_item("Item").build();
        assert_eq!(graphic.node_count(), 1);
        let xml = graphic_frame(5, 1000, 2000, 4000, 3000, "rId10");
        assert!(xml.contains("r:dm=\"rId10\""));
        assert!(xml.contains("r:lo=\"rIdlo\""));
    }

    #[test]
    fn relationship_id_is_escaped_without_panicking_on_empty_input() {
        let xml = graphic_frame(1, 0, 0, 1, 1, "r&Id");
        assert!(xml.contains("r:dm=\"r&amp;Id\""));
    }
}
