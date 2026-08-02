//! SmartArt/Diagram support for PowerPoint presentations.
//!
//! SmartArt graphics are represented as diagrams in OOXML. The semantic model
//! (`SmartArt`, `DiagramNode`, `DiagramType`), the builder, and the diagram
//! part generators live in the format-agnostic
//! [`litchi_drawingml::diagram`] module and are re-exported here for the
//! PowerPoint-specific facade.
//! This module adds the PowerPoint-specific graphic-frame anchor.

pub use litchi_drawingml::diagram::{
    DiagramNode, DiagramType, SmartArt, SmartArtBuilder, generate_smartart_colors_xml,
    generate_smartart_data_xml, generate_smartart_drawing_xml, generate_smartart_layout_xml,
    generate_smartart_quickstyle_xml,
};

/// Generate graphic frame XML for embedding SmartArt on a slide.
pub fn generate_smartart_graphic_frame(
    shape_id: u32,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    data_rel_id: &str,
) -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(
        "<p:graphicFrame xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">",
    );
    xml.push_str("<p:nvGraphicFramePr>");
    xml.push_str(&format!(
        r#"<p:cNvPr id="{}" name="SmartArt {}"/>"#,
        shape_id, shape_id
    ));
    xml.push_str(r#"<p:cNvGraphicFramePr/>"#);
    xml.push_str("<p:nvPr/>");
    xml.push_str("</p:nvGraphicFramePr>");

    xml.push_str("<p:xfrm>");
    xml.push_str(&format!(
        r#"<a:off xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="{}" y="{}"/>"#,
        x, y
    ));
    xml.push_str(&format!(r#"<a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" cx="{}" cy="{}"/>"#, width, height));
    xml.push_str("</p:xfrm>");

    xml.push_str(r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
    xml.push_str(
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#,
    );
    let base_id = &data_rel_id[..data_rel_id.len() - 2];
    xml.push_str(&format!(
        r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="{}" r:lo="{}lo" r:qs="{}qs" r:cs="{}cs"/>"#,
        data_rel_id, base_id, base_id, base_id
    ));
    xml.push_str("</a:graphicData>");
    xml.push_str("</a:graphic>");

    xml.push_str("</p:graphicFrame>");

    xml
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_diagram_type_from_uri() {
        assert_eq!(
            DiagramType::from_layout_uri("urn:microsoft.com/office/list"),
            DiagramType::List
        );
        assert_eq!(
            DiagramType::from_layout_uri("urn:microsoft.com/office/orgChart"),
            DiagramType::Hierarchy
        );
    }

    #[test]
    fn test_smartart_builder() {
        let smartart = SmartArtBuilder::new(DiagramType::List)
            .layout_name("Basic List")
            .add_items(vec!["Item 1", "Item 2", "Item 3"])
            .build();

        assert_eq!(smartart.diagram_type, DiagramType::List);
        assert_eq!(smartart.node_count(), 3);
        assert!(smartart.text().contains("Item 1"));
    }

    #[test]
    fn test_generate_graphic_frame() {
        let xml = generate_smartart_graphic_frame(5, 1000, 2000, 4000, 3000, "rId10");
        assert!(xml.contains("<p:graphicFrame"));
        assert!(xml.contains(r#"r:dm="rId10""#));
        assert!(xml.contains("dgm:relIds"));
        assert!(xml.contains("</p:graphicFrame>"));
    }
}
