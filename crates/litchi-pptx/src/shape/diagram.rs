//! Presentation-specific `SmartArt` semantics.
//!
//! The diagram model and its five part codecs live in `litchi-drawingml`.
//! This focused module supplies the `PresentationML` graphic-frame anchor and
//! gives the shared model contextual names (`Graphic`, `Node`, `Kind`, and
//! `Builder`) without copying the diagram payload.

pub use litchi_drawingml::diagram::{
    DiagramNode as Node, DiagramType as Kind, Limits, SmartArt as Graphic,
    SmartArtBuilder as Builder, generate_smartart_colors_xml as colors_xml,
    generate_smartart_data_xml as data_xml, generate_smartart_drawing_xml as drawing_xml,
    generate_smartart_layout_xml as layout_xml, generate_smartart_quickstyle_xml as quickstyle_xml,
};

/// Generate a `PresentationML` graphic frame for a diagram.
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
    limits: &Limits,
) -> litchi_drawingml::Result<String> {
    limits.validate()?;
    let max_input = limits.input_bytes_value().min(16 * 1024 * 1024);
    if data_rel_id.len() > max_input {
        return Err(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::InputBytes,
            observed: data_rel_id.len(),
            maximum: max_input,
        });
    }
    let escaped_len = data_rel_id.bytes().try_fold(0_usize, |length, byte| {
        let extra = match byte {
            b'&' => 4,
            b'<' | b'>' => 3,
            b'"' | b'\'' => 5,
            _ => 0,
        };
        length.checked_add(1 + extra).ok_or_else(|| {
            litchi_drawingml::Error::Invalid("diagram relationship ID length overflows".to_owned())
        })
    })?;
    let estimate = 2048_usize
        .checked_add(escaped_len.checked_mul(4).ok_or_else(|| {
            litchi_drawingml::Error::Invalid("diagram graphic-frame length overflows".to_owned())
        })?)
        .ok_or_else(|| {
            litchi_drawingml::Error::Invalid("diagram graphic-frame length overflows".to_owned())
        })?;
    let max_output = limits.output_bytes_value().min(16 * 1024 * 1024);
    if estimate > max_output {
        return Err(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::DiagramDrawingBytes,
            observed: estimate,
            maximum: max_output,
        });
    }
    let mut xml = String::with_capacity(estimate);
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
    if xml.len() > max_output {
        return Err(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::DiagramDrawingBytes,
            observed: xml.len(),
            maximum: max_output,
        });
    }
    Ok(xml)
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
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
        let xml = graphic_frame(5, 1000, 2000, 4000, 3000, "rId10", &Limits::default()).unwrap();
        assert!(xml.contains("r:dm=\"rId10\""));
        assert!(xml.contains("r:lo=\"rIdlo\""));
    }

    #[test]
    fn relationship_id_is_escaped_without_panicking_on_empty_input() {
        let xml = graphic_frame(1, 0, 0, 1, 1, "r&Id", &Limits::default()).unwrap();
        assert!(xml.contains("r:dm=\"r&amp;Id\""));
    }

    #[test]
    fn custom_frame_values_are_preserved_and_output_is_bounded() {
        let xml = graphic_frame(77, -10, 20, 300, 400, "rId42", &Limits::default()).unwrap();
        assert!(xml.contains("id=\"77\""));
        assert!(xml.contains("x=\"-10\" y=\"20\""));
        assert!(xml.contains("cx=\"300\" cy=\"400\""));
        assert!(xml.contains("r:lo=\"rIdlo\""));

        let defaults = Limits::default();
        let tiny = Limits::new(
            defaults.input_bytes_value(),
            1,
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            defaults.points_value(),
            defaults.connections_value(),
            defaults.data_text_bytes_value(),
            defaults.nodes_value(),
            defaults.node_depth_value(),
            defaults.node_text_bytes_value(),
        );
        assert!(matches!(
            graphic_frame(77, 0, 0, 1, 1, "rId1", &tiny),
            Err(litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::DiagramDrawingBytes,
                ..
            })
        ));
    }
}
