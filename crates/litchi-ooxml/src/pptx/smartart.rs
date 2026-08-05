//! Migration adapter for the standalone PresentationML diagram owner.

pub use litchi_pptx::shape::diagram::{
    Builder as SmartArtBuilder, Graphic as SmartArt, Kind as DiagramType, Node as DiagramNode,
    colors_xml as generate_smartart_colors_xml, data_xml as generate_smartart_data_xml,
    drawing_xml as generate_smartart_drawing_xml, layout_xml as generate_smartart_layout_xml,
    quickstyle_xml as generate_smartart_quickstyle_xml,
};

/// Generate a PresentationML graphic frame for a diagram.
pub fn generate_smartart_graphic_frame(
    shape_id: u32,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    data_rel_id: &str,
) -> String {
    litchi_pptx::shape::diagram::graphic_frame(shape_id, x, y, width, height, data_rel_id)
}
