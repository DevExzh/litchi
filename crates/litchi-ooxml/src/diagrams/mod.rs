//! DrawingML diagram (SmartArt) support shared across OOXML formats.
//!
//! A SmartArt graphic is stored as a set of diagram parts — the data model
//! (`dataN.xml`), layout (`layoutN.xml`), quick style (`quickStyleN.xml`),
//! colors (`colorsN.xml`), and an optional pre-rendered drawing
//! (`drawingN.xml`) — anchored from the host document through a
//! `dgm:relIds` (Word) or graphic-frame (PowerPoint) reference.
//!
//! This module holds the format-agnostic pieces:
//!
//! - [`model`]: the semantic `SmartArt`/`DiagramNode` model, its builder, and
//!   the generators for the five diagram parts.
//! - [`data`]: a bounded, inert parser for the `dgm:dataModel` node and
//!   connection graph.
//! - [`definition`]: inert header metadata for the layout, quick-style, and
//!   colors definition parts.
//!
//! Both the transitional and the ISO Strict `drawingml/diagram` namespaces are
//! supported. Format-specific anchoring lives in `crate::docx::smartart` and
//! `crate::pptx::smartart`.

pub mod data;
pub mod definition;
pub mod model;

pub use data::{
    DiagramConnection, DiagramConnectionType, DiagramDataModel, DiagramPoint, DiagramPointType,
};
pub use definition::{DiagramCategory, DiagramDefinition};
pub use model::{
    DiagramNode, DiagramType, SmartArt, SmartArtBuilder, generate_smartart_colors_xml,
    generate_smartart_data_xml, generate_smartart_drawing_xml, generate_smartart_layout_xml,
    generate_smartart_quickstyle_xml,
};

/// DrawingML diagram namespace (transitional).
pub const DGM_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
/// DrawingML diagram namespace (ISO/IEC 29500 Strict).
pub const DGM_NAMESPACE_STRICT: &str = "http://purl.oclc.org/ooxml/drawingml/diagram";

/// Diagram data relationship type (transitional).
pub const DIAGRAM_DATA_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
/// Diagram layout relationship type (transitional).
pub const DIAGRAM_LAYOUT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout";
/// Diagram quick-style relationship type (transitional).
pub const DIAGRAM_QUICK_STYLE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle";
/// Diagram colors relationship type (transitional).
pub const DIAGRAM_COLORS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors";

/// Diagram data relationship type (ISO/IEC 29500 Strict).
pub const STRICT_DIAGRAM_DATA_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData";
/// Diagram layout relationship type (ISO/IEC 29500 Strict).
pub const STRICT_DIAGRAM_LAYOUT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout";
/// Diagram quick-style relationship type (ISO/IEC 29500 Strict).
pub const STRICT_DIAGRAM_QUICK_STYLE_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle";
/// Diagram colors relationship type (ISO/IEC 29500 Strict).
pub const STRICT_DIAGRAM_COLORS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors";

/// Microsoft extension relationship type for the pre-rendered diagram drawing.
///
/// This URI is not translated in Strict packages; both dialects use it.
pub const MS_DIAGRAM_DRAWING_REL: &str =
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
