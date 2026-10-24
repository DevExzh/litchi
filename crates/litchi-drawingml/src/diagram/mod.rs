//! `DrawingML` diagram (`SmartArt`) support shared across OOXML formats.
//!
//! A `SmartArt` graphic is stored as a set of diagram parts — the data model
//! (`dataN.xml`), layout (`layoutN.xml`), quick style (`quickStyleN.xml`),
//! colors (`colorsN.xml`), and an optional pre-rendered drawing
//! (`drawingN.xml`) — anchored from the host document through a
//! `dgm:relIds` (Word) or graphic-frame (`PowerPoint`) reference.
//!
//! This module holds the format-agnostic pieces:
//!
//! - [`model`]: the semantic `SmartArt`/`DiagramNode` model, its builder, and
//!   the generators for the five diagram parts.
//! - [`data`]: a bounded, inert parser for the `dgm:dataModel` node and
//!   connection graph.
//! - [`definition`]: inert header metadata for the layout, quick-style, and
//!   colors definition parts.
//! - [`authoring`]: deterministic authoring of a graphic from a text outline
//!   with a built-in layout, including Litchi's own cached drawing.
//! - [`edit`]: the outline of a data model and source-preserving text and
//!   drawing-reference edits.
//!
//! Both the transitional and the ISO Strict `drawingml/diagram` namespaces are
//! supported. Format-specific anchoring remains in the concrete `litchi-docx`
//! and `litchi-pptx` crates.
//!
//! The shared part vocabulary follows `[MS-ODRAWXML]` diagram structures;
//! `[MS-OI29500]` package relationships and the host anchors described by
//! `[MS-PPTX]` remain outside this module.

pub mod authoring;
pub mod data;
pub mod definition;
pub mod edit;
pub mod model;

/// Caller-selected bounds for diagram raw readers, editors and generators.
///
/// The defaults retain the historical hard ceilings of the diagram codecs.
/// A caller may select narrower bounds, but cannot widen those ceilings.  All
/// bounded public diagram operations take this value explicitly; the value is
/// validated at the operation boundary so invalid policies report the typed
/// `InvalidLimit` error before any input or output allocation is attempted.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[must_use]
pub struct Limits {
    input_bytes: usize,
    output_bytes: usize,
    xml_nodes: usize,
    xml_depth: usize,
    points: usize,
    connections: usize,
    data_text_bytes: usize,
    nodes: usize,
    node_depth: usize,
    node_text_bytes: usize,
}

impl Limits {
    /// Historical bounded diagram policy.
    pub const DEFAULT: Self = Self {
        input_bytes: 16 * 1024 * 1024,
        output_bytes: 16 * 1024 * 1024,
        xml_nodes: 200_000,
        xml_depth: 128,
        points: 100_000,
        connections: 100_000,
        data_text_bytes: 1024 * 1024,
        nodes: 200_000,
        node_depth: 128,
        node_text_bytes: 1024 * 1024,
    };

    /// Construct a caller-selected diagram policy.
    ///
    /// Values are checked when the policy is passed to a public operation;
    /// zero and over-hard-cap values then return `Error::InvalidLimit` with
    /// the corresponding `LimitKind`.
    #[must_use]
    #[allow(clippy::too_many_arguments)]
    pub const fn new(
        input_bytes: usize,
        output_bytes: usize,
        xml_nodes: usize,
        xml_depth: usize,
        points: usize,
        connections: usize,
        data_text_bytes: usize,
        nodes: usize,
        node_depth: usize,
        node_text_bytes: usize,
    ) -> Self {
        Self {
            input_bytes,
            output_bytes,
            xml_nodes,
            xml_depth,
            points,
            connections,
            data_text_bytes,
            nodes,
            node_depth,
            node_text_bytes,
        }
    }

    /// Maximum input bytes admitted by the policy.
    #[must_use]
    pub const fn input_bytes_value(self) -> usize {
        self.input_bytes
    }

    /// Maximum output bytes admitted by the policy.
    #[must_use]
    pub const fn output_bytes_value(self) -> usize {
        self.output_bytes
    }

    /// Maximum XML nodes admitted by the policy.
    #[must_use]
    pub const fn xml_nodes_value(self) -> usize {
        self.xml_nodes
    }

    /// Maximum XML nesting depth admitted by the policy.
    #[must_use]
    pub const fn xml_depth_value(self) -> usize {
        self.xml_depth
    }

    /// Maximum data-model points admitted by the policy.
    #[must_use]
    pub const fn points_value(self) -> usize {
        self.points
    }

    /// Maximum data-model connections admitted by the policy.
    #[must_use]
    pub const fn connections_value(self) -> usize {
        self.connections
    }

    /// Maximum bytes of one data-model text value admitted by the policy.
    #[must_use]
    pub const fn data_text_bytes_value(self) -> usize {
        self.data_text_bytes
    }

    /// Maximum semantic/content nodes admitted by the policy.
    #[must_use]
    pub const fn nodes_value(self) -> usize {
        self.nodes
    }

    /// Maximum semantic/content node depth admitted by the policy.
    #[must_use]
    pub const fn node_depth_value(self) -> usize {
        self.node_depth
    }

    /// Maximum bytes of one semantic/content node text value.
    #[must_use]
    pub const fn node_text_bytes_value(self) -> usize {
        self.node_text_bytes
    }

    pub fn validate(self) -> crate::Result<()> {
        for (kind, value, maximum) in [
            (
                crate::LimitKind::InputBytes,
                self.input_bytes,
                16 * 1024 * 1024,
            ),
            (
                crate::LimitKind::OutputBytes,
                self.output_bytes,
                16 * 1024 * 1024,
            ),
            (crate::LimitKind::XmlNodes, self.xml_nodes, 200_000),
            (crate::LimitKind::XmlDepth, self.xml_depth, 128),
            (crate::LimitKind::DiagramPoints, self.points, 100_000),
            (
                crate::LimitKind::DiagramConnections,
                self.connections,
                100_000,
            ),
            (
                crate::LimitKind::DiagramDataModelTextBytes,
                self.data_text_bytes,
                1024 * 1024,
            ),
            (crate::LimitKind::DiagramNodes, self.nodes, 200_000),
            (crate::LimitKind::DiagramNodeDepth, self.node_depth, 128),
            (
                crate::LimitKind::DiagramNodeTextBytes,
                self.node_text_bytes,
                1024 * 1024,
            ),
        ] {
            if value == 0 || value > maximum {
                return Err(crate::Error::invalid_limit(kind, value, maximum));
            }
        }
        Ok(())
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

pub use data::{
    Conformance, Connection, ConnectionType, DiagramDataModel, Id, IdError, Point, PointType,
};
pub use definition::{DiagramCategory, DiagramDefinition};
pub use model::{
    DiagramNode, DiagramType, SmartArt, SmartArtBuilder, generate_smartart_colors_xml,
    generate_smartart_data_xml, generate_smartart_drawing_xml, generate_smartart_layout_xml,
    generate_smartart_quickstyle_xml,
};

/// `DrawingML` diagram namespace (transitional).
pub const DGM_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
/// `DrawingML` diagram namespace (ISO/IEC 29500 Strict).
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
