//! SmartArt (DrawingML diagram) authoring for DOCX documents.
//!
//! A SmartArt graphic is anchored in the document body as a `w:drawing` whose
//! `a:graphicData` (in the `drawingml/diagram` namespace) carries a
//! `dgm:relIds` element with the relationship IDs of the four diagram parts
//! (data, layout, quick style, colors) generated under `/word/diagrams/`.
//! The parts are produced by the shared generators in
//! [`crate::diagrams::model`] and are discoverable after save and reopen
//! through [`crate::docx::Document::smart_arts`].
//!
//! The optional pre-rendered `drawingN.xml` part is deliberately not
//! generated: Word and LibreOffice re-render the diagram from the layout and
//! data parts when it is absent.

use crate::diagrams::DGM_NAMESPACE;
use crate::diagrams::model::{
    SmartArt, generate_smartart_colors_xml, generate_smartart_data_xml,
    generate_smartart_layout_xml, generate_smartart_quickstyle_xml,
};
use crate::error::{OoxmlError, Result};
use litchi_core::unit::EMUS_PER_INCH;
use litchi_core::xml::escape_xml;
use std::fmt::Write as FmtWrite;

/// Maximum SmartArt diagrams in one document, matching the read-side bound.
pub const MAX_SMART_ARTS: usize = 64;
/// Maximum total nodes in an authored diagram.
const MAX_DIAGRAM_NODES: usize = 4096;
/// Maximum node-tree depth in an authored diagram.
const MAX_DIAGRAM_DEPTH: u32 = 64;
/// Default anchor width (4 inches) when no explicit size is set.
const DEFAULT_WIDTH_EMU: i64 = 4 * EMUS_PER_INCH;
/// Default anchor height (2 inches) when no explicit size is set.
const DEFAULT_HEIGHT_EMU: i64 = 2 * EMUS_PER_INCH;

/// The four relationship IDs of a SmartArt anchor (`dgm:relIds`).
#[derive(Clone, Debug)]
pub(crate) struct SmartArtRelIds {
    /// Relationship ID of the data-model part (`r:dm`).
    pub(crate) data: String,
    /// Relationship ID of the layout part (`r:lo`).
    pub(crate) layout: String,
    /// Relationship ID of the quick-style part (`r:qs`).
    pub(crate) quick_style: String,
    /// Relationship ID of the colors part (`r:cs`).
    pub(crate) colors: String,
}

/// The generated content of one SmartArt part graph.
pub(crate) struct SmartArtPartXml {
    /// `dgm:dataModel` part XML.
    pub(crate) data_xml: String,
    /// `dgm:layoutDef` part XML.
    pub(crate) layout_xml: String,
    /// `dgm:styleDef` part XML.
    pub(crate) quick_style_xml: String,
    /// `dgm:colorsDef` part XML.
    pub(crate) colors_xml: String,
}

/// A mutable SmartArt diagram being authored in a document.
///
/// Wraps a semantic [`SmartArt`] built with
/// [`crate::diagrams::model::SmartArtBuilder`], adding the drawing anchor
/// identity and extents. Attach it to a document via
/// [`crate::docx::writer::MutableDocument::add_smart_art`].
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::docx::{DiagramType, MutableSmartArt, Package, SmartArtBuilder};
///
/// let mut package = Package::new()?;
/// let smartart = SmartArtBuilder::new(DiagramType::Process)
///     .add_items(["Plan", "Build", "Ship"])
///     .build();
/// package
///     .document_mut()?
///     .add_smart_art(MutableSmartArt::new(smartart)?)?;
/// package.save("out.docx")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug)]
pub struct MutableSmartArt {
    /// The semantic diagram model (nodes, type, layout name).
    pub(crate) smartart: SmartArt,
    /// Drawing name written to `wp:docPr@name`.
    pub(crate) name: String,
    /// Drawing element ID written to `wp:docPr@id`.
    pub(crate) doc_pr_id: u32,
    /// Anchor width in EMUs (English Metric Units, 1 inch = 914400 EMUs).
    pub(crate) width_emu: i64,
    /// Anchor height in EMUs.
    pub(crate) height_emu: i64,
    /// Per-document key used to bind the four relationship IDs at save time;
    /// assigned by `MutableDocument::add_smart_art`.
    pub(crate) anchor_key: String,
}

impl MutableSmartArt {
    /// Wrap a built [`SmartArt`] for authoring.
    ///
    /// Validates the diagram: it must contain at least one node, and the node
    /// tree is bounded to `MAX_DIAGRAM_NODES` nodes and `MAX_DIAGRAM_DEPTH`
    /// levels.
    pub fn new(smartart: SmartArt) -> Result<Self> {
        validate_smartart(&smartart)?;
        Ok(Self {
            smartart,
            name: "SmartArt".to_string(),
            // Matches the writer convention for inline pictures (`wp:docPr
            // id="1"`); callers composing several drawings should assign
            // distinct IDs with `set_id`.
            doc_pr_id: 1,
            width_emu: DEFAULT_WIDTH_EMU,
            height_emu: DEFAULT_HEIGHT_EMU,
            anchor_key: String::new(),
        })
    }

    /// Get the semantic diagram model.
    pub fn smartart(&self) -> &SmartArt {
        &self.smartart
    }

    /// Set the drawing name (`wp:docPr@name`).
    pub fn set_name(&mut self, name: impl Into<String>) -> &mut Self {
        self.name = name.into();
        self
    }

    /// Set the drawing element ID (`wp:docPr@id`).
    pub fn set_id(&mut self, id: u32) -> &mut Self {
        self.doc_pr_id = id;
        self
    }

    /// Set the anchor extents in EMUs (both must be positive).
    pub fn set_size_emu(&mut self, width_emu: i64, height_emu: i64) -> Result<&mut Self> {
        if width_emu <= 0 || height_emu <= 0 {
            return Err(OoxmlError::InvalidFormat(
                "SmartArt anchor extents must be positive EMU values".to_string(),
            ));
        }
        self.width_emu = width_emu;
        self.height_emu = height_emu;
        Ok(self)
    }

    /// Get the assigned anchor key (empty until the diagram is added to a
    /// document).
    pub fn anchor_key(&self) -> &str {
        &self.anchor_key
    }

    /// Generate the four diagram part bodies (data, layout, quick style,
    /// colors) for this diagram.
    pub(crate) fn generate_parts(&self) -> SmartArtPartXml {
        SmartArtPartXml {
            data_xml: generate_smartart_data_xml(&self.smartart),
            layout_xml: generate_smartart_layout_xml(&self.smartart),
            quick_style_xml: generate_smartart_quickstyle_xml(),
            colors_xml: generate_smartart_colors_xml(),
        }
    }

    /// Serialize the `<w:drawing>` anchor with the `dgm:relIds` reference.
    ///
    /// When the relationship IDs are unavailable (placeholder serialization),
    /// deterministic `{{SMARTART_*}}` placeholders are emitted, matching the
    /// writer's hyperlink/image placeholder convention. The enclosing `<w:r>`
    /// wrapper is emitted by the paragraph writer.
    pub(crate) fn to_xml(&self, xml: &mut String, rel_ids: Option<&SmartArtRelIds>) -> Result<()> {
        if self.anchor_key.is_empty() {
            return Err(OoxmlError::InvalidFormat(
                "SmartArt has no anchor key; add it via MutableDocument::add_smart_art".to_string(),
            ));
        }
        let placeholder = |part: &str| format!("{{{{SMARTART_{part}_{}}}}}", self.anchor_key);
        let placeholders;
        let (data, layout, quick_style, colors) = match rel_ids {
            Some(ids) => (
                ids.data.as_str(),
                ids.layout.as_str(),
                ids.quick_style.as_str(),
                ids.colors.as_str(),
            ),
            None => {
                placeholders = (
                    placeholder("DM"),
                    placeholder("LO"),
                    placeholder("QS"),
                    placeholder("CS"),
                );
                (
                    placeholders.0.as_str(),
                    placeholders.1.as_str(),
                    placeholders.2.as_str(),
                    placeholders.3.as_str(),
                )
            },
        };
        write!(
            xml,
            r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{}" cy="{}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{}" name="{}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="{DGM_NAMESPACE}"><dgm:relIds xmlns:dgm="{DGM_NAMESPACE}" r:dm="{}" r:lo="{}" r:qs="{}" r:cs="{}"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#,
            self.width_emu,
            self.height_emu,
            self.doc_pr_id,
            escape_xml(&self.name),
            data,
            layout,
            quick_style,
            colors,
        )
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        Ok(())
    }
}

/// Validate the authored diagram against the bounded limits.
fn validate_smartart(smartart: &SmartArt) -> Result<()> {
    if smartart.nodes.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "SmartArt diagram must contain at least one node".to_string(),
        ));
    }
    let mut nodes = 0usize;
    let mut stack: Vec<(&crate::diagrams::model::DiagramNode, u32)> =
        smartart.nodes.iter().map(|node| (node, 1)).collect();
    while let Some((node, depth)) = stack.pop() {
        nodes += 1;
        if nodes > MAX_DIAGRAM_NODES || depth > MAX_DIAGRAM_DEPTH {
            return Err(OoxmlError::InvalidFormat(
                "SmartArt diagram node tree limit exceeded".to_string(),
            ));
        }
        stack.extend(node.children.iter().map(|child| (child, depth + 1)));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::diagrams::DiagramType;
    use crate::diagrams::model::{DiagramNode, SmartArtBuilder};

    fn built() -> SmartArt {
        SmartArtBuilder::new(DiagramType::Process)
            .add_items(["Plan", "Build", "Ship"])
            .build()
    }

    #[test]
    fn validates_diagram_bounds() {
        assert!(MutableSmartArt::new(built()).is_ok());
        assert!(MutableSmartArt::new(SmartArt::new(DiagramType::Process)).is_err());

        let mut deep = SmartArt::new(DiagramType::Hierarchy);
        let mut node = DiagramNode::new("root");
        let mut cursor = &mut node;
        for level in 0..MAX_DIAGRAM_DEPTH {
            cursor.add_child(DiagramNode::new(format!("level {level}")));
            cursor = cursor.children.last_mut().unwrap();
        }
        deep.add_node(node);
        assert!(MutableSmartArt::new(deep).is_err());
    }

    #[test]
    fn serializes_anchor_with_relationship_ids() {
        let mut smartart = MutableSmartArt::new(built()).unwrap();
        smartart.set_id(7).set_name("Org Plan");
        smartart.anchor_key = "smartart1".to_string();
        let rel_ids = SmartArtRelIds {
            data: "rIdDm".to_string(),
            layout: "rIdLo".to_string(),
            quick_style: "rIdQs".to_string(),
            colors: "rIdCs".to_string(),
        };
        let mut xml = String::new();
        smartart.to_xml(&mut xml, Some(&rel_ids)).unwrap();
        assert!(xml.contains("<wp:inline"));
        assert!(xml.contains(r#"<wp:docPr id="7" name="Org Plan"/>"#));
        assert!(xml.contains(&format!(
            r#"<a:graphicData uri="{DGM_NAMESPACE}"><dgm:relIds xmlns:dgm="{DGM_NAMESPACE}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>"#
        )));
    }

    #[test]
    fn serializes_placeholders_and_requires_anchor_key() {
        let smartart = MutableSmartArt::new(built()).unwrap();
        assert!(smartart.to_xml(&mut String::new(), None).is_err());

        let mut smartart = MutableSmartArt::new(built()).unwrap();
        smartart.anchor_key = "smartart2".to_string();
        let mut xml = String::new();
        smartart.to_xml(&mut xml, None).unwrap();
        assert!(xml.contains(r#"r:dm="{{SMARTART_DM_smartart2}}""#));
        assert!(xml.contains(r#"r:cs="{{SMARTART_CS_smartart2}}""#));
    }

    #[test]
    fn rejects_invalid_extents() {
        let mut smartart = MutableSmartArt::new(built()).unwrap();
        assert!(smartart.set_size_emu(0, 914400).is_err());
        assert!(smartart.set_size_emu(914400, -1).is_err());
        assert!(smartart.set_size_emu(914400, 914400).is_ok());
    }

    #[test]
    fn round_trips_smartart_through_saved_package() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            document.add_paragraph_with_text("before the diagram");

            let mut hierarchy = DiagramNode::new("Root");
            hierarchy.add_child(DiagramNode::new("Child A"));
            hierarchy.add_child(DiagramNode::new("Child B"));
            let mut smartart_model = SmartArt::new(DiagramType::Hierarchy);
            smartart_model.add_node(hierarchy);
            let mut smartart = MutableSmartArt::new(smartart_model).unwrap();
            smartart.set_name("Org Chart").set_id(31);
            document.add_smart_art(smartart).unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();
        assert!(document.text().unwrap().contains("before the diagram"));

        let smart_arts = document.smart_arts().unwrap();
        assert_eq!(smart_arts.len(), 1);
        let discovered = &smart_arts[0];
        assert_eq!(discovered.diagram_type(), DiagramType::Hierarchy);
        assert_eq!(discovered.text(), "Root\nChild A\nChild B");
        assert_eq!(discovered.data_part_name, "/word/diagrams/data1.xml");
        assert_eq!(
            discovered.layout_part_name.as_deref(),
            Some("/word/diagrams/layout1.xml")
        );
        assert_eq!(
            discovered.quick_style_part_name.as_deref(),
            Some("/word/diagrams/quickStyle1.xml")
        );
        assert_eq!(
            discovered.colors_part_name.as_deref(),
            Some("/word/diagrams/colors1.xml")
        );
        // The pre-rendered drawing part is deliberately not generated.
        assert_eq!(discovered.drawing_part_name, None);
        // Definition metadata parses back from the generated parts.
        assert!(discovered.layout.is_some());
        assert!(discovered.quick_style.is_some());
        assert!(discovered.colors.is_some());
    }

    #[test]
    fn multiple_smartarts_get_distinct_parts_and_rels() {
        use crate::docx::Package;
        use tempfile::NamedTempFile;

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();
            let first = MutableSmartArt::new(built()).unwrap();
            let first = document.add_smart_art(first).unwrap();
            assert_eq!(first.anchor_key(), "smartart1");
            let second = MutableSmartArt::new(
                SmartArtBuilder::new(DiagramType::Cycle)
                    .add_items(["Alpha", "Beta"])
                    .build(),
            )
            .unwrap();
            let second = document.add_smart_art(second).unwrap();
            assert_eq!(second.anchor_key(), "smartart2");
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let smart_arts = reopened.document().unwrap().smart_arts().unwrap();
        assert_eq!(smart_arts.len(), 2);
        let mut part_names: Vec<&str> = smart_arts
            .iter()
            .map(|smartart| smartart.data_part_name.as_str())
            .collect();
        part_names.sort_unstable();
        assert_eq!(
            part_names,
            ["/word/diagrams/data1.xml", "/word/diagrams/data2.xml"]
        );
        let mut texts: Vec<String> = smart_arts.iter().map(|s| s.text()).collect();
        texts.sort();
        assert_eq!(texts, ["Alpha\nBeta", "Plan\nBuild\nShip"]);
        assert_ne!(
            smart_arts[0].data_relationship_id, smart_arts[1].data_relationship_id,
            "each diagram anchor references its own data relationship"
        );
    }

    #[test]
    fn coexists_with_text_boxes_images_and_ole_objects() {
        use crate::docx::textbox::TextBoxAnchor;
        use crate::docx::{MutableOleObject, MutableTextBox, Package};
        use tempfile::NamedTempFile;

        const PNG_HEADER: &[u8] = &[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        let file = NamedTempFile::with_suffix(".docx").unwrap();
        let mut package = Package::new().unwrap();
        {
            let document = package.document_mut().unwrap();

            let mut text_box = MutableTextBox::new("Companion Box", 914400, 457200).unwrap();
            text_box.add_run("box story");
            document.add_text_box(text_box);

            document
                .add_paragraph()
                .add_picture_from_bytes(PNG_HEADER.to_vec(), Some(914400), Some(914400))
                .unwrap();

            document
                .add_ole_object(MutableOleObject::new("Package", b"payload".to_vec()).unwrap())
                .unwrap();

            document
                .add_smart_art(MutableSmartArt::new(built()).unwrap())
                .unwrap();
        }
        package.save(file.path()).unwrap();

        let reopened = Package::open(file.path()).unwrap();
        let document = reopened.document().unwrap();

        let smart_arts = document.smart_arts().unwrap();
        assert_eq!(smart_arts.len(), 1);
        assert_eq!(smart_arts[0].text(), "Plan\nBuild\nShip");

        let text_boxes = document.text_boxes().unwrap();
        assert_eq!(text_boxes.len(), 1);
        assert_eq!(text_boxes[0].anchor, TextBoxAnchor::Inline);
        assert_eq!(text_boxes[0].text(), "box story");

        let embedded = reopened.embedded_parts().unwrap();
        assert_eq!(embedded.len(), 1);

        let picture_count = document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| {
                paragraph
                    .drawing_objects()
                    .unwrap()
                    .iter()
                    .filter(|drawing| drawing.name() == "Picture")
                    .count()
            })
            .sum::<usize>();
        assert_eq!(picture_count, 1);
    }
}
