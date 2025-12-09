//! SmartArt/Diagram support for PowerPoint presentations.
//!
//! SmartArt graphics are represented as diagrams in OOXML. This module provides
//! read support for extracting diagram information from presentations.

use crate::ooxml::error::{OoxmlError, Result};
use quick_xml::Reader;
use quick_xml::events::Event;

/// SmartArt diagram type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramType {
    /// List diagram
    List,
    /// Process diagram
    Process,
    /// Cycle diagram
    Cycle,
    /// Hierarchy diagram
    Hierarchy,
    /// Relationship diagram
    Relationship,
    /// Matrix diagram
    Matrix,
    /// Pyramid diagram
    Pyramid,
    /// Picture diagram
    Picture,
    /// Unknown diagram type
    Unknown,
}

impl DiagramType {
    /// Parse diagram type from layout type URI.
    pub fn from_layout_uri(uri: &str) -> Self {
        let uri_lower = uri.to_lowercase();
        if uri_lower.contains("list") {
            DiagramType::List
        } else if uri_lower.contains("process") {
            DiagramType::Process
        } else if uri_lower.contains("cycle") {
            DiagramType::Cycle
        } else if uri_lower.contains("hierarchy") || uri_lower.contains("orgchart") {
            DiagramType::Hierarchy
        } else if uri_lower.contains("relationship") || uri_lower.contains("venn") {
            DiagramType::Relationship
        } else if uri_lower.contains("matrix") {
            DiagramType::Matrix
        } else if uri_lower.contains("pyramid") {
            DiagramType::Pyramid
        } else if uri_lower.contains("picture") {
            DiagramType::Picture
        } else {
            DiagramType::Unknown
        }
    }
}

/// A SmartArt diagram node/item.
#[derive(Debug, Clone)]
pub struct DiagramNode {
    /// Node text content
    pub text: String,
    /// Child nodes
    pub children: Vec<DiagramNode>,
    /// Node depth level (0 = root)
    pub depth: u32,
}

impl DiagramNode {
    /// Create a new diagram node.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            children: Vec::new(),
            depth: 0,
        }
    }

    /// Add a child node.
    pub fn add_child(&mut self, child: DiagramNode) {
        self.children.push(child);
    }

    /// Get all text from this node and its children.
    pub fn all_text(&self) -> String {
        let mut result = self.text.clone();
        for child in &self.children {
            if !result.is_empty() && !child.text.is_empty() {
                result.push('\n');
            }
            result.push_str(&child.all_text());
        }
        result
    }
}

/// SmartArt diagram information.
#[derive(Debug, Clone)]
pub struct SmartArt {
    /// Diagram type
    pub diagram_type: DiagramType,
    /// Root nodes of the diagram
    pub nodes: Vec<DiagramNode>,
    /// Layout name/description if available
    pub layout_name: Option<String>,
    /// Unique ID of the diagram
    pub id: Option<String>,
}

impl SmartArt {
    /// Create a new SmartArt diagram.
    pub fn new(diagram_type: DiagramType) -> Self {
        Self {
            diagram_type,
            nodes: Vec::new(),
            layout_name: None,
            id: None,
        }
    }

    /// Add a root node.
    pub fn add_node(&mut self, node: DiagramNode) {
        self.nodes.push(node);
    }

    /// Get all text content from the diagram.
    pub fn text(&self) -> String {
        self.nodes
            .iter()
            .map(|n| n.all_text())
            .collect::<Vec<_>>()
            .join("\n")
    }

    /// Get the number of root nodes.
    pub fn node_count(&self) -> usize {
        self.nodes.len()
    }

    /// Parse SmartArt data XML (dgm:data).
    pub fn parse_data_xml(xml: &str) -> Result<Vec<DiagramNode>> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(true);

        let mut nodes = Vec::new();
        let mut current_text = String::new();
        let mut in_text = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) => {
                    let local = e.local_name();
                    if local.as_ref() == b"t" {
                        in_text = true;
                        current_text.clear();
                    }
                },
                Ok(Event::End(e)) => {
                    let local = e.local_name();
                    if local.as_ref() == b"t" {
                        in_text = false;
                        if !current_text.trim().is_empty() {
                            nodes.push(DiagramNode::new(current_text.trim()));
                        }
                    }
                },
                Ok(Event::Text(e)) => {
                    if in_text && let Ok(text) = std::str::from_utf8(e.as_ref()) {
                        current_text.push_str(text);
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        Ok(nodes)
    }
}

/// Builder for creating SmartArt diagrams.
pub struct SmartArtBuilder {
    diagram_type: DiagramType,
    nodes: Vec<DiagramNode>,
    layout_name: Option<String>,
}

impl SmartArtBuilder {
    /// Create a new SmartArt builder with the specified diagram type.
    pub fn new(diagram_type: DiagramType) -> Self {
        Self {
            diagram_type,
            nodes: Vec::new(),
            layout_name: None,
        }
    }

    /// Set the layout name.
    pub fn layout_name(mut self, name: impl Into<String>) -> Self {
        self.layout_name = Some(name.into());
        self
    }

    /// Add a text item to the diagram.
    pub fn add_item(mut self, text: impl Into<String>) -> Self {
        self.nodes.push(DiagramNode::new(text));
        self
    }

    /// Add multiple text items.
    pub fn add_items<I, S>(mut self, items: I) -> Self
    where
        I: IntoIterator<Item = S>,
        S: Into<String>,
    {
        for item in items {
            self.nodes.push(DiagramNode::new(item));
        }
        self
    }

    /// Build the SmartArt diagram.
    pub fn build(self) -> SmartArt {
        SmartArt {
            diagram_type: self.diagram_type,
            nodes: self.nodes,
            layout_name: self.layout_name,
            id: None,
        }
    }
}

/// Generate SmartArt data XML.
pub fn generate_smartart_data_xml(smartart: &SmartArt) -> String {
    let mut xml = String::with_capacity(2048);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#,
    );
    xml.push_str("<dgm:ptLst>");

    // Root point: declare layout, quick style, and color identifiers via prSet.
    // These identifiers follow the OOXML SmartArt semantics used by Office and
    // reference implementations. We choose layout based on DiagramType, while
    // reusing the same quick style and color scheme that we generate in
    // quickStyle/colors parts.
    //
    // loTypeId values are aligned with the built-in layout IDs used by
    // PowerPoint (e.g. process1, cycle2, orgChart1, venn1, matrix3, pyramid1)
    // so that each diagram type can use its distinct default layout. For these
    // built-in layouts, PowerPoint typically leaves loCatId empty and derives
    // category from the layout definition's catLst instead, so we follow that
    // pattern and use an empty string for loCatId.
    let (lo_type_id, lo_cat_id) = match smartart.diagram_type {
        DiagramType::List => (
            "urn:microsoft.com/office/officeart/2005/8/layout/default",
            "",
        ),
        DiagramType::Process => (
            "urn:microsoft.com/office/officeart/2005/8/layout/process1",
            "",
        ),
        DiagramType::Cycle => (
            "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
            "",
        ),
        DiagramType::Hierarchy => (
            "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1",
            "",
        ),
        DiagramType::Relationship => ("urn:microsoft.com/office/officeart/2005/8/layout/venn1", ""),
        DiagramType::Matrix => (
            "urn:microsoft.com/office/officeart/2005/8/layout/matrix3",
            "",
        ),
        DiagramType::Pyramid => (
            "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
            "",
        ),
        DiagramType::Picture | DiagramType::Unknown => (
            "urn:microsoft.com/office/officeart/2005/8/layout/default",
            "",
        ),
    };

    // Quick style and colors: use the same built-in identifiers as in our
    // quickStyle/colors parts (simple1, accent1_1).
    let qs_type_id = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
    let qs_cat_id = "simple";
    let cs_type_id = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_1";
    let cs_cat_id = "accent1";

    xml.push_str("<dgm:pt modelId=\"0\" type=\"doc\">");
    xml.push_str(&format!(
        "<dgm:prSet loTypeId=\"{}\" loCatId=\"{}\" qsTypeId=\"{}\" qsCatId=\"{}\" csTypeId=\"{}\" csCatId=\"{}\"/>",
        lo_type_id, lo_cat_id, qs_type_id, qs_cat_id, cs_type_id, cs_cat_id
    ));
    xml.push_str(
        r#"<dgm:spPr/><dgm:t><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:endParaRPr/></a:p></dgm:t></dgm:pt>"#,
    );

    // Diagram types for which we emit a full presentation model (pres points,
    // transitions, and a SmartArt-style layout). Picture/Unknown remain
    // minimal for now.
    let has_presentation_model = matches!(
        smartart.diagram_type,
        DiagramType::List
            | DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    );

    // Flatten diagram nodes (including hierarchy children) into a linear
    // sequence with parent references so that we can build the SmartArt data
    // graph.
    let mut flat_nodes: Vec<(i32, i32, &DiagramNode)> = Vec::new();
    let mut next_model_id: i32 = 1;

    fn flatten_nodes<'a>(
        out: &mut Vec<(i32, i32, &'a DiagramNode)>,
        next_id: &mut i32,
        parent_id: i32,
        node: &'a DiagramNode,
    ) {
        let id = *next_id;
        *next_id += 1;
        out.push((id, parent_id, node));
        for child in &node.children {
            flatten_nodes(out, next_id, id, child);
        }
    }

    for root in &smartart.nodes {
        flatten_nodes(&mut flat_nodes, &mut next_model_id, 0, root);
    }

    let node_count = flat_nodes.len();

    // Add content nodes (type defaults to "node").
    for (model_id, _, node) in &flat_nodes {
        xml.push_str(&format!(
            r#"<dgm:pt modelId="{}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p></dgm:t></dgm:pt>"#,
            model_id,
            escape_xml(&node.text)
        ));
    }

    // For supported diagrams, add minimal presentation points (type="pres")
    // that map logical nodes to presentation nodes, plus a diagram-level
    // presentation point. This models the basic pres/presParOf graph used by
    // SmartArt layouts without copying concrete data from sample files.
    if has_presentation_model && node_count > 0 {
        let base_pres_id = 1000_i32;
        let diagram_pres_id = base_pres_id;
        let first_node_pres_id = base_pres_id + 1;

        // Diagram-level presentation point associated with the document root.
        xml.push_str(&format!(
            r#"<dgm:pt modelId="{}" type="pres"><dgm:prSet presAssocID="0" presName="diagram" presStyleCnt="0"><dgm:presLayoutVars><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#,
            diagram_pres_id,
        ));

        // One presentation point per logical node, each associated with that
        // node and using a simple bullet-enabled layout.
        for (idx, _) in smartart.nodes.iter().enumerate() {
            let node_id = (idx + 1) as i32;
            let pres_id = first_node_pres_id + idx as i32;
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="pres"><dgm:prSet presAssocID="{}" presName="node" presStyleLbl="node0" presStyleIdx="0" presStyleCnt="1"><dgm:presLayoutVars><dgm:bulletEnabled val="1"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#,
                pres_id,
                node_id,
            ));
        }

        // For each structural connection (doc -> node), define transition
        // points of type parTrans and sibTrans. These are referenced from the
        // CT_Cxn entries via parTransId/sibTransId and are used by layouts
        // when iterating over siblings (axis="followSib").
        let base_partrans_id = 2000_i32;
        let base_sibtrans_id = 2100_i32;
        for (idx, _) in flat_nodes.iter().enumerate() {
            let cxn_model_id = 100 + idx as i32;
            let par_pt_id = base_partrans_id + idx as i32;
            let sib_pt_id = base_sibtrans_id + idx as i32;

            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="parTrans" cxnId="{}"><dgm:prSet/><dgm:spPr/></dgm:pt>"#,
                par_pt_id,
                cxn_model_id,
            ));
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="sibTrans" cxnId="{}"><dgm:prSet/><dgm:spPr/></dgm:pt>"#,
                sib_pt_id,
                cxn_model_id,
            ));
        }
    }

    xml.push_str("</dgm:ptLst>");

    // Connection list: structural parent-of connections from the document root
    // to each node, plus (for supported diagrams) basic presentation mappings.
    xml.push_str("<dgm:cxnLst>");

    // parOf connections: doc (0) -> each node, or parent -> child for
    // hierarchical diagrams. We use the flattened node list and connect each
    // node to its recorded parent (0 = document root).
    for (idx, (node_id, parent_id, _)) in flat_nodes.iter().enumerate() {
        let cxn_model_id = 100 + idx as i32;

        // Compute srcOrd as the index of this child among siblings with the
        // same parent.
        let src_ord = flat_nodes[..idx]
            .iter()
            .filter(|(_, p, _)| p == parent_id)
            .count();

        // For diagrams with a full presentation model, attach
        // parTransId/sibTransId pointing to the corresponding transition
        // points. For minimal diagram types, we omit these attributes.
        if has_presentation_model && node_count > 0 {
            let base_partrans_id = 2000_i32;
            let base_sibtrans_id = 2100_i32;
            let par_pt_id = base_partrans_id + idx as i32;
            let sib_pt_id = base_sibtrans_id + idx as i32;

            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" parTransId="{}" sibTransId="{}"/>"#,
                cxn_model_id,
                parent_id,
                node_id,
                src_ord,
                par_pt_id,
                sib_pt_id,
            ));
        } else {
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" srcId="{}" destId="{}" srcOrd="{}" destOrd="0"/>"#,
                cxn_model_id, parent_id, node_id, src_ord,
            ));
        }
    }

    if has_presentation_model && node_count > 0 {
        let base_pres_id = 1000_i32;
        let diagram_pres_id = base_pres_id;
        let first_node_pres_id = base_pres_id + 1;
        let pres_layout_id = lo_type_id;

        // presOf connection: document root -> diagram presentation node. This
        // ties the top-level layout node to the diagram-level presentation
        // point.
        xml.push_str(&format!(
            r#"<dgm:cxn modelId="{}" type="presOf" srcId="0" destId="{}" srcOrd="0" destOrd="0" presId="{}"/>"#,
            250,
            diagram_pres_id,
            pres_layout_id,
        ));

        // presOf connections: each logical node -> its presentation node.
        for (idx, (node_id, _, _)) in flat_nodes.iter().enumerate() {
            let pres_id = first_node_pres_id + idx as i32;
            let cxn_id = 300 + idx as i32;
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="presOf" srcId="{}" destId="{}" srcOrd="0" destOrd="0" presId="{}"/>"#,
                cxn_id,
                node_id,
                pres_id,
                pres_layout_id,
            ));
        }

        // presParOf connections: diagram presentation node -> each
        // presentation node for simple linear ordering.
        for (idx, _) in flat_nodes.iter().enumerate() {
            let pres_id = first_node_pres_id + idx as i32;
            let cxn_id = 500 + idx as i32;
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="presParOf" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" presId="{}"/>"#,
                cxn_id,
                diagram_pres_id,
                pres_id,
                idx,
                pres_layout_id,
            ));
        }
    }

    xml.push_str("</dgm:cxnLst>");

    xml.push_str("<dgm:bg/><dgm:whole/>");
    xml.push_str("</dgm:dataModel>");

    xml
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Generate SmartArt layout XML.
///
/// This generates a simple layout definition for the diagram.
pub fn generate_smartart_layout_xml(smartart: &SmartArt) -> String {
    let mut xml = String::with_capacity(4096);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    // Use a layout type identifier for SmartArt diagrams that matches the
    // loTypeId we emit in the root dgm:prSet so that the data model correctly
    // binds to this embedded layout definition.
    let layout_type_id = match smartart.diagram_type {
        DiagramType::List => "urn:microsoft.com/office/officeart/2005/8/layout/default",
        DiagramType::Process => "urn:microsoft.com/office/officeart/2005/8/layout/process1",
        DiagramType::Cycle => "urn:microsoft.com/office/officeart/2005/8/layout/cycle2",
        DiagramType::Hierarchy => "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1",
        DiagramType::Relationship => "urn:microsoft.com/office/officeart/2005/8/layout/venn1",
        DiagramType::Matrix => "urn:microsoft.com/office/officeart/2005/8/layout/matrix3",
        DiagramType::Pyramid => "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
        DiagramType::Picture | DiagramType::Unknown => {
            "urn:microsoft.com/office/officeart/2005/8/layout/default"
        },
    };
    xml.push_str(&format!(
        "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" uniqueId=\"{}\">",
        layout_type_id,
    ));

    // Category type for the layout, aligned with loCatId in the data model.
    let cat_type = match smartart.diagram_type {
        DiagramType::List => "list",
        DiagramType::Process => "process",
        DiagramType::Cycle => "cycle",
        DiagramType::Hierarchy => "hierarchy",
        DiagramType::Relationship => "relationship",
        DiagramType::Matrix => "matrix",
        DiagramType::Pyramid => "pyramid",
        DiagramType::Picture | DiagramType::Unknown => "list",
    };

    // Per-diagram node geometry used in the layout tree. This controls the
    // visible shape type that the SmartArt engine creates for each logical
    // node (e.g. chevrons for process, circles for cycle, trapezoids for
    // pyramid, etc.).
    let node_shape_type = match smartart.diagram_type {
        DiagramType::Process => "chevron",
        DiagramType::Cycle => "ellipse",
        DiagramType::Relationship => "ellipse",
        DiagramType::Matrix => "roundRect",
        DiagramType::Pyramid => "trapezoid",
        _ => "rect",
    };

    // Diagram types for which we provide a full SmartArt-style layout tree
    // (diagram/node/sibTrans with forEach). Picture/Unknown use a minimal
    // fallback layout.
    let has_presentation_model = matches!(
        smartart.diagram_type,
        DiagramType::List
            | DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    );

    // Simple title for layout
    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");

    // Category list
    xml.push_str("<dgm:catLst>");
    xml.push_str(&format!("<dgm:cat type=\"{}\" pri=\"1000\"/>", cat_type));
    // Many built-in layouts also advertise themselves as convertible to/from
    // other layouts using a secondary "convert" category. Mirroring this
    // helps Office treat our embedded layouts more like full SmartArt
    // definitions instead of falling back to the default list layout.
    if matches!(
        smartart.diagram_type,
        DiagramType::Process
            | DiagramType::Cycle
            | DiagramType::Hierarchy
            | DiagramType::Relationship
            | DiagramType::Matrix
            | DiagramType::Pyramid
    ) {
        xml.push_str("<dgm:cat type=\"convert\" pri=\"15000\"/>");
    }
    xml.push_str("</dgm:catLst>");

    // Sample data. We provide a small, schema-valid sample data model that
    // exercises a document root plus a few placeholder nodes. This is not
    // used directly for rendering but helps Office recognize the layout as a
    // proper SmartArt definition.
    xml.push_str("<dgm:sampData>");
    xml.push_str("<dgm:dataModel>");
    xml.push_str("<dgm:ptLst>");
    xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
    for i in 1..=3 {
        xml.push_str(&format!(r#"<dgm:pt modelId="{}"/>"#, i));
    }
    xml.push_str("</dgm:ptLst>");
    xml.push_str("<dgm:cxnLst/>");
    xml.push_str("<dgm:bg/><dgm:whole/>");
    xml.push_str("</dgm:dataModel>");
    xml.push_str("</dgm:sampData>");

    // Minimal style and color data models. Built-in PowerPoint layouts
    // include styleData and clrData sections alongside sampData. Emitting
    // lightweight versions here makes our embedded layouts look more like
    // those full definitions while keeping the XML compact.
    if has_presentation_model {
        // Style data: root doc node plus two child placeholders, with a
        // simple parent-of connection graph.
        xml.push_str("<dgm:styleData>");
        xml.push_str("<dgm:dataModel>");
        xml.push_str("<dgm:ptLst>");
        xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
        xml.push_str(r#"<dgm:pt modelId="1"/><dgm:pt modelId="2"/>"#);
        xml.push_str("</dgm:ptLst>");
        xml.push_str("<dgm:cxnLst>");
        xml.push_str(r#"<dgm:cxn modelId="3" srcId="0" destId="1" srcOrd="0" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="4" srcId="0" destId="2" srcOrd="1" destOrd="0"/>"#);
        xml.push_str("</dgm:cxnLst>");
        xml.push_str("<dgm:bg/><dgm:whole/>");
        xml.push_str("</dgm:dataModel>");
        xml.push_str("</dgm:styleData>");

        // Color data: similar miniature graph with a few more placeholders so
        // that color engines have something to bind to if needed.
        xml.push_str("<dgm:clrData>");
        xml.push_str("<dgm:dataModel>");
        xml.push_str("<dgm:ptLst>");
        xml.push_str(r#"<dgm:pt modelId="0" type="doc"/>"#);
        xml.push_str(r#"<dgm:pt modelId="1"/><dgm:pt modelId="2"/><dgm:pt modelId="3"/>"#);
        xml.push_str("</dgm:ptLst>");
        xml.push_str("<dgm:cxnLst>");
        xml.push_str(r#"<dgm:cxn modelId="5" srcId="0" destId="1" srcOrd="0" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="6" srcId="0" destId="2" srcOrd="1" destOrd="0"/>"#);
        xml.push_str(r#"<dgm:cxn modelId="7" srcId="0" destId="3" srcOrd="2" destOrd="0"/>"#);
        xml.push_str("</dgm:cxnLst>");
        xml.push_str("<dgm:bg/><dgm:whole/>");
        xml.push_str("</dgm:dataModel>");
        xml.push_str("</dgm:clrData>");
    }

    // Layout node definition
    if has_presentation_model {
        xml.push_str("<dgm:layoutNode name=\"diagram\">");
        xml.push_str("<dgm:varLst>");
        xml.push_str("<dgm:dir/>");
        xml.push_str(r#"<dgm:resizeHandles val="exact"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str("<dgm:alg type=\"lin\"/>");
        xml.push_str("<dgm:shape/>");
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst>");
        xml.push_str(r#"<dgm:constr type=\"primFontSz\" val=\"65\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"w\" for=\"ch\" forName=\"node\" refType=\"w\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"h\" for=\"ch\" forName=\"node\" refType=\"w\" refFor=\"ch\" refForName=\"node\" fact=\"0.5\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"w\" for=\"ch\" forName=\"sibTrans\" refType=\"w\" refFor=\"ch\" refForName=\"node\" fact=\"0.15\"/>"#);
        xml.push_str(
            r#"<dgm:constr type=\"sp\" refType=\"w\" refFor=\"ch\" refForName=\"sibTrans\"/>"#,
        );
        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst/>");

        xml.push_str(r#"<dgm:forEach name=\"items\" axis=\"ch\" ptType=\"node\">"#);
        xml.push_str(r#"<dgm:layoutNode name=\"node\">"#);
        xml.push_str("<dgm:varLst>");
        xml.push_str(r#"<dgm:bulletEnabled val=\"1\"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str(r#"<dgm:alg type=\"tx\"/>"#);
        xml.push_str(&format!("<dgm:shape type=\"{}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:blip=\"\"><dgm:adjLst/></dgm:shape>", node_shape_type));
        xml.push_str(r#"<dgm:presOf axis=\"desOrSelf\" ptType=\"node\"/>"#);
        xml.push_str("<dgm:constrLst>");
        xml.push_str(r#"<dgm:constr type=\"lMarg\" refType=\"primFontSz\" fact=\"0.4\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"rMarg\" refType=\"primFontSz\" fact=\"0.4\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"tMarg\" refType=\"primFontSz\" fact=\"0.4\"/>"#);
        xml.push_str(r#"<dgm:constr type=\"bMarg\" refType=\"primFontSz\" fact=\"0.4\"/>"#);
        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst>");
        xml.push_str(r#"<dgm:rule type=\"primFontSz\" val=\"5\"/>"#);
        xml.push_str("</dgm:ruleLst>");
        xml.push_str("</dgm:layoutNode>");

        xml.push_str(
            r#"<dgm:forEach name=\"spacing\" axis=\"followSib\" ptType=\"sibTrans\" cnt=\"1\">"#,
        );
        xml.push_str(r#"<dgm:layoutNode name=\"sibTrans\">"#);
        xml.push_str(r#"<dgm:alg type=\"sp\"/>"#);
        xml.push_str(r#"<dgm:shape xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:blip=\"\"><dgm:adjLst/></dgm:shape>"#);
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst/>");
        xml.push_str("<dgm:ruleLst/>");
        xml.push_str("</dgm:layoutNode>");
        xml.push_str("</dgm:forEach>");

        xml.push_str("</dgm:forEach>");
        xml.push_str("</dgm:layoutNode>");
    } else {
        xml.push_str("<dgm:layoutNode name=\"root\">");
        xml.push_str("<dgm:varLst>");
        xml.push_str(r#"<dgm:dir val=\"norm\"/>"#);
        xml.push_str("</dgm:varLst>");
        xml.push_str("<dgm:alg type=\"lin\"/>");
        xml.push_str("<dgm:shape/>");
        xml.push_str("<dgm:presOf/>");
        xml.push_str("<dgm:constrLst>");

        // Add constraints based on diagram type
        let spacing = match smartart.diagram_type {
            DiagramType::Process => 150,
            DiagramType::Hierarchy => 200,
            _ => 100,
        };
        xml.push_str(r#"<dgm:constr type="primFontSz" val="65"/>"#);
        xml.push_str(&format!(r#"<dgm:constr type="sp" val="{}"/>"#, spacing));

        xml.push_str("</dgm:constrLst>");
        xml.push_str("<dgm:ruleLst/>");
        xml.push_str("</dgm:layoutNode>");
    }

    xml.push_str("</dgm:layoutDef>");

    xml
}

/// Generate SmartArt drawing XML.
///
/// This converts the SmartArt to a DrawingML representation.
pub fn generate_smartart_drawing_xml(
    smartart: &SmartArt,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
) -> String {
    let mut xml = String::with_capacity(4096);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    // Use the Microsoft extension namespace for diagram drawings (dsp), as used by
    // PowerPoint, Apache POI, and LibreOffice for SmartArt pre-rendered shapes.
    xml.push_str(
        r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" "#,
    );
    xml.push_str(r#"xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#);
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);

    xml.push_str("<dsp:spTree>");

    // Group shape container (match typical PowerPoint output: empty grpSpPr)
    xml.push_str("<dsp:nvGrpSpPr>");
    xml.push_str("<dsp:cNvPr id=\"0\" name=\"\"/>");
    xml.push_str("<dsp:cNvGrpSpPr/>");
    xml.push_str("</dsp:nvGrpSpPr>");
    xml.push_str("<dsp:grpSpPr/>");

    // Generate shapes for each node
    let node_count = smartart.nodes.len().max(1);
    let node_width = width / node_count as i64;
    let node_height = height * 8 / 10; // 80% height

    for (idx, node) in smartart.nodes.iter().enumerate() {
        let node_x = x + (idx as i64 * node_width);
        let node_y = y + (height - node_height) / 2;

        // Use a numeric modelId per shape. ST_ModelId is a union of xsd:int and
        // GUID; using integers keeps us schema-valid. For diagram types with a
        // full presentation model, we align the shape modelIds with the
        // presentation points (type="pres") defined in the data model (1001,
        // 1002, ...). For minimal types, we fall back to the node ids
        // (1-based).
        let has_presentation_model = matches!(
            smartart.diagram_type,
            DiagramType::List
                | DiagramType::Process
                | DiagramType::Cycle
                | DiagramType::Hierarchy
                | DiagramType::Relationship
                | DiagramType::Matrix
                | DiagramType::Pyramid
        );
        let model_id: i32 = if has_presentation_model {
            let base_pres_id = 1000_i32;
            let first_node_pres_id = base_pres_id + 1;
            first_node_pres_id + idx as i32
        } else {
            (idx + 1) as i32
        };

        xml.push_str(&format!("<dsp:sp modelId=\"{}\">", model_id));
        xml.push_str("<dsp:nvSpPr>");
        xml.push_str(&format!(
            "<dsp:cNvPr id=\"{}\" name=\"Shape {}\"/>",
            idx + 1,
            idx + 1
        ));
        xml.push_str("<dsp:cNvSpPr/>");
        xml.push_str("</dsp:nvSpPr>");

        xml.push_str("<dsp:spPr>");
        xml.push_str("<a:xfrm>");
        xml.push_str(&format!(r#"<a:off x="{}" y="{}"/>"#, node_x, node_y));
        xml.push_str(&format!(
            r#"<a:ext cx="{}" cy="{}"/>"#,
            node_width * 9 / 10,
            node_height
        ));
        xml.push_str("</a:xfrm>");

        // Shape type based on diagram type
        let shape_type = match smartart.diagram_type {
            DiagramType::Process => "chevron",
            DiagramType::Cycle => "ellipse",
            DiagramType::Hierarchy => "rect",
            _ => "roundRect",
        };
        xml.push_str(&format!(
            r#"<a:prstGeom prst="{}"><a:avLst/></a:prstGeom>"#,
            shape_type
        ));

        // Default fill and outline
        xml.push_str("<a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill>");
        xml.push_str("<a:ln><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln>");
        xml.push_str("</dsp:spPr>");

        // Basic style block, mirroring the structure used in typical
        // PowerPoint-generated SmartArt drawings.
        xml.push_str("<dsp:style>");
        xml.push_str("<a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef>");
        xml.push_str("<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>");
        xml.push_str("<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>");
        xml.push_str("<a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef>");
        xml.push_str("</dsp:style>");

        // Text body
        xml.push_str("<dsp:txBody>");
        xml.push_str(r#"<a:bodyPr anchor="ctr"/>"#);
        xml.push_str("<a:lstStyle/>");
        xml.push_str("<a:p>");
        xml.push_str("<a:pPr algn=\"ctr\"/>");
        xml.push_str("<a:r>");
        xml.push_str(r#"<a:rPr lang="en-US"/>"#);
        xml.push_str(&format!("<a:t>{}</a:t>", escape_xml(&node.text)));
        xml.push_str("</a:r>");
        xml.push_str("</a:p>");
        xml.push_str("</dsp:txBody>");

        // Text transform (position/size of text box inside the shape). We
        // mirror the shape's geometry so PowerPoint and POI have clear anchors.
        xml.push_str("<dsp:txXfrm>");
        xml.push_str(&format!(r#"<a:off x="{}" y="{}"/>"#, node_x, node_y));
        xml.push_str(&format!(
            r#"<a:ext cx="{}" cy="{}"/>"#,
            node_width * 9 / 10,
            node_height
        ));
        xml.push_str("</dsp:txXfrm>");

        xml.push_str("</dsp:sp>");
    }

    xml.push_str("</dsp:spTree>");
    xml.push_str("</dsp:drawing>");

    xml
}

/// Generate SmartArt colors XML.
pub fn generate_smartart_colors_xml() -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
    );
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_1">"#);

    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");
    xml.push_str("<dgm:catLst><dgm:cat type=\"accent1\" pri=\"10100\"/></dgm:catLst>");

    xml.push_str("<dgm:styleLbl name=\"node0\">");
    xml.push_str("<dgm:fillClrLst meth=\"repeat\"><a:schemeClr val=\"accent1\"/></dgm:fillClrLst>");
    xml.push_str("<dgm:linClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:linClrLst>");
    xml.push_str("<dgm:effectClrLst/>");
    xml.push_str("<dgm:txLinClrLst/>");
    xml.push_str("<dgm:txFillClrLst meth=\"repeat\"><a:schemeClr val=\"lt1\"/></dgm:txFillClrLst>");
    xml.push_str("<dgm:txEffectClrLst/>");
    xml.push_str("</dgm:styleLbl>");

    xml.push_str("</dgm:colorsDef>");

    xml
}

/// Generate SmartArt quick style XML.
pub fn generate_smartart_quickstyle_xml() -> String {
    let mut xml = String::with_capacity(1024);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" "#,
    );
    xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1">"#);

    xml.push_str("<dgm:title val=\"\"/>");
    xml.push_str("<dgm:desc val=\"\"/>");
    xml.push_str("<dgm:catLst><dgm:cat type=\"simple\" pri=\"10000\"/></dgm:catLst>");

    xml.push_str("<dgm:styleLbl name=\"node0\">");
    xml.push_str("<dgm:scene3d><a:camera prst=\"orthographicFront\"/><a:lightRig rig=\"threePt\" dir=\"t\"/></dgm:scene3d>");
    xml.push_str("<dgm:txPr/>");
    xml.push_str("<dgm:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></a:fontRef></dgm:style>");
    xml.push_str("</dgm:styleLbl>");

    xml.push_str("</dgm:styleDef>");

    xml
}

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
}
