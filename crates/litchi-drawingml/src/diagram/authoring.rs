//! Authoring `SmartArt` graphics from a text outline (record 0284).
//!
//! A [`Spec`] is a built-in [`Layout`] and a tree of [`Node`]s. It writes the
//! five diagram parts deterministically: the data model
//! ([`Spec::data_xml`]), the layout definition ([`Spec::layout_xml`]), the
//! quick style ([`quick_style_xml`]), the colors ([`colors_xml`]) and the
//! cached drawing ([`Spec::drawing_xml`]), which Litchi lays out itself so
//! consumers that draw the cache (LibreOffice, viewers) and consumers that lay
//! the diagram out again from the data model see the same nodes.
//!
//! [`Spec::from_parts`] recognizes the parts this module wrote: it answers a
//! `Spec` only when regenerating the data model and the layout definition
//! from what they say reproduces them byte for byte, so a host can rebuild
//! such a graphic after a structural edit without losing anything.
//!
//! Layout identifiers are the built-in `PowerPoint` ones
//! (`urn:microsoft.com/office/officeart/2005/8/layout/…`); the embedded
//! layout definitions are Litchi's own and simpler than `PowerPoint`'s.
//!
//! ```
//! use litchi_drawingml::diagram::{Limits, authoring::{Layout, Spec}};
//!
//! let limits = Limits::default();
//! let spec = Spec::parse_outline(Layout::Process, "Plan\nBuild\n\tTest\nShip", &limits).unwrap();
//! assert_eq!(spec.nodes().len(), 3);
//! let data = spec.data_xml(Some("rId9"), &limits).unwrap();
//! assert!(data.contains("layout/process1"));
//! assert_eq!(spec.outline(&limits).unwrap(), "Plan\nBuild\n\tTest\nShip");
//! ```

mod drawing;
#[cfg(test)]
mod tests;

use litchi_core::xml::escape_xml;

use super::Limits;
use super::edit::{DRAWING_NAMESPACE, Outline};
use crate::{Error, Result};

const MAX_AUTHORING_OUTPUT: usize = 16 * 1024 * 1024;

fn ensure_input(bytes: usize, limits: &Limits, kind: crate::LimitKind) -> Result<()> {
    let maximum = limits.input_bytes_value().min(MAX_AUTHORING_OUTPUT);
    if bytes > maximum {
        return Err(Error::limit(kind, bytes, maximum));
    }
    Ok(())
}

fn ensure_output(bytes: usize, limits: &Limits, kind: crate::LimitKind) -> Result<()> {
    let maximum = limits.output_bytes_value().min(MAX_AUTHORING_OUTPUT);
    if bytes > maximum {
        return Err(Error::limit(kind, bytes, maximum));
    }
    Ok(())
}

fn estimated_output(
    base: usize,
    per_node: usize,
    nodes: usize,
    text_bytes: usize,
    text_multiplier: usize,
    limits: &Limits,
    kind: crate::LimitKind,
) -> Result<usize> {
    let text = text_bytes
        .checked_mul(text_multiplier)
        .ok_or_else(|| Error::Invalid("SmartArt output length overflows".to_owned()))?;
    let dynamic = per_node
        .checked_mul(nodes)
        .and_then(|value| value.checked_add(text))
        .and_then(|value| value.checked_add(base))
        .ok_or_else(|| Error::Invalid("SmartArt output length overflows".to_owned()))?;
    ensure_output(dynamic, limits, kind)?;
    Ok(dynamic)
}

fn escaped_xml_len(text: &str) -> Result<usize> {
    text.as_bytes().iter().try_fold(0_usize, |length, byte| {
        let replacement = match byte {
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        length
            .checked_add(replacement)
            .ok_or_else(|| Error::Invalid("escaped SmartArt text length overflows".to_owned()))
    })
}

fn text_stats(flat: &[Flat<'_>]) -> Result<(usize, usize)> {
    flat.iter()
        .try_fold((0_usize, 0_usize), |(bytes, lines), item| {
            let next_bytes = bytes
                .checked_add(item.node.text.len())
                .ok_or_else(|| Error::Invalid("SmartArt text length overflows".to_owned()))?;
            let next_lines = lines
                .checked_add(item.node.text.bytes().filter(|byte| *byte == b'\n').count())
                .and_then(|value| value.checked_add(1))
                .ok_or_else(|| Error::Invalid("SmartArt line count overflows".to_owned()))?;
            Ok((next_bytes, next_lines))
        })
}

/// Most nodes a spec holds.
pub const MAX_NODES: usize = 1000;
/// Most nesting levels below the top level a spec holds.
pub const MAX_DEPTH: u32 = 7;
/// Most text bytes of one node.
pub const MAX_NODE_TEXT: usize = 32 * 1024;

const LAYOUT_PREFIX: &str = "urn:microsoft.com/office/officeart/2005/8/layout/";
const QUICK_STYLE_ID: &str = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
const COLORS_ID: &str = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";
const DGM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// A built-in `SmartArt` layout Litchi authors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Layout {
    /// Basic Block List (`default`): top-level nodes in a grid of boxes.
    List,
    /// Basic Process (`process1`): boxes in a row joined by arrows.
    Process,
    /// Basic Cycle (`cycle2`): circles around a ring.
    Cycle,
    /// Hierarchy (`hierarchy1`): every node a box, children below parents.
    Hierarchy,
    /// Basic Pyramid (`pyramid1`): stacked slices, the first on top.
    Pyramid,
    /// Basic Matrix (`matrix3`): boxes in a two-column grid.
    Matrix,
    /// Basic Venn (`venn1`): overlapping circles.
    Venn,
}

impl Layout {
    /// Every layout.
    pub const ALL: [Self; 7] = [
        Self::List,
        Self::Process,
        Self::Cycle,
        Self::Hierarchy,
        Self::Pyramid,
        Self::Matrix,
        Self::Venn,
    ];

    /// The layout's token: `list`, `process`, `cycle`, `hierarchy`,
    /// `pyramid`, `matrix` or `venn`.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::List => "list",
            Self::Process => "process",
            Self::Cycle => "cycle",
            Self::Hierarchy => "hierarchy",
            Self::Pyramid => "pyramid",
            Self::Matrix => "matrix",
            Self::Venn => "venn",
        }
    }

    /// The layout of a token ([`Self::name`], ASCII case-insensitive).
    #[must_use]
    pub fn parse(token: &str) -> Option<Self> {
        Self::ALL
            .into_iter()
            .find(|layout| layout.name().eq_ignore_ascii_case(token.trim()))
    }

    /// The last segment of the layout's unique identifier.
    const fn id_suffix(self) -> &'static str {
        match self {
            Self::List => "default",
            Self::Process => "process1",
            Self::Cycle => "cycle2",
            Self::Hierarchy => "hierarchy1",
            Self::Pyramid => "pyramid1",
            Self::Matrix => "matrix3",
            Self::Venn => "venn1",
        }
    }

    /// The layout's unique identifier (`loTypeId`, `layoutDef/@uniqueId`).
    #[must_use]
    pub fn unique_id(self) -> String {
        format!("{LAYOUT_PREFIX}{}", self.id_suffix())
    }

    /// The layout whose unique identifier is exactly `id`.
    #[must_use]
    pub fn from_unique_id(id: &str) -> Option<Self> {
        let suffix = id.strip_prefix(LAYOUT_PREFIX)?;
        Self::ALL
            .into_iter()
            .find(|layout| layout.id_suffix() == suffix)
    }

    /// The layout category (`loCatId`).
    #[must_use]
    pub const fn category(self) -> &'static str {
        match self {
            Self::List => "list",
            Self::Process => "process",
            Self::Cycle => "cycle",
            Self::Hierarchy => "hierarchy",
            Self::Pyramid => "pyramid",
            Self::Matrix => "matrix",
            Self::Venn => "relationship",
        }
    }

    /// Whether every node is drawn as its own shape; otherwise only the
    /// top-level nodes are, with their descendants as bullet lines.
    const fn draws_every_node(self) -> bool {
        matches!(self, Self::Hierarchy)
    }
}

/// One node of a [`Spec`]: its text and its children.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Node {
    /// The node's text; each line is a paragraph.
    pub text: String,
    /// The node's children, in order.
    pub children: Vec<Node>,
}

impl Node {
    /// A node without children.
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            children: Vec::new(),
        }
    }

    /// Append a child.
    #[must_use]
    pub fn with_child(mut self, child: Self) -> Self {
        self.children.push(child);
        self
    }
}

/// A `SmartArt` graphic to author: a layout and a node tree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Spec {
    layout: Layout,
    nodes: Vec<Node>,
}

/// One flattened node: its number (`modelId`), parent number (0: the
/// document), sibling position, depth and the number of the top-level node
/// or (for layouts that draw every node) the node itself that presents it.
#[derive(Debug, Clone, Copy)]
pub(super) struct Flat<'s> {
    pub(super) id: u32,
    pub(super) parent: u32,
    pub(super) order: u32,
    pub(super) depth: u32,
    pub(super) node: &'s Node,
}

impl Spec {
    /// An empty graphic with `layout`.
    #[must_use]
    pub const fn new(layout: Layout) -> Self {
        Self {
            layout,
            nodes: Vec::new(),
        }
    }

    /// Append a top-level node.
    #[must_use]
    pub fn with_node(mut self, node: Node) -> Self {
        self.nodes.push(node);
        self
    }

    /// The layout.
    #[must_use]
    pub const fn layout(&self) -> Layout {
        self.layout
    }

    /// Replace the layout.
    #[must_use]
    pub const fn with_layout(mut self, layout: Layout) -> Self {
        self.layout = layout;
        self
    }

    /// The top-level nodes.
    #[must_use]
    pub fn nodes(&self) -> &[Node] {
        &self.nodes
    }

    /// The top-level nodes, to edit.
    pub fn nodes_mut(&mut self) -> &mut Vec<Node> {
        &mut self.nodes
    }

    /// A graphic from a text outline: one node per non-blank line, nested by
    /// its indentation (one level per leading tab or per two leading spaces).
    /// OfficeCLI's escapes `\n` and `\t` are read as a line break and a tab.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for a line indented more than one level
    /// below the previous one, or a graphic [`Self::validate`] refuses.
    pub fn parse_outline(layout: Layout, text: &str, limits: &Limits) -> Result<Self> {
        limits.validate()?;
        ensure_input(text.len(), limits, crate::LimitKind::InputBytes)?;
        let mut normalized = String::with_capacity(text.len());
        let mut chars = text.chars().peekable();
        while let Some(character) = chars.next() {
            if character == '\\' {
                match chars.peek().copied() {
                    Some('n') => {
                        chars.next();
                        normalized.push('\n');
                        continue;
                    },
                    Some('t') => {
                        chars.next();
                        normalized.push('\t');
                        continue;
                    },
                    _ => {},
                }
            }
            normalized.push(character);
        }
        let mut spec = Self::new(layout);
        // The path of child indexes to the last node at each depth.
        let path_capacity = limits.node_depth_value().min(MAX_DEPTH as usize);
        let mut path: Vec<usize> = Vec::with_capacity(path_capacity);
        let max_nodes = limits.nodes_value().min(MAX_NODES);
        let max_depth = limits.node_depth_value().min(MAX_DEPTH as usize);
        let max_text = limits.node_text_bytes_value().min(MAX_NODE_TEXT);
        let mut node_count = 0_usize;
        for line in normalized.split('\n') {
            let line = line.strip_suffix('\r').unwrap_or(line);
            if line.trim().is_empty() {
                continue;
            }
            let mut level = 0_usize;
            let mut rest = line;
            loop {
                if let Some(next) = rest.strip_prefix('\t') {
                    rest = next;
                } else if let Some(next) = rest.strip_prefix("  ") {
                    rest = next;
                } else {
                    break;
                }
                level += 1;
            }
            let rest = rest.trim();
            if level > max_depth {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeDepth,
                    level,
                    max_depth,
                ));
            }
            if level > path.len() {
                return Err(Error::Invalid(format!(
                    "outline line \"{rest}\" is indented more than one level below the line before it"
                )));
            }
            node_count = node_count
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("SmartArt node count overflows".to_owned()))?;
            if node_count > max_nodes {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodes,
                    node_count,
                    max_nodes,
                ));
            }
            if rest.len() > max_text {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeTextBytes,
                    rest.len(),
                    max_text,
                ));
            }
            if rest.chars().any(|character| {
                !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
            }) {
                return Err(Error::Invalid(
                    "SmartArt node text holds a character XML cannot carry".to_owned(),
                ));
            }
            path.truncate(level);
            let mut siblings = &mut spec.nodes;
            for &index in &path {
                siblings = &mut siblings[index].children;
            }
            siblings.push(Node::new(rest));
            path.push(siblings.len() - 1);
        }
        spec.validate_for_limits(limits)?;
        Ok(spec)
    }

    /// The text outline of the graphic: one line per node, indented by one
    /// tab per level; a node's own line breaks become spaces.
    #[must_use]
    pub fn outline(&self, limits: &Limits) -> Result<String> {
        self.validate_for_limits(limits)?;
        let mut capacity = 0_usize;
        let mut first = true;
        fn measure(
            nodes: &[Node],
            depth: usize,
            capacity: &mut usize,
            first: &mut bool,
        ) -> Result<()> {
            for node in nodes {
                if !*first {
                    *capacity = capacity.checked_add(1).ok_or_else(|| {
                        Error::Invalid("SmartArt outline length overflows".to_owned())
                    })?;
                }
                *first = false;
                *capacity = capacity
                    .checked_add(node.text.len())
                    .and_then(|value| value.checked_add(depth))
                    .ok_or_else(|| {
                        Error::Invalid("SmartArt outline length overflows".to_owned())
                    })?;
                measure(&node.children, depth + 1, capacity, first)?;
            }
            Ok(())
        }
        measure(&self.nodes, 0, &mut capacity, &mut first)?;
        ensure_output(capacity, limits, crate::LimitKind::OutputBytes)?;
        fn walk(out: &mut String, nodes: &[Node], depth: usize) {
            for node in nodes {
                if !out.is_empty() {
                    out.push('\n');
                }
                for _ in 0..depth {
                    out.push('\t');
                }
                out.push_str(&node.text.replace(['\r', '\n'], " "));
                walk(out, &node.children, depth + 1);
            }
        }
        let mut out = String::with_capacity(capacity);
        walk(&mut out, &self.nodes, 0);
        ensure_output(out.len(), limits, crate::LimitKind::OutputBytes)?;
        Ok(out)
    }

    /// Check the node count, depth and text against the module's bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] for too many nodes, levels or text bytes,
    /// and [`Error::Invalid`] for text XML cannot carry.
    pub fn validate(&self) -> Result<()> {
        let mut count = 0_usize;
        let mut stack: Vec<(&Node, u32)> = self.nodes.iter().map(|node| (node, 0)).collect();
        while let Some((node, depth)) = stack.pop() {
            count += 1;
            if count > MAX_NODES {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodes,
                    count,
                    MAX_NODES,
                ));
            }
            if depth > MAX_DEPTH {
                let observed = usize::try_from(depth)
                    .map_err(|_| Error::Invalid("SmartArt depth exceeds usize".into()))?;
                let maximum = usize::try_from(MAX_DEPTH)
                    .map_err(|_| Error::Invalid("SmartArt depth limit exceeds usize".into()))?;
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeDepth,
                    observed,
                    maximum,
                ));
            }
            if node.text.len() > MAX_NODE_TEXT {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeTextBytes,
                    node.text.len(),
                    MAX_NODE_TEXT,
                ));
            }
            if node.text.chars().any(|character| {
                !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
            }) {
                return Err(Error::Invalid(
                    "SmartArt node text holds a character XML cannot carry".to_owned(),
                ));
            }
            stack.extend(node.children.iter().map(|child| (child, depth + 1)));
        }
        Ok(())
    }

    fn validate_for_limits(&self, limits: &Limits) -> Result<()> {
        limits.validate()?;
        self.validate()?;
        let max_nodes = limits.nodes_value().min(MAX_NODES);
        let max_depth = limits.node_depth_value().min(MAX_DEPTH as usize);
        let max_text = limits.node_text_bytes_value().min(MAX_NODE_TEXT);
        let mut count = 0_usize;
        let mut stack: Vec<(&Node, usize)> =
            self.nodes.iter().rev().map(|node| (node, 0)).collect();
        while let Some((node, depth)) = stack.pop() {
            count = count
                .checked_add(1)
                .ok_or_else(|| Error::Invalid("SmartArt node count overflows".to_owned()))?;
            if count > max_nodes {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodes,
                    count,
                    max_nodes,
                ));
            }
            if depth > max_depth {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeDepth,
                    depth,
                    max_depth,
                ));
            }
            if node.text.len() > max_text {
                return Err(Error::limit(
                    crate::LimitKind::SmartArtNodeTextBytes,
                    node.text.len(),
                    max_text,
                ));
            }
            stack.extend(node.children.iter().rev().map(|child| (child, depth + 1)));
        }
        Ok(())
    }

    /// The nodes in document order with their numbers.
    pub(super) fn flatten(&self) -> Vec<Flat<'_>> {
        fn walk<'s>(out: &mut Vec<Flat<'s>>, nodes: &'s [Node], parent: u32, depth: u32) {
            for (order, node) in nodes.iter().enumerate() {
                let id = u32::try_from(out.len() + 1).unwrap_or(u32::MAX);
                out.push(Flat {
                    id,
                    parent,
                    order: u32::try_from(order).unwrap_or(u32::MAX),
                    depth,
                    node,
                });
                walk(out, &node.children, id, depth + 1);
            }
        }
        let mut out = Vec::new();
        walk(&mut out, &self.nodes, 0, 0);
        out
    }

    /// A spec equal to the parts `data` and `layout`, when this module wrote
    /// them: regenerating both from what they say reproduces them exactly.
    /// `None` for any other graphic.
    ///
    /// # Errors
    ///
    /// Returns an error when `data` is not a readable diagram data model.
    pub fn from_parts(data: &[u8], layout: &[u8], limits: &Limits) -> Result<Option<Self>> {
        Self::recognize(&Outline::read(data, limits)?, data, layout, limits)
    }

    /// [`Self::from_parts`] for a data model already read as `outline`.
    ///
    /// # Errors
    ///
    /// Returns an error when regenerating the data model fails.
    pub fn recognize(
        outline: &Outline,
        data: &[u8],
        layout: &[u8],
        limits: &Limits,
    ) -> Result<Option<Self>> {
        limits.validate()?;
        ensure_input(data.len(), limits, crate::LimitKind::DiagramDataModelBytes)?;
        ensure_input(layout.len(), limits, crate::LimitKind::XmlBytes)?;
        let Some(kind) = outline
            .layout_type_id
            .as_deref()
            .and_then(Layout::from_unique_id)
        else {
            return Ok(None);
        };
        let spec = Self::from_outline(kind, outline);
        if spec.validate().is_err() {
            return Ok(None);
        }
        let regenerated = spec.data_xml(outline.drawing_relationship.as_deref(), limits)?;
        if regenerated.as_bytes() != data || spec.layout_xml(limits)?.as_bytes() != layout {
            return Ok(None);
        }
        Ok(Some(spec))
    }

    /// The node tree an outline reads, with `layout`.
    #[must_use]
    pub fn from_outline(layout: Layout, outline: &Outline) -> Self {
        fn build(outline: &Outline, parent: Option<usize>) -> Vec<Node> {
            outline
                .children(parent)
                .map(|index| Node {
                    text: outline.entries[index].text.clone(),
                    children: build(outline, Some(index)),
                })
                .collect()
        }
        Self {
            layout,
            nodes: build(outline, None),
        }
    }

    /// The data-model part (`dgm:dataModel`). `drawing_relationship` is the
    /// host relationship of the cached drawing, recorded in the model's
    /// `dsp:dataModelExt`.
    ///
    /// # Errors
    ///
    /// Returns an error when the spec fails [`Self::validate`].
    pub fn data_xml(&self, drawing_relationship: Option<&str>, limits: &Limits) -> Result<String> {
        self.validate_for_limits(limits)?;
        let flat = self.flatten();
        let (text_bytes, line_count) = text_stats(&flat)?;
        let text_bound = text_bytes
            .checked_add(
                line_count
                    .checked_mul(128)
                    .ok_or_else(|| Error::Invalid("SmartArt line count overflows".to_owned()))?,
            )
            .ok_or_else(|| Error::Invalid("SmartArt text length overflows".to_owned()))?;
        let relationship_bytes = drawing_relationship
            .map(escaped_xml_len)
            .transpose()?
            .unwrap_or(0);
        let capacity = estimated_output(
            8 * 1024 + relationship_bytes,
            5 * 1024,
            flat.len(),
            text_bound,
            6,
            limits,
            crate::LimitKind::DiagramDataModelBytes,
        )?;
        let layout_id = self.layout.unique_id();
        let mut xml = String::with_capacity(capacity);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(&format!(
            r#"<dgm:dataModel xmlns:dgm="{DGM}" xmlns:a="{A}"><dgm:ptLst>"#
        ));
        xml.push_str(&format!(
            r#"<dgm:pt modelId="0" type="doc"><dgm:prSet loTypeId="{layout_id}" loCatId="{}" qsTypeId="{QUICK_STYLE_ID}" qsCatId="simple" csTypeId="{COLORS_ID}" csCatId="accent1" phldr="0"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>"#,
            self.layout.category()
        ));
        for item in &flat {
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/>"#,
                item.id
            ));
            push_paragraphs(&mut xml, &item.node.text);
            xml.push_str("</dgm:t></dgm:pt>");
            for (base, kind) in [(20_000, "parTrans"), (30_000, "sibTrans")] {
                xml.push_str(&format!(
                    r#"<dgm:pt modelId="{}" type="{kind}" cxnId="{}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></dgm:t></dgm:pt>"#,
                    base + item.id,
                    10_000 + item.id
                ));
            }
        }
        // Presentation points.
        let shown: Vec<&Flat<'_>> = flat
            .iter()
            .filter(|item| item.depth == 0 || self.layout.draws_every_node())
            .collect();
        if !flat.is_empty() {
            xml.push_str(r#"<dgm:pt modelId="40000" type="pres"><dgm:prSet presAssocID="0" presName="diagram" presStyleCnt="0"><dgm:presLayoutVars><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#);
        }
        for (index, item) in shown.iter().enumerate() {
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="pres"><dgm:prSet presAssocID="{}" presName="node" presStyleLbl="node1" presStyleIdx="{index}" presStyleCnt="{}"><dgm:presLayoutVars><dgm:bulletEnabled val="1"/></dgm:presLayoutVars></dgm:prSet><dgm:spPr/></dgm:pt>"#,
                40_000 + item.id,
                item.id,
                shown.len()
            ));
        }
        let transitions = self.transitions(&flat);
        for (index, (item, kind)) in transitions.iter().enumerate() {
            let (name, label, base) = match kind {
                Transition::Sibling => ("sibTrans", "sibTrans2D1", 30_000),
                Transition::Parent => ("parTrans", "parChTrans1D1", 20_000),
            };
            xml.push_str(&format!(
                r#"<dgm:pt modelId="{}" type="pres"><dgm:prSet presAssocID="{}" presName="{name}" presStyleLbl="{label}" presStyleIdx="{index}" presStyleCnt="{}"/><dgm:spPr/></dgm:pt>"#,
                70_000 + item.id,
                base + item.id,
                transitions.len()
            ));
        }
        xml.push_str("</dgm:ptLst><dgm:cxnLst>");
        for item in &flat {
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" parTransId="{}" sibTransId="{}"/>"#,
                10_000 + item.id,
                item.parent,
                item.id,
                item.order,
                20_000 + item.id,
                30_000 + item.id
            ));
        }
        if !flat.is_empty() {
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="50000" type="presOf" srcId="0" destId="40000" srcOrd="0" destOrd="0" presId="{layout_id}"/>"#
            ));
        }
        // presOf: each node to the shape that shows it, in that shape's order.
        let mut shown_order: std::collections::HashMap<u32, u32> = std::collections::HashMap::new();
        for item in &flat {
            let presenter = if self.layout.draws_every_node() {
                item.id
            } else {
                top_ancestor(&flat, item)
            };
            let order = shown_order.entry(presenter).or_insert(0);
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="presOf" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" presId="{layout_id}"/>"#,
                50_000 + item.id,
                item.id,
                40_000 + presenter,
                *order
            ));
            *order += 1;
        }
        for (item, kind) in &transitions {
            let source = match kind {
                Transition::Sibling => 30_000 + item.id,
                Transition::Parent => 20_000 + item.id,
            };
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="presOf" srcId="{source}" destId="{}" srcOrd="0" destOrd="0" presId="{layout_id}"/>"#,
                80_000 + item.id,
                70_000 + item.id
            ));
        }
        // presParOf: the presentation tree.
        let mut order = 0_u32;
        for item in &shown {
            if self.layout.draws_every_node() && item.parent != 0 {
                continue;
            }
            xml.push_str(&format!(
                r#"<dgm:cxn modelId="{}" type="presParOf" srcId="40000" destId="{}" srcOrd="{order}" destOrd="0" presId="{layout_id}"/>"#,
                60_000 + item.id,
                40_000 + item.id
            ));
            order += 1;
            if let Some((transition, _)) = transitions
                .iter()
                .find(|(other, kind)| other.id == item.id && *kind == Transition::Sibling)
            {
                xml.push_str(&format!(
                    r#"<dgm:cxn modelId="{}" type="presParOf" srcId="40000" destId="{}" srcOrd="{order}" destOrd="0" presId="{layout_id}"/>"#,
                    90_000 + transition.id,
                    70_000 + transition.id
                ));
                order += 1;
            }
        }
        if self.layout.draws_every_node() {
            let mut child_order: std::collections::HashMap<u32, u32> =
                std::collections::HashMap::new();
            for item in flat.iter().filter(|item| item.parent != 0) {
                let next = child_order.entry(item.parent).or_insert(0);
                xml.push_str(&format!(
                    r#"<dgm:cxn modelId="{}" type="presParOf" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" presId="{layout_id}"/>"#,
                    90_000 + item.id,
                    40_000 + item.parent,
                    70_000 + item.id,
                    *next
                ));
                xml.push_str(&format!(
                    r#"<dgm:cxn modelId="{}" type="presParOf" srcId="{}" destId="{}" srcOrd="{}" destOrd="0" presId="{layout_id}"/>"#,
                    60_000 + item.id,
                    40_000 + item.parent,
                    40_000 + item.id,
                    *next + 1
                ));
                *next += 2;
            }
        }
        xml.push_str("</dgm:cxnLst><dgm:bg/><dgm:whole/>");
        if let Some(relationship) = drawing_relationship {
            xml.push_str(&format!(
                r#"<dgm:extLst><a:ext uri="{DRAWING_NAMESPACE}"><dsp:dataModelExt xmlns:dsp="{DRAWING_NAMESPACE}" relId="{}" minVer="{DGM}"/></a:ext></dgm:extLst>"#,
                escape_xml(relationship)
            ));
        }
        xml.push_str("</dgm:dataModel>");
        ensure_output(xml.len(), limits, crate::LimitKind::DiagramDataModelBytes)?;
        Ok(xml)
    }

    /// The transition shapes the layout draws: an arrow after every
    /// top-level process step but the last, a link above every non-root
    /// hierarchy node.
    fn transitions<'f, 's>(&self, flat: &'f [Flat<'s>]) -> Vec<(&'f Flat<'s>, Transition)> {
        match self.layout {
            Layout::Process => {
                let top: Vec<&Flat<'s>> = flat.iter().filter(|item| item.depth == 0).collect();
                top.iter()
                    .take(top.len().saturating_sub(1))
                    .map(|item| (*item, Transition::Sibling))
                    .collect()
            },
            Layout::Hierarchy => flat
                .iter()
                .filter(|item| item.parent != 0)
                .map(|item| (item, Transition::Parent))
                .collect(),
            _ => Vec::new(),
        }
    }

    /// The layout-definition part (`dgm:layoutDef`).
    #[must_use]
    pub fn layout_xml(&self, limits: &Limits) -> Result<String> {
        limits.validate()?;
        estimated_output(16 * 1024, 0, 0, 0, 1, limits, crate::LimitKind::XmlBytes)?;
        let xml = layout_definition(self.layout);
        ensure_output(xml.len(), limits, crate::LimitKind::XmlBytes)?;
        Ok(xml)
    }

    /// The cached drawing part (`dsp:drawing`) for a frame of `width` ×
    /// `height` EMU; shape positions are relative to the frame.
    ///
    /// # Errors
    ///
    /// Returns an error for non-positive extents or a spec that fails
    /// [`Self::validate`].
    pub fn drawing_xml(&self, width: i64, height: i64, limits: &Limits) -> Result<String> {
        self.validate_for_limits(limits)?;
        if width <= 0 || height <= 0 {
            return Err(Error::Invalid(
                "SmartArt frame extents must be positive".to_owned(),
            ));
        }
        let flat = self.flatten();
        let (text_bytes, line_count) = text_stats(&flat)?;
        let text_bound = text_bytes
            .checked_add(
                line_count
                    .checked_mul(128)
                    .ok_or_else(|| Error::Invalid("SmartArt line count overflows".to_owned()))?,
            )
            .ok_or_else(|| Error::Invalid("SmartArt text length overflows".to_owned()))?;
        let capacity = estimated_output(
            8 * 1024,
            6 * 1024,
            flat.len(),
            text_bound,
            12,
            limits,
            crate::LimitKind::DiagramDrawingBytes,
        )?;
        let xml = drawing::write(self, width, height, capacity);
        ensure_output(xml.len(), limits, crate::LimitKind::DiagramDrawingBytes)?;
        Ok(xml)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Transition {
    Sibling,
    Parent,
}

fn top_ancestor(flat: &[Flat<'_>], item: &Flat<'_>) -> u32 {
    let mut current = *item;
    while current.parent != 0 {
        current = flat[current.parent as usize - 1];
    }
    current.id
}

/// One `a:p` per line of `text`.
fn push_paragraphs(xml: &mut String, text: &str) {
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        if line.is_empty() {
            xml.push_str(r#"<a:p><a:endParaRPr lang="en-US"/></a:p>"#);
        } else {
            xml.push_str(&format!(
                r#"<a:p><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p>"#,
                escape_xml(line)
            ));
        }
    }
}

/// The quick-style part (`dgm:styleDef`) of authored graphics.
pub fn quick_style_xml(limits: &Limits) -> Result<String> {
    limits.validate()?;
    estimated_output(8 * 1024, 0, 0, 0, 1, limits, crate::LimitKind::XmlBytes)?;
    let mut xml = String::with_capacity(2048);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<dgm:styleDef xmlns:dgm="{DGM}" xmlns:a="{A}" uniqueId="{QUICK_STYLE_ID}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:styleLbl name="node0"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>"#
    ));
    for (name, line, font) in [
        ("node1", "2", "lt1"),
        ("vennNode1", "2", "tx1"),
        ("sibTrans2D1", "0", "lt1"),
        ("parChTrans1D1", "2", "tx1"),
    ] {
        xml.push_str(&format!(
            r#"<dgm:styleLbl name="{name}"><dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style><a:lnRef idx="{line}"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="{font}"/></a:fontRef></dgm:style></dgm:styleLbl>"#
        ));
    }
    xml.push_str("</dgm:styleDef>");
    ensure_output(xml.len(), limits, crate::LimitKind::XmlBytes)?;
    Ok(xml)
}

/// The colors part (`dgm:colorsDef`) of authored graphics.
pub fn colors_xml(limits: &Limits) -> Result<String> {
    limits.validate()?;
    estimated_output(8 * 1024, 0, 0, 0, 1, limits, crate::LimitKind::XmlBytes)?;
    let mut xml = String::with_capacity(2048);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<dgm:colorsDef xmlns:dgm="{DGM}" xmlns:a="{A}" uniqueId="{COLORS_ID}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>"#
    ));
    for (name, fill, line, text) in [
        (
            "node0",
            r#"<a:schemeClr val="accent1"/>"#,
            r#"<a:schemeClr val="lt1"/>"#,
            r#"<a:schemeClr val="lt1"/>"#,
        ),
        (
            "node1",
            r#"<a:schemeClr val="accent1"/>"#,
            r#"<a:schemeClr val="lt1"/>"#,
            r#"<a:schemeClr val="lt1"/>"#,
        ),
        (
            "vennNode1",
            r#"<a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>"#,
            r#"<a:schemeClr val="lt1"/>"#,
            r#"<a:schemeClr val="tx1"/>"#,
        ),
        (
            "sibTrans2D1",
            r#"<a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr>"#,
            r#"<a:schemeClr val="accent1"><a:tint val="60000"/></a:schemeClr>"#,
            r#"<a:schemeClr val="lt1"/>"#,
        ),
        (
            "parChTrans1D1",
            r#"<a:schemeClr val="accent1"/>"#,
            r#"<a:schemeClr val="accent1"><a:shade val="60000"/></a:schemeClr>"#,
            r#"<a:schemeClr val="tx1"/>"#,
        ),
    ] {
        xml.push_str(&format!(
            r#"<dgm:styleLbl name="{name}"><dgm:fillClrLst meth="repeat">{fill}</dgm:fillClrLst><dgm:linClrLst meth="repeat">{line}</dgm:linClrLst><dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat">{text}</dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>"#
        ));
    }
    xml.push_str("</dgm:colorsDef>");
    ensure_output(xml.len(), limits, crate::LimitKind::XmlBytes)?;
    Ok(xml)
}

/// Litchi's layout definition for `layout`.
fn layout_definition(layout: Layout) -> String {
    let (algorithm, shape, node_label) = match layout {
        Layout::List => (
            r#"<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="off" val="ctr"/></dgm:alg>"#,
            "rect",
            "node1",
        ),
        Layout::Process => (r#"<dgm:alg type="lin"/>"#, "roundRect", "node1"),
        Layout::Cycle => (
            r#"<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>"#,
            "ellipse",
            "node1",
        ),
        Layout::Hierarchy => (r#"<dgm:alg type="hierChild"/>"#, "roundRect", "node1"),
        Layout::Pyramid => (
            r#"<dgm:alg type="pyra"><dgm:param type="linDir" val="fromT"/></dgm:alg>"#,
            "trapezoid",
            "node1",
        ),
        Layout::Matrix => (
            r#"<dgm:alg type="snake"><dgm:param type="grDir" val="tL"/><dgm:param type="flowDir" val="row"/><dgm:param type="contDir" val="sameDir"/><dgm:param type="bkpt" val="fixed"/><dgm:param type="bkPtFixedVal" val="2"/></dgm:alg>"#,
            "roundRect",
            "node1",
        ),
        Layout::Venn => (
            r#"<dgm:alg type="cycle"><dgm:param type="stAng" val="0"/><dgm:param type="spanAng" val="360"/></dgm:alg>"#,
            "ellipse",
            "vennNode1",
        ),
    };
    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(&format!(
        r#"<dgm:layoutDef xmlns:dgm="{DGM}" xmlns:a="{A}" uniqueId="{}"><dgm:title val=""/><dgm:desc val=""/><dgm:catLst><dgm:cat type="{}" pri="1000"/></dgm:catLst>"#,
        layout.unique_id(),
        layout.category()
    ));
    xml.push_str(r#"<dgm:sampData><dgm:dataModel><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="1"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="2"><dgm:prSet phldr="1"/></dgm:pt><dgm:pt modelId="3"><dgm:prSet phldr="1"/></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="4" srcId="0" destId="1" srcOrd="0" destOrd="0"/><dgm:cxn modelId="5" srcId="0" destId="2" srcOrd="1" destOrd="0"/><dgm:cxn modelId="6" srcId="0" destId="3" srcOrd="2" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel></dgm:sampData>"#);
    xml.push_str(r#"<dgm:layoutNode name="diagram"><dgm:varLst><dgm:dir/><dgm:resizeHandles val="exact"/></dgm:varLst>"#);
    xml.push_str(algorithm);
    xml.push_str(&format!(
        r#"<dgm:shape xmlns:r="{R}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/><dgm:constr type="h" for="ch" forName="node" refType="w" refFor="ch" refForName="node" fact="0.6"/><dgm:constr type="w" for="ch" forName="sibTrans" refType="w" refFor="ch" refForName="node" fact="0.3"/><dgm:constr type="h" for="ch" forName="sibTrans" refType="h" refFor="ch" refForName="node" fact="0.4"/><dgm:constr type="sp" refType="w" refFor="ch" refForName="node" fact="0.1"/><dgm:constr type="primFontSz" for="ch" forName="node" op="equ" val="65"/></dgm:constrLst><dgm:ruleLst/>"#
    ));
    let node = format!(
        r#"<dgm:layoutNode name="node" styleLbl="{node_label}"><dgm:varLst><dgm:bulletEnabled val="1"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape xmlns:r="{R}" type="{shape}" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="{}" ptType="node"/><dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst><dgm:ruleLst><dgm:rule type="primFontSz" val="5" fact="NaN" max="NaN"/></dgm:ruleLst>"#,
        if layout.draws_every_node() {
            "self"
        } else {
            "desOrSelf"
        }
    );
    match layout {
        Layout::Process => {
            xml.push_str(r#"<dgm:forEach name="nodesForEach" axis="ch" ptType="node">"#);
            xml.push_str(&node);
            xml.push_str("</dgm:layoutNode>");
            xml.push_str(&format!(
                r#"<dgm:forEach name="sibTransForEach" axis="followSib" ptType="sibTrans" cnt="1"><dgm:layoutNode name="sibTrans" styleLbl="sibTrans2D1"><dgm:alg type="conn"/><dgm:shape xmlns:r="{R}" type="rightArrow" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach>"#
            ));
            xml.push_str("</dgm:forEach>");
        },
        Layout::Hierarchy => {
            xml.push_str(r#"<dgm:forEach name="rootForEach" axis="ch" ptType="node">"#);
            xml.push_str(r#"<dgm:layoutNode name="hierRoot"><dgm:alg type="hierRoot"/><dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/>"#);
            xml.push_str(&node);
            xml.push_str("</dgm:layoutNode>");
            xml.push_str(r#"<dgm:layoutNode name="hierChildren"><dgm:alg type="hierChild"/><dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst/><dgm:ruleLst/><dgm:forEach name="childForEach" axis="ch" ptType="node"><dgm:forEach name="parTransForEach" axis="self" ptType="parTrans" st="-1" cnt="1"><dgm:layoutNode name="parTrans" styleLbl="parChTrans1D1"><dgm:alg type="conn"><dgm:param type="dim" val="1D"/><dgm:param type="endSty" val="noArr"/></dgm:alg><dgm:shape xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="conn" r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf axis="self"/><dgm:constrLst/><dgm:ruleLst/></dgm:layoutNode></dgm:forEach><dgm:forEach name="repeat" ref="rootForEach"/></dgm:forEach></dgm:layoutNode>"#);
            xml.push_str("</dgm:layoutNode></dgm:forEach>");
        },
        _ => {
            xml.push_str(r#"<dgm:forEach name="nodesForEach" axis="ch" ptType="node">"#);
            xml.push_str(&node);
            xml.push_str("</dgm:layoutNode></dgm:forEach>");
        },
    }
    xml.push_str("</dgm:layoutNode></dgm:layoutDef>");
    xml
}
