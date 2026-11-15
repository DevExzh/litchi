//! Office Math equations in slide text.
//!
//! PowerPoint writes an equation as `a14:m` inside a paragraph of a text
//! body, in the `mc:Choice Requires="a14"` branch of a markup-compatibility
//! block whose fallback is a picture of the equation. The shape scene reads
//! the fallback, so [`equations`] scans the slide's own XML: every `a14:m`,
//! in document order, with the `cNvPr` ID of the shape that holds it and
//! its OMML rendered as a LaTeX-like formula and as plain text.

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const A14: &[u8] = b"http://schemas.microsoft.com/office/drawing/2010/main";
const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const MATH: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const MAX_DEPTH: usize = 256;
const MAX_EQUATIONS: usize = 4096;
const MAX_NODES: usize = 1_000_000;

/// One equation of a slide.
#[derive(Clone, PartialEq, Eq)]
pub struct Equation {
    shape_id: Option<u32>,
    formula: String,
    text: String,
}

impl std::fmt::Debug for Equation {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Equation")
            .field("shape_id", &self.shape_id)
            .finish_non_exhaustive()
    }
}

impl Equation {
    /// The `cNvPr` ID of the shape whose text holds the equation.
    #[must_use]
    pub const fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    /// The equation as a LaTeX-like formula: fractions as `\frac{a}{b}`,
    /// scripts as `x^{2}` and `x_{i}`, radicals as `\sqrt{x}`, n-ary
    /// operators as `\sum`, `\int` and `\prod` with their limits, and
    /// delimiters with their characters; other structures read as their
    /// text.
    #[must_use]
    pub fn formula(&self) -> &str {
        &self.formula
    }

    /// The equation's text: every `m:t` in document order.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// One OMML element, parsed for rendering.
#[derive(Default)]
struct Node {
    local: Vec<u8>,
    /// `m:val`/`m:chr`-style attribute values: (local name, value).
    attributes: Vec<(Vec<u8>, String)>,
    children: Vec<Node>,
    text: String,
}

impl Node {
    fn child(&self, local: &[u8]) -> Option<&Node> {
        self.children.iter().find(|child| child.local == local)
    }

    fn attribute(&self, local: &[u8]) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(name, _)| name == local)
            .map(|(_, value)| value.as_str())
    }

    /// The `m:val` of a property child (`m:chr m:val="∑"`).
    fn property(&self, properties: &[u8], local: &[u8]) -> Option<&str> {
        self.child(properties)?.child(local)?.attribute(b"val")
    }

    fn text(&self, out: &mut String) {
        if self.local == b"t" {
            out.push_str(&self.text);
        }
        for child in &self.children {
            child.text(out);
        }
    }

    fn render(&self, out: &mut String) {
        let group = |node: Option<&Node>, out: &mut String| {
            out.push('{');
            if let Some(node) = node {
                node.render(out);
            }
            out.push('}');
        };
        match self.local.as_slice() {
            b"t" => out.push_str(&self.text),
            // Properties carry no content.
            local if local.ends_with(b"Pr") => {},
            b"f" => {
                out.push_str("\\frac");
                group(self.child(b"num"), out);
                group(self.child(b"den"), out);
            },
            b"sSup" | b"sSub" | b"sSubSup" | b"sPre" => {
                if let Some(base) = self.child(b"e") {
                    base.render(out);
                }
                if let Some(sub) = self.child(b"sub") {
                    out.push('_');
                    group(Some(sub), out);
                }
                if let Some(sup) = self.child(b"sup") {
                    out.push('^');
                    group(Some(sup), out);
                }
            },
            b"rad" => {
                out.push_str("\\sqrt");
                if let Some(degree) = self.child(b"deg")
                    && !degree.children.is_empty()
                {
                    out.push('[');
                    degree.render(out);
                    out.push(']');
                }
                group(self.child(b"e"), out);
            },
            b"nary" => {
                let operator = match self.property(b"naryPr", b"chr").unwrap_or("∫") {
                    "∑" => "\\sum",
                    "∏" => "\\prod",
                    "∐" => "\\coprod",
                    "∬" => "\\iint",
                    "∭" => "\\iiint",
                    "∮" => "\\oint",
                    "⋃" => "\\bigcup",
                    "⋂" => "\\bigcap",
                    _ => "\\int",
                };
                out.push_str(operator);
                if let Some(sub) = self.child(b"sub").filter(|node| !node.children.is_empty()) {
                    out.push('_');
                    group(Some(sub), out);
                }
                if let Some(sup) = self.child(b"sup").filter(|node| !node.children.is_empty()) {
                    out.push('^');
                    group(Some(sup), out);
                }
                if let Some(base) = self.child(b"e") {
                    out.push(' ');
                    base.render(out);
                }
            },
            b"d" => {
                let begin = self.property(b"dPr", b"begChr").unwrap_or("(").to_owned();
                let end = self.property(b"dPr", b"endChr").unwrap_or(")").to_owned();
                out.push_str(&begin);
                let mut first = true;
                for child in self.children.iter().filter(|child| child.local == b"e") {
                    if !first {
                        out.push(',');
                    }
                    first = false;
                    child.render(out);
                }
                out.push_str(&end);
            },
            _ => {
                for child in &self.children {
                    child.render(out);
                }
            },
        }
    }
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("slide equations: {message}"))
}

/// Every equation of a slide's XML, in document order.
///
/// # Errors
///
/// Returns an error for malformed XML or exceeded bounds.
pub fn equations(xml: &[u8]) -> Result<Vec<Equation>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    // For each open element: whether it is a `p:sp` and the `cNvPr` ID
    // seen inside it.
    let mut shapes: Vec<(usize, Option<u32>)> = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut out = Vec::new();
    // The open OMML subtree of the current equation: its stack of nodes,
    // and for each element open inside it whether it pushed a node.
    let mut math: Option<(usize, Vec<Node>)> = None;
    let mut pushed: Vec<bool> = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let a14 = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == A14);
        let in_math = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MATH);
        let pml = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == PML || value == STRICT_PML);
        nodes += 1;
        if nodes > MAX_NODES {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                observed: nodes,
                maximum: MAX_NODES,
            });
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let local = element.local_name().as_ref().to_vec();
                if let Some((_, stack)) = math.as_mut()
                    && in_math
                {
                    let mut node = Node {
                        local: local.clone(),
                        ..Node::default()
                    };
                    for attribute in element.checked_attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Implicit1_0,
                                reader.decoder(),
                            )
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        node.attributes.push((
                            attribute.key.local_name().as_ref().to_vec(),
                            value.into_owned(),
                        ));
                    }
                    if empty {
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(node);
                        }
                    } else {
                        stack.push(node);
                        pushed.push(true);
                    }
                } else if math.is_some() && !empty {
                    pushed.push(false);
                }
                if pml && local == b"sp" && !empty {
                    shapes.push((depth, None));
                }
                if pml
                    && local == b"cNvPr"
                    && let Some((_, id)) = shapes.last_mut()
                    && id.is_none()
                {
                    for attribute in element.checked_attributes().flatten() {
                        if attribute.key.as_ref() == b"id" {
                            *id = std::str::from_utf8(&attribute.value)
                                .ok()
                                .and_then(|value| value.parse().ok());
                        }
                    }
                }
                if a14 && local == b"m" && math.is_none() && !empty {
                    math = Some((depth, vec![Node::default()]));
                }
                if !empty {
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(invalid("the slide is nested too deeply"));
                    }
                }
            },
            Event::Text(ref text) => {
                if let Some((_, stack)) = math.as_mut()
                    && let Some(node) = stack.last_mut()
                    && node.local == b"t"
                {
                    let decoded = text
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    node.text.push_str(&decoded);
                }
            },
            Event::GeneralRef(ref reference) => {
                if let Some((_, stack)) = math.as_mut()
                    && let Some(node) = stack.last_mut()
                    && node.local == b"t"
                {
                    let name = reference
                        .decode()
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    let resolved = match name.as_ref() {
                        "amp" => Some('&'),
                        "lt" => Some('<'),
                        "gt" => Some('>'),
                        "quot" => Some('"'),
                        "apos" => Some('\''),
                        _ => reference.resolve_char_ref().ok().flatten(),
                    };
                    if let Some(ch) = resolved {
                        node.text.push(ch);
                    }
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced slide XML"))?;
                if let Some((start, stack)) = math.as_mut() {
                    if depth == *start {
                        let root = stack.pop().unwrap_or_default();
                        let mut formula = String::new();
                        root.render(&mut formula);
                        let mut text = String::new();
                        root.text(&mut text);
                        if out.len() >= MAX_EQUATIONS {
                            return Err(Error::Limit {
                                kind: crate::opened::LimitKind::Codec(
                                    crate::opened::CodecLimitKind::Equations,
                                ),
                                observed: out.len().checked_add(1).unwrap_or(usize::MAX),
                                maximum: MAX_EQUATIONS,
                            });
                        }
                        out.push(Equation {
                            shape_id: shapes.last().and_then(|(_, id)| *id),
                            formula: formula.trim().to_owned(),
                            text,
                        });
                        math = None;
                    } else if pushed.pop() == Some(true) && stack.len() > 1 {
                        let node = stack.pop().unwrap_or_default();
                        if let Some(parent) = stack.last_mut() {
                            parent.children.push(node);
                        }
                    }
                }
                if shapes.last().is_some_and(|(at, _)| *at == depth) {
                    shapes.pop();
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

/// Whether `span` of a slide-like owner's XML overlaps an
/// `mc:AlternateContent` whose `mc:Choice` holds an Office Math equation
/// (`a14:m`): the span lies inside such a block, contains one, or crosses
/// its boundary.
///
/// Such a block pairs the equation with a fallback; the shape scene selects
/// the fallback, so a text write there would leave the equation unchanged
/// while the branches diverge. A block may sit around a whole shape or inside
/// its text body, since markup compatibility is allowed anywhere. XML without
/// the markup-compatibility and drawing-2010 namespaces cannot hold such a
/// block and is not scanned.
pub(crate) fn span_overlaps_equation_block(
    xml: &[u8],
    span: &std::ops::Range<usize>,
) -> Result<bool> {
    /// One open element: an `mc:AlternateContent` (its start offset and
    /// whether its Choice holds `a14:m`), an `mc:Choice`, or anything else.
    enum Open {
        Block { start: usize, math: bool },
        Choice,
        Other,
    }
    if memchr::memmem::find(xml, MCE).is_none() || memchr::memmem::find(xml, A14).is_none() {
        return Ok(false);
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    // The caller's scene read already bounded the owner's size, and this
    // pass is linear in it.
    let mut open: Vec<Open> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the slide is too large"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let mce = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == MCE);
        let a14 = matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == A14);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name();
                if a14 && local.as_ref() == b"m" {
                    // Mark every block whose Choice branch encloses this
                    // equation: a Choice's parent is its block.
                    let mut choice_above = false;
                    for frame in open.iter_mut().rev() {
                        match frame {
                            Open::Choice => choice_above = true,
                            Open::Block { math, .. } => {
                                if choice_above {
                                    *math = true;
                                }
                                choice_above = false;
                            },
                            Open::Other => choice_above = false,
                        }
                    }
                }
                if matches!(event, Event::Start(_)) {
                    let frame = if mce && local.as_ref() == b"AlternateContent" {
                        Open::Block { start, math: false }
                    } else if mce && local.as_ref() == b"Choice" {
                        Open::Choice
                    } else {
                        Open::Other
                    };
                    open.push(frame);
                }
            },
            Event::End(_) => {
                let frame = open.pop().ok_or_else(|| invalid("unbalanced slide XML"))?;
                if let Open::Block { start, math: true } = frame {
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_error| invalid("the slide is too large"))?;
                    if start < span.end && span.start < end {
                        return Ok(true);
                    }
                }
            },
            Event::Eof => return Ok(false),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn span_of(xml: &str, needle: &str) -> std::ops::Range<usize> {
        let start = xml.find(needle).unwrap();
        start..start + needle.len()
    }

    /// A shape tree with the given namespace declarations on its root, the
    /// `outside` shape and `body` after it.
    fn tree(declarations: &str, body: &str) -> String {
        format!(
            r#"<p:spTree xmlns:p="{}" xmlns:a="a"{declarations}><p:sp><a:t>o</a:t></p:sp>{body}</p:spTree>"#,
            std::str::from_utf8(PML).unwrap(),
        )
    }

    fn declare(mc: &str, a14: &str) -> String {
        format!(
            r#" xmlns:{mc}="{}" xmlns:{a14}="{}""#,
            std::str::from_utf8(MCE).unwrap(),
            std::str::from_utf8(A14).unwrap(),
        )
    }

    fn overlaps(xml: &str, needle: &str) -> bool {
        span_overlaps_equation_block(xml.as_bytes(), &span_of(xml, needle)).unwrap()
    }

    const OUTSIDE: &str = "<p:sp><a:t>o</a:t></p:sp>";
    const FALLBACK: &str = "<p:sp><a:t>f</a:t></p:sp>";

    #[test]
    fn finds_spans_overlapping_equation_blocks_only() {
        let mc = declare("mc", "a14");
        // A shape in the fallback of a block whose Choice holds `a14:m`.
        let around = tree(
            &mc,
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><a:p><a14:m/></a:p></p:sp></mc:Choice><mc:Fallback>{FALLBACK}</mc:Fallback></mc:AlternateContent>"#
            ),
        );
        assert!(overlaps(&around, FALLBACK));
        assert!(!overlaps(&around, OUTSIDE));

        // A block inside the text body of the selected shape.
        let inner = r#"<p:sp><a:p><mc:AlternateContent><mc:Choice Requires="a14"><a14:m><m:oMath xmlns:m="m"><m:r><m:t>y</m:t></m:r></m:oMath></a14:m></mc:Choice><mc:Fallback><a:r><a:t>y</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p></p:sp>"#;
        let inline = tree(&mc, inner);
        assert!(overlaps(&inline, inner));
        assert!(!overlaps(&inline, OUTSIDE));

        // A block without math is not an equation block.
        let plain = tree(
            &mc,
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><a14:other/></p:sp></mc:Choice><mc:Fallback>{FALLBACK}</mc:Fallback></mc:AlternateContent>"#
            ),
        );
        assert!(!overlaps(&plain, FALLBACK));

        // `a14:m` only in the fallback branch does not make an equation block.
        let fallback_only = r#"<p:sp><a:t>f</a:t><a14:m/></p:sp>"#;
        let fallback_math = tree(
            &mc,
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><a:p/></p:sp></mc:Choice><mc:Fallback>{fallback_only}</mc:Fallback></mc:AlternateContent>"#
            ),
        );
        assert!(!overlaps(&fallback_math, fallback_only));

        // XML without the markup-compatibility namespace is not scanned.
        assert!(!span_overlaps_equation_block(b"<not-xml", &(0..1)).unwrap());
    }

    #[test]
    fn equation_blocks_are_found_by_namespace_not_by_prefix() {
        let block = |mc: &str, a14: &str, local: &str| {
            format!(
                r#"<{mc}:AlternateContent{local}><{mc}:Choice Requires="{a14}"><p:sp><a:p><{a14}:m/></a:p></p:sp></{mc}:Choice><{mc}:Fallback>{FALLBACK}</{mc}:Fallback></{mc}:AlternateContent>"#
            )
        };
        // Renamed prefixes declared on the root.
        let renamed = tree(&declare("zz", "q"), &block("zz", "q", ""));
        assert!(overlaps(&renamed, FALLBACK));
        // Default prefixes declared on the block, not the root.
        let local = tree("", &block("mc", "a14", &declare("mc", "a14")));
        assert!(overlaps(&local, FALLBACK));
        // The familiar prefixes bound to other namespaces are not a block,
        // even though the right namespaces occur elsewhere in the XML.
        let decoy = tree(
            &declare("real", "real14"),
            &block(
                "mc",
                "a14",
                r#" xmlns:mc="urn:other" xmlns:a14="urn:other14""#,
            ),
        );
        assert!(!overlaps(&decoy, FALLBACK));
    }

    #[test]
    fn math_in_comments_or_cdata_is_not_an_equation() {
        let mc = declare("mc", "a14");
        let hidden = tree(
            &mc,
            &format!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><a:p><!--<a14:m/>--><![CDATA[<a14:m/>]]></a:p></p:sp></mc:Choice><mc:Fallback>{FALLBACK}</mc:Fallback></mc:AlternateContent>"#
            ),
        );
        assert!(!overlaps(&hidden, FALLBACK));
    }

    #[test]
    fn reads_the_choice_branch_equation_of_a_shape() {
        let xml = include_bytes!(
            "../../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx"
        );
        let slide = {
            let package = crate::Package::from_bytes(xml).unwrap();
            let presentation = package.presentation().unwrap();
            presentation
                .slide(0)
                .unwrap()
                .unwrap()
                .part()
                .part()
                .blob()
                .to_vec()
        };
        let found = equations(&slide).unwrap();
        assert_eq!(found.len(), 1);
        assert_eq!(found[0].shape_id(), Some(5));
        assert_eq!(found[0].formula(), "𝑎=\\frac{𝑏}{𝑐}");
        assert_eq!(found[0].text(), "𝑎=𝑏𝑐");
    }
}
