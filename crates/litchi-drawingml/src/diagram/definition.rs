//! Metadata parsing for `DrawingML` diagram definition parts.
//!
//! The layout (`dgm:layoutDef`), quick-style (`dgm:styleDef`), and colors
//! (`dgm:colorsDef`) parts share a common header shape: a `uniqueId` attribute
//! plus `dgm:title`/`dgm:desc`/`dgm:catLst` children. This module reads that
//! header as inert metadata; the layout algorithm and style bodies are not
//! interpreted.

use crate::diagram::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT, Limits};
use crate::{Error, LimitKind, Result};
use litchi_ooxml_common::mce::process_str;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_DEFINITION_XML: usize = 4 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 64;

/// A category entry (`dgm:cat`) advertised by a diagram definition part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramCategory {
    /// Category type (e.g. `list`, `process`, `cycle`, `accent1`, `simple`).
    pub category_type: String,
    /// Category priority (`pri`), when it is a valid non-negative integer.
    pub priority: Option<u32>,
}

/// Inert header metadata of a diagram layout, quick-style, or colors part.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DiagramDefinition {
    /// Definition identifier (`uniqueId`, e.g. an `officeart` layout URI).
    pub unique_id: Option<String>,
    /// Title (`dgm:title@val`), when non-empty.
    pub title: Option<String>,
    /// Description (`dgm:desc@val`), when non-empty.
    pub description: Option<String>,
    /// Advertised categories (`dgm:catLst/dgm:cat`).
    pub categories: Vec<DiagramCategory>,
}

impl DiagramDefinition {
    /// Parse a diagram layout part (`dgm:layoutDef`).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_layout(xml: &str, limits: &Limits) -> Result<Self> {
        Self::parse(xml, "layoutDef", limits)
    }

    /// Parse a diagram quick-style part (`dgm:styleDef`).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_quick_style(xml: &str, limits: &Limits) -> Result<Self> {
        Self::parse(xml, "styleDef", limits)
    }

    /// Parse a diagram colors part (`dgm:colorsDef`).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_colors(xml: &str, limits: &Limits) -> Result<Self> {
        Self::parse(xml, "colorsDef", limits)
    }

    fn parse(xml: &str, root_name: &str, limits: &Limits) -> Result<Self> {
        limits.validate()?;
        let max_input = limits.input_bytes_value().min(MAX_DEFINITION_XML);
        if xml.len() > max_input {
            return Err(limit(LimitKind::XmlBytes, xml.len(), max_input));
        }
        let processed = process_str(xml)?;
        if processed.len() > max_input {
            return Err(limit(LimitKind::XmlBytes, processed.len(), max_input));
        }
        let mut reader = NsReader::from_reader(processed.as_bytes());
        reader.config_mut().trim_text(true);

        let mut scan = DefinitionScan {
            definition: DiagramDefinition::default(),
            root_name,
            phase: DefinitionPhase::Prolog,
            nodes: 0,
            max_nodes: limits.xml_nodes_value().min(MAX_NODES),
            max_depth: limits.xml_depth_value().min(MAX_DEPTH),
            max_categories: limits.nodes_value(),
        };
        let mut buffer = Vec::new();
        let mut depth = 0usize;

        loop {
            match reader.read_event_into(&mut buffer).map_err(xml_error)? {
                Event::Start(element) => {
                    if scan.phase == DefinitionPhase::Closed {
                        return Err(invalid("trailing markup follows the diagram definition"));
                    }
                    if depth == 0 && scan.phase != DefinitionPhase::Prolog {
                        return Err(invalid("the diagram definition has multiple roots"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("definition XML depth overflows"))?;
                    scan.visit(&reader, &element, depth)?;
                },
                Event::Empty(element) => {
                    if scan.phase == DefinitionPhase::Closed {
                        return Err(invalid("trailing markup follows the diagram definition"));
                    }
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("definition XML depth overflows"))?;
                    if depth == 0 && scan.phase != DefinitionPhase::Prolog {
                        return Err(invalid("the diagram definition has multiple roots"));
                    }
                    scan.visit(&reader, &element, child_depth)?;
                    if depth == 0 {
                        scan.phase = DefinitionPhase::Closed;
                    }
                },
                Event::End(_) => {
                    if depth == 0 || scan.phase == DefinitionPhase::Prolog {
                        return Err(invalid("unexpected definition closing element"));
                    }
                    depth -= 1;
                    if depth == 0 {
                        scan.phase = DefinitionPhase::Closed;
                    }
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::CData(_) => return Err(invalid("CDATA is rejected")),
                Event::Eof => break,
                Event::Text(text) => {
                    if scan.phase != DefinitionPhase::Open
                        && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())
                    {
                        return Err(invalid(
                            "non-whitespace text is outside the diagram definition",
                        ));
                    }
                },
                Event::GeneralRef(_) if scan.phase != DefinitionPhase::Open => {
                    return Err(invalid(
                        "entity reference is outside the diagram definition",
                    ));
                },
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
        if scan.phase != DefinitionPhase::Closed || depth != 0 {
            return Err(invalid("missing or unterminated definition root"));
        }
        Ok(scan.definition)
    }
}

struct DefinitionScan<'a> {
    definition: DiagramDefinition,
    root_name: &'a str,
    phase: DefinitionPhase,
    nodes: usize,
    max_nodes: usize,
    max_depth: usize,
    max_categories: usize,
}

/// Where the definition header reader is relative to the part's one root.
/// Nesting inside the root is not part of it: the depth counter holds that.
///
/// ```text
/// Prolog --root start tag-------------------------> Open
/// Prolog --root empty-element tag-----------------> Closed (through Open, inside one event)
/// Open   --nested start, empty or end tag---------> Open
/// Open   --end tag back to depth 0----------------> Closed
/// any    --comment, declaration or PI-------------> unchanged (accepted anywhere)
/// Prolog | Closed --whitespace--------------------> unchanged
/// Open   --text, entity reference-----------------> Open
/// ```
///
/// Refused: a root that is not the expected `dgm` element; any element after
/// `Closed`; a second element at depth 0; an end tag in `Prolog`; text other
/// than whitespace or an entity reference outside `Open`; a DOCTYPE or CDATA
/// anywhere; and the end of input anywhere but `Closed` at depth 0.
#[derive(Clone, Copy, PartialEq, Eq)]
enum DefinitionPhase {
    /// No element has been read.
    Prolog,
    /// The root start tag was read and its end tag was not.
    Open,
    /// The root element is complete.
    Closed,
}

impl DefinitionScan<'_> {
    fn visit(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        depth: usize,
    ) -> Result<()> {
        self.nodes = self
            .nodes
            .checked_add(1)
            .ok_or_else(|| invalid("definition XML node count overflows"))?;
        if depth > self.max_depth {
            return Err(limit(LimitKind::XmlDepth, depth, self.max_depth));
        }
        if self.nodes > self.max_nodes {
            return Err(limit(LimitKind::XmlNodes, self.nodes, self.max_nodes));
        }
        let namespace = element_namespace(reader, element)?;
        let local = local_name(element)?;
        if self.phase == DefinitionPhase::Prolog {
            if !is_dgm(&namespace) || local != self.root_name {
                return Err(invalid("invalid diagram definition root or namespace"));
            }
            self.phase = DefinitionPhase::Open;
            self.definition.unique_id = attribute(element, "uniqueId")?;
            return Ok(());
        }
        if is_dgm(&namespace) && local == self.root_name {
            return Err(invalid("the diagram definition has a nested root"));
        }
        if !is_dgm(&namespace) {
            return Ok(());
        }
        match local.as_str() {
            "title" => {
                if let Some(value) = attribute(element, "val")?
                    && !value.is_empty()
                {
                    self.definition.title.get_or_insert(value);
                }
            },
            "desc" => {
                if let Some(value) = attribute(element, "val")?
                    && !value.is_empty()
                {
                    self.definition.description.get_or_insert(value);
                }
            },
            "cat" => {
                if let Some(category_type) = attribute(element, "type")? {
                    if self.definition.categories.len() >= self.max_categories {
                        return Err(limit(
                            LimitKind::DiagramNodes,
                            self.definition.categories.len() + 1,
                            self.max_categories,
                        ));
                    }
                    let priority = attribute(element, "pri")?.and_then(|value| value.parse().ok());
                    self.definition.categories.push(DiagramCategory {
                        category_type,
                        priority,
                    });
                }
            },
            _ => {},
        }
        Ok(())
    }
}

fn attribute(element: &BytesStart<'_>, name: &str) -> Result<Option<String>> {
    for item in element.checked_attributes() {
        let item = item.map_err(xml_error)?;
        if item.key.local_name().as_ref() == name.as_bytes() {
            let value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
            return Ok(Some(
                quick_xml::escape::unescape(value)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn element_namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    match reader.resolver().resolve_element(element.name()).0 {
        ResolveResult::Bound(Namespace(namespace)) => Ok(std::str::from_utf8(namespace)
            .map_err(xml_error)?
            .to_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn local_name(element: &BytesStart<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn is_dgm(namespace: &str) -> bool {
    matches!(namespace, DGM_NAMESPACE | DGM_NAMESPACE_STRICT)
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_layout(xml: &str) -> Result<DiagramDefinition> {
        DiagramDefinition::parse_layout(xml, &Limits::default())
    }

    fn parse_quick_style(xml: &str) -> Result<DiagramDefinition> {
        DiagramDefinition::parse_quick_style(xml, &Limits::default())
    }

    fn parse_colors(xml: &str) -> Result<DiagramDefinition> {
        DiagramDefinition::parse_colors(xml, &Limits::default())
    }

    #[test]
    fn parses_layout_definition() {
        let xml = concat!(
            "<?xml version=\"1.0\"?>",
            "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" ",
            "uniqueId=\"urn:microsoft.com/office/officeart/2005/8/layout/cycle2\">",
            "<dgm:title val=\"Cycle\"/><dgm:desc val=\"\"/>",
            "<dgm:catLst><dgm:cat type=\"cycle\" pri=\"1000\"/><dgm:cat type=\"convert\" pri=\"10000\"/></dgm:catLst>",
            "<dgm:layoutNode name=\"diagram\"/>",
            "</dgm:layoutDef>"
        );
        let definition = parse_layout(xml).unwrap();
        assert_eq!(
            definition.unique_id.as_deref(),
            Some("urn:microsoft.com/office/officeart/2005/8/layout/cycle2")
        );
        assert_eq!(definition.title.as_deref(), Some("Cycle"));
        assert_eq!(definition.description, None);
        assert_eq!(
            definition.categories,
            vec![
                DiagramCategory {
                    category_type: "cycle".to_string(),
                    priority: Some(1000),
                },
                DiagramCategory {
                    category_type: "convert".to_string(),
                    priority: Some(10000),
                },
            ]
        );
    }

    #[test]
    fn parses_strict_quick_style_and_colors_definitions() {
        let style = concat!(
            "<dgm:styleDef xmlns:dgm=\"http://purl.oclc.org/ooxml/drawingml/diagram\" ",
            "uniqueId=\"urn:test/quickstyle/simple1\">",
            "<dgm:catLst><dgm:cat type=\"simple\" pri=\"10100\"/></dgm:catLst>",
            "</dgm:styleDef>"
        );
        let definition = parse_quick_style(style).unwrap();
        assert_eq!(
            definition.unique_id.as_deref(),
            Some("urn:test/quickstyle/simple1")
        );
        assert_eq!(definition.categories.len(), 1);

        let colors = concat!(
            "<dgm:colorsDef xmlns:dgm=\"http://purl.oclc.org/ooxml/drawingml/diagram\" ",
            "uniqueId=\"urn:test/colors/accent1_2\">",
            "<dgm:catLst><dgm:cat type=\"accent1\" pri=\"11200\"/></dgm:catLst>",
            "</dgm:colorsDef>"
        );
        let definition = parse_colors(colors).unwrap();
        assert_eq!(
            definition.unique_id.as_deref(),
            Some("urn:test/colors/accent1_2")
        );
        assert_eq!(definition.categories[0].category_type, "accent1");
    }

    #[test]
    fn rejects_mismatched_root() {
        let xml = "<dgm:colorsDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>";
        assert!(parse_layout(xml).is_err());
        assert!(parse_colors(xml).is_ok());
    }

    #[test]
    fn rejects_duplicate_or_trailing_layout_roots() {
        let duplicate = concat!(
            "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>",
            "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>"
        );
        assert!(matches!(parse_layout(duplicate), Err(Error::Invalid(_))));

        let trailing = concat!(
            "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>",
            "trailing"
        );
        assert!(matches!(parse_layout(trailing), Err(Error::Invalid(_))));
    }

    #[test]
    fn rejects_unterminated_or_nested_layout_root() {
        let unterminated = "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">";
        assert!(matches!(parse_layout(unterminated), Err(Error::Invalid(_))));

        let nested = concat!(
            "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
            "<dgm:layoutDef/>",
            "</dgm:layoutDef>"
        );
        assert!(matches!(parse_layout(nested), Err(Error::Invalid(_))));
    }
}

#[cfg(test)]
mod envelope_tests {
    //! What the diagram definition reader accepts before, around and after its root, and the exact error reported otherwise. Each
    //! expected line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
            ("reference before root", format!("&amp;{d}")),
            ("reference after root", format!("{d}&amp;")),
            ("CDATA after root", format!("{d}<![CDATA[x]]>")),
            ("end tag before root", format!("</foreign>{d}")),
            (
                "declaration inside root",
                format!("{open}{DECLARATION}{body}{close}"),
            ),
            (
                "text, CDATA and reference inside root",
                format!("{open}{body}text<![CDATA[x]]>&amp;{close}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&str) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(&xml)))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The definition header reader, through the colors part.
    #[test]
    fn colors_definition_envelope() {
        let root = Root {
            open: r#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:test/colors">"#,
            body: r#"<dgm:title val="T"/><dgm:catLst><dgm:cat type="accent1" pri="1"/></dgm:catLst>"#,
            close: r"</dgm:colorsDef>",
            empty: r#"<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:test/colors"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"missing or unterminated definition root\")",
                "whitespace only: Invalid(\"missing or unterminated definition root\")",
                "declaration only: Invalid(\"missing or unterminated definition root\")",
                "empty-element root: Ok(Some(\"urn:test/colors\") None 0)",
                "start and end root: Ok(Some(\"urn:test/colors\") None 0)",
                "root then whitespace: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "root then comment: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "root then PI: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "root then text: Invalid(\"non-whitespace text is outside the diagram definition\")",
                "root then second root: Invalid(\"trailing markup follows the diagram definition\")",
                "root then foreign start: Invalid(\"trailing markup follows the diagram definition\")",
                "root then foreign empty: Invalid(\"trailing markup follows the diagram definition\")",
                "declaration after root: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "two declarations: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "comment before declaration: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "text before root: Invalid(\"non-whitespace text is outside the diagram definition\")",
                "unterminated root: Invalid(\"missing or unterminated definition root\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"DTDs are rejected\")",
                "top-level CDATA: Invalid(\"CDATA is rejected\")",
                "declaration then root: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "whitespace before declaration: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "PI before declaration: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "comment then root: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "PI then root: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "root then text then second root: Invalid(\"non-whitespace text is outside the diagram definition\")",
                "empty-element root then root: Invalid(\"trailing markup follows the diagram definition\")",
                "declaration after empty-element root: Ok(Some(\"urn:test/colors\") None 0)",
                "reference before root: Invalid(\"entity reference is outside the diagram definition\")",
                "reference after root: Invalid(\"entity reference is outside the diagram definition\")",
                "CDATA after root: Invalid(\"CDATA is rejected\")",
                "end tag before root: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "declaration inside root: Ok(Some(\"urn:test/colors\") Some(\"T\") 1)",
                "text, CDATA and reference inside root: Invalid(\"CDATA is rejected\")",
            ],
            |xml| {
                outcome(
                    super::DiagramDefinition::parse_colors(xml, &super::Limits::default()),
                    |value| {
                        format!(
                            "{:?} {:?} {}",
                            value.unique_id,
                            value.title,
                            value.categories.len()
                        )
                    },
                )
            },
        );
    }
}
