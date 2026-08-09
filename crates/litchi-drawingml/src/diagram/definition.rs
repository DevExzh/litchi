//! Metadata parsing for `DrawingML` diagram definition parts.
//!
//! The layout (`dgm:layoutDef`), quick-style (`dgm:styleDef`), and colors
//! (`dgm:colorsDef`) parts share a common header shape: a `uniqueId` attribute
//! plus `dgm:title`/`dgm:desc`/`dgm:catLst` children. This module reads that
//! header as inert metadata; the layout algorithm and style bodies are not
//! interpreted.

use crate::diagram::{DGM_NAMESPACE, DGM_NAMESPACE_STRICT};
use crate::{Error, Result};
use litchi_ooxml_common::mce::process_str;
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
    pub fn parse_layout(xml: &str) -> Result<Self> {
        Self::parse(xml, "layoutDef")
    }

    /// Parse a diagram quick-style part (`dgm:styleDef`).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_quick_style(xml: &str) -> Result<Self> {
        Self::parse(xml, "styleDef")
    }

    /// Parse a diagram colors part (`dgm:colorsDef`).
    /// # Errors
    ///
    /// Returns an error when input violates DrawingML constraints, exceeds a configured
    /// bound, or an underlying XML, MCE, I/O, or formatting operation fails.
    pub fn parse_colors(xml: &str) -> Result<Self> {
        Self::parse(xml, "colorsDef")
    }

    fn parse(xml: &str, root_name: &str) -> Result<Self> {
        if xml.len() > MAX_DEFINITION_XML {
            return Err(limit("definition XML bytes"));
        }
        let processed = process_str(xml)?;
        if processed.len() > MAX_DEFINITION_XML {
            return Err(limit("processed definition XML bytes"));
        }
        let mut reader = NsReader::from_reader(processed.as_bytes());
        reader.config_mut().trim_text(true);

        let mut scan = DefinitionScan {
            definition: DiagramDefinition::default(),
            root_name,
            root_seen: false,
            nodes: 0,
        };
        let mut buffer = Vec::new();
        let mut depth = 0usize;

        loop {
            match reader.read_event_into(&mut buffer).map_err(xml_error)? {
                Event::Start(element) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("definition XML depth"))?;
                    scan.visit(&reader, &element, depth)?;
                },
                Event::Empty(element) => {
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("definition XML depth"))?;
                    scan.visit(&reader, &element, child_depth)?;
                },
                Event::End(_) => {
                    if depth == 0 {
                        return Err(invalid("unexpected definition closing element"));
                    }
                    depth -= 1;
                },
                Event::DocType(_) => return Err(invalid("DTDs are rejected")),
                Event::CData(_) => return Err(invalid("CDATA is rejected")),
                Event::Eof => break,
                Event::Text(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
            buffer.clear();
        }
        if !scan.root_seen || depth != 0 {
            return Err(invalid("missing or unterminated definition root"));
        }
        Ok(scan.definition)
    }
}

struct DefinitionScan<'a> {
    definition: DiagramDefinition,
    root_name: &'a str,
    root_seen: bool,
    nodes: usize,
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
            .ok_or_else(|| limit("definition XML node count"))?;
        if self.nodes > MAX_NODES || depth > MAX_DEPTH {
            return Err(limit("definition XML structure"));
        }
        let namespace = element_namespace(reader, element)?;
        let local = local_name(element)?;
        if !self.root_seen {
            if !is_dgm(&namespace) || local != self.root_name {
                return Err(invalid("invalid diagram definition root or namespace"));
            }
            self.root_seen = true;
            self.definition.unique_id = attribute(element, "uniqueId")?;
            return Ok(());
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
    for item in element.attributes().with_checks(true) {
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

fn limit(label: &str) -> Error {
    invalid(format!("diagram {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;

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
        let definition = DiagramDefinition::parse_layout(xml).unwrap();
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
        let definition = DiagramDefinition::parse_quick_style(style).unwrap();
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
        let definition = DiagramDefinition::parse_colors(colors).unwrap();
        assert_eq!(
            definition.unique_id.as_deref(),
            Some("urn:test/colors/accent1_2")
        );
        assert_eq!(definition.categories[0].category_type, "accent1");
    }

    #[test]
    fn rejects_mismatched_root() {
        let xml = "<dgm:colorsDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>";
        assert!(DiagramDefinition::parse_layout(xml).is_err());
        assert!(DiagramDefinition::parse_colors(xml).is_ok());
    }
}
