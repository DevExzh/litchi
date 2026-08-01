//! Table styles part reader (`/ppt/tableStyles.xml`, `a:tblStyleLst`).
//!
//! The table styles part declares the default table style (`def`) and any
//! custom table style definitions (`a:tblStyle`) a presentation carries.
//! Tables reference one of these styles through `a:tblPr/a:tableStyleId`.
//!
//! This module reports only the stored style inventory: which styles exist
//! and which conditional part styles (header row, banding, corner cells, ...)
//! each one defines. Cell style payloads (fills, borders, text styles) are
//! retained as presence metadata only and are never resolved or rendered.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::xml::{is_drawingml_name, unqualified_attribute_value};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 64;
const MAX_XML_ATTRIBUTES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;
const MAX_STYLES: usize = 4_096;

const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

/// Relationship type from the presentation part to the table styles part
/// (strict dialect).
const STRICT_TABLE_STYLES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles";

/// A conditional part of a table that a table style may define.
///
/// Element names follow ECMA-376 Part 1, 21.1.2.1.9 (`a:tblStyle`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum TableStylePartKind {
    /// `a:wholeTbl` — style applied to the entire table.
    WholeTable,
    /// `a:band1H` — style applied to odd horizontal (row) bands.
    OddRowBand,
    /// `a:band2H` — style applied to even horizontal (row) bands.
    EvenRowBand,
    /// `a:band1V` — style applied to odd vertical (column) bands.
    OddColumnBand,
    /// `a:band2V` — style applied to even vertical (column) bands.
    EvenColumnBand,
    /// `a:firstCol` — style applied to the first column.
    FirstColumn,
    /// `a:lastCol` — style applied to the last column.
    LastColumn,
    /// `a:firstRow` — style applied to the first (header) row.
    FirstRow,
    /// `a:lastRow` — style applied to the last (totals) row.
    LastRow,
    /// `a:seCell` — style applied to the bottom-right cell.
    SouthEastCell,
    /// `a:swCell` — style applied to the bottom-left cell.
    SouthWestCell,
    /// `a:neCell` — style applied to the top-right cell.
    NorthEastCell,
    /// `a:nwCell` — style applied to the top-left cell.
    NorthWestCell,
}

impl TableStylePartKind {
    /// The DrawingML element name declaring this part style.
    pub fn xml_name(self) -> &'static str {
        match self {
            TableStylePartKind::WholeTable => "wholeTbl",
            TableStylePartKind::OddRowBand => "band1H",
            TableStylePartKind::EvenRowBand => "band2H",
            TableStylePartKind::OddColumnBand => "band1V",
            TableStylePartKind::EvenColumnBand => "band2V",
            TableStylePartKind::FirstColumn => "firstCol",
            TableStylePartKind::LastColumn => "lastCol",
            TableStylePartKind::FirstRow => "firstRow",
            TableStylePartKind::LastRow => "lastRow",
            TableStylePartKind::SouthEastCell => "seCell",
            TableStylePartKind::SouthWestCell => "swCell",
            TableStylePartKind::NorthEastCell => "neCell",
            TableStylePartKind::NorthWestCell => "nwCell",
        }
    }

    fn from_xml_name(name: &[u8]) -> Option<Self> {
        match name {
            b"wholeTbl" => Some(TableStylePartKind::WholeTable),
            b"band1H" => Some(TableStylePartKind::OddRowBand),
            b"band2H" => Some(TableStylePartKind::EvenRowBand),
            b"band1V" => Some(TableStylePartKind::OddColumnBand),
            b"band2V" => Some(TableStylePartKind::EvenColumnBand),
            b"firstCol" => Some(TableStylePartKind::FirstColumn),
            b"lastCol" => Some(TableStylePartKind::LastColumn),
            b"firstRow" => Some(TableStylePartKind::FirstRow),
            b"lastRow" => Some(TableStylePartKind::LastRow),
            b"seCell" => Some(TableStylePartKind::SouthEastCell),
            b"swCell" => Some(TableStylePartKind::SouthWestCell),
            b"neCell" => Some(TableStylePartKind::NorthEastCell),
            b"nwCell" => Some(TableStylePartKind::NorthWestCell),
            _ => None,
        }
    }
}

/// One `a:tblStyle` definition: its identity and the part styles it defines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableStyleDefinition {
    /// The style GUID (`styleId`), as stored; tables reference this value
    /// through `a:tblPr/a:tableStyleId`.
    pub style_id: Option<String>,
    /// The display name of the style (`styleName`), as stored.
    pub style_name: Option<String>,
    /// The conditional part styles this definition declares, sorted by
    /// [`TableStylePartKind`] order without duplicates.
    pub parts: Vec<TableStylePartKind>,
}

impl TableStyleDefinition {
    /// Whether this definition declares the given part style.
    pub fn has(&self, part: TableStylePartKind) -> bool {
        self.parts.contains(&part)
    }
}

/// The parsed table styles part (`a:tblStyleLst`).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableStyleList {
    /// The default table style GUID (`def`), when declared.
    pub default_style_id: Option<String>,
    /// The custom table style definitions declared by the presentation.
    pub styles: Vec<TableStyleDefinition>,
}

impl TableStyleList {
    /// Look up a style definition by its `styleId` GUID.
    pub fn find(&self, style_id: &str) -> Option<&TableStyleDefinition> {
        self.styles
            .iter()
            .find(|style| style.style_id.as_deref() == Some(style_id))
    }

    /// Parse a table styles part from its XML bytes.
    pub fn parse(xml_bytes: &[u8]) -> Result<Self> {
        if xml_bytes.len() > MAX_XML_BYTES {
            return Err(limit("table styles XML bytes"));
        }
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut list = TableStyleList::default();
        let mut nodes = 0usize;
        let mut depth = 0usize;
        let mut saw_root = false;
        let mut closed_root = false;
        let mut open_style_depth: Option<usize> = None;

        loop {
            let decoder = reader.decoder();
            let event = reader
                .read_event()
                .map_err(|error| OoxmlError::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);

            match event {
                Event::Start(element) => {
                    increment_nodes(&mut nodes)?;
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("table styles XML depth"))?;
                    if depth > MAX_XML_DEPTH {
                        return Err(limit("table styles XML depth"));
                    }
                    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
                        return Err(limit("table styles XML attribute count"));
                    }
                    if depth == 1 {
                        if saw_root
                            || !is_drawingml_name(&namespace, element.name(), b"tblStyleLst")
                        {
                            return Err(invalid(
                                "table styles part must have one DrawingML tblStyleLst root",
                            ));
                        }
                        saw_root = true;
                        list.default_style_id = bounded_optional(
                            unqualified_attribute_value(&element, b"def", decoder)?,
                            "default table style ID",
                        )?;
                    } else if depth == 2
                        && is_drawingml_name(&namespace, element.name(), b"tblStyle")
                    {
                        if list.styles.len() >= MAX_STYLES {
                            return Err(limit("table style count"));
                        }
                        list.styles.push(parse_style_definition(&element, decoder)?);
                        open_style_depth = Some(depth);
                    } else if open_style_depth == Some(depth - 1) {
                        record_part_style(&mut list, &namespace, element.name())?;
                    }
                },
                Event::Empty(element) => {
                    increment_nodes(&mut nodes)?;
                    let child_depth = depth
                        .checked_add(1)
                        .ok_or_else(|| limit("table styles XML depth"))?;
                    if child_depth > MAX_XML_DEPTH {
                        return Err(limit("table styles XML depth"));
                    }
                    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
                        return Err(limit("table styles XML attribute count"));
                    }
                    if child_depth == 1 {
                        if saw_root
                            || !is_drawingml_name(&namespace, element.name(), b"tblStyleLst")
                        {
                            return Err(invalid(
                                "table styles part must have one DrawingML tblStyleLst root",
                            ));
                        }
                        saw_root = true;
                        closed_root = true;
                        list.default_style_id = bounded_optional(
                            unqualified_attribute_value(&element, b"def", decoder)?,
                            "default table style ID",
                        )?;
                    } else if child_depth == 2
                        && is_drawingml_name(&namespace, element.name(), b"tblStyle")
                    {
                        if list.styles.len() >= MAX_STYLES {
                            return Err(limit("table style count"));
                        }
                        list.styles.push(parse_style_definition(&element, decoder)?);
                    } else if open_style_depth == Some(child_depth - 1) {
                        record_part_style(&mut list, &namespace, element.name())?;
                    }
                },
                Event::End(element) => {
                    if depth == 0 {
                        return Err(invalid("invalid table styles XML nesting"));
                    }
                    if depth == 1 {
                        if !is_drawingml_name(&namespace, element.name(), b"tblStyleLst") {
                            return Err(invalid(
                                "table styles part must close with a DrawingML tblStyleLst element",
                            ));
                        }
                        closed_root = true;
                    }
                    if open_style_depth == Some(depth)
                        && is_drawingml_name(&namespace, element.name(), b"tblStyle")
                    {
                        let style = list
                            .styles
                            .last_mut()
                            .ok_or_else(|| invalid("missing open table style"))?;
                        style.parts.sort_unstable();
                        open_style_depth = None;
                    }
                    depth -= 1;
                },
                Event::DocType(_) => {
                    return Err(invalid("table styles part must not contain a DTD"));
                },
                Event::Eof => {
                    if !saw_root || !closed_root || depth != 0 || open_style_depth.is_some() {
                        return Err(invalid(
                            "unterminated or missing DrawingML tblStyleLst root",
                        ));
                    }
                    break;
                },
                _ => {},
            }
        }

        Ok(list)
    }
}

fn parse_style_definition(
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<TableStyleDefinition> {
    Ok(TableStyleDefinition {
        style_id: bounded_optional(
            unqualified_attribute_value(element, b"styleId", decoder)?,
            "table style ID",
        )?,
        style_name: bounded_optional(
            unqualified_attribute_value(element, b"styleName", decoder)?,
            "table style name",
        )?,
        parts: Vec::new(),
    })
}

fn record_part_style(
    list: &mut TableStyleList,
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
) -> Result<()> {
    let in_drawingml = matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    );
    let part = in_drawingml
        .then(|| TableStylePartKind::from_xml_name(name.local_name().as_ref()))
        .flatten();
    let Some(part) = part else {
        return Err(invalid("unexpected element in table style definition"));
    };
    let style = list
        .styles
        .last_mut()
        .ok_or_else(|| invalid("missing open table style"))?;
    if style.parts.contains(&part) {
        return Err(invalid(
            "table style definition declares a part style twice",
        ));
    }
    style.parts.push(part);
    Ok(())
}

/// Load the presentation's table styles part, if the package declares one.
pub(crate) fn load_table_styles(package: &OpcPackage) -> Result<Option<TableStyleList>> {
    let presentation = package.main_document_part()?;
    let mut found = presentation.rels().iter().filter(|rel| {
        matches!(
            rel.reltype(),
            rt::TABLE_STYLES | STRICT_TABLE_STYLES_RELATIONSHIP
        )
    });
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid(
            "presentation has multiple table-styles relationships",
        ));
    }
    if rel.is_external() {
        return Err(invalid("table-styles relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != ct::PML_TABLE_STYLES {
        return Err(OoxmlError::InvalidContentType {
            expected: ct::PML_TABLE_STYLES.to_string(),
            got: part.content_type().to_string(),
        });
    }
    Ok(Some(TableStyleList::parse(part.blob())?))
}

fn bounded_optional(value: Option<String>, what: &str) -> Result<Option<String>> {
    if let Some(value) = &value
        && value.len() > MAX_ATTRIBUTE_BYTES
    {
        return Err(limit(what));
    }
    Ok(value)
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("table styles XML node count"))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("table styles XML node count"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    #[test]
    fn parses_default_only_list() {
        let xml = format!(
            r#"<a:tblStyleLst xmlns:a="{A_NS}" def="{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}"/>"#
        );
        let list = TableStyleList::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            list.default_style_id.as_deref(),
            Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        );
        assert!(list.styles.is_empty());
    }

    #[test]
    fn parses_style_definitions_with_part_styles() {
        let xml = format!(
            r#"<a:tblStyleLst xmlns:a="{A_NS}" def="{{A}}"><a:tblStyle styleId="{{A}}" styleName="Branded"><a:wholeTbl><a:tcTxStyle/></a:wholeTbl><a:band1H/><a:firstRow/><a:nwCell/></a:tblStyle><a:tblStyle styleId="{{B}}"><a:lastCol/></a:tblStyle></a:tblStyleLst>"#
        );
        let list = TableStyleList::parse(xml.as_bytes()).unwrap();
        assert_eq!(list.styles.len(), 2);

        let branded = list.find("{A}").unwrap();
        assert_eq!(branded.style_name.as_deref(), Some("Branded"));
        let mut expected = vec![
            TableStylePartKind::OddRowBand,
            TableStylePartKind::FirstRow,
            TableStylePartKind::NorthWestCell,
            TableStylePartKind::WholeTable,
        ];
        expected.sort_unstable();
        assert_eq!(branded.parts, expected);
        assert!(branded.has(TableStylePartKind::WholeTable));
        assert!(branded.has(TableStylePartKind::FirstRow));
        assert!(!branded.has(TableStylePartKind::LastRow));

        let second = list.find("{B}").unwrap();
        assert_eq!(second.style_name, None);
        assert!(second.has(TableStylePartKind::LastColumn));
    }

    #[test]
    fn accepts_the_strict_drawingml_dialect() {
        let xml = r#"<a:tblStyleLst xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><a:tblStyle styleId="{S}"><a:seCell/></a:tblStyle></a:tblStyleLst>"#;
        let list = TableStyleList::parse(xml.as_bytes()).unwrap();
        assert!(list.default_style_id.is_none());
        assert!(
            list.find("{S}")
                .unwrap()
                .has(TableStylePartKind::SouthEastCell)
        );
    }

    #[test]
    fn rejects_duplicates_unknown_children_and_foreign_roots() {
        let xml = format!(
            r#"<a:tblStyleLst xmlns:a="{A_NS}"><a:tblStyle><a:firstRow/><a:firstRow/></a:tblStyle></a:tblStyleLst>"#
        );
        assert!(TableStyleList::parse(xml.as_bytes()).is_err());
        let xml = format!(
            r#"<a:tblStyleLst xmlns:a="{A_NS}"><a:tblStyle><a:wholeTblx/></a:tblStyle></a:tblStyleLst>"#
        );
        assert!(TableStyleList::parse(xml.as_bytes()).is_err());
        let xml = r#"<a:tblStyleLst xmlns:a="urn:wrong"/>"#;
        assert!(TableStyleList::parse(xml.as_bytes()).is_err());
        let xml = format!(r#"<!DOCTYPE a:tblStyleLst><a:tblStyleLst xmlns:a="{A_NS}"/>"#);
        assert!(TableStyleList::parse(xml.as_bytes()).is_err());
    }

    #[test]
    fn part_kind_xml_names_round_trip() {
        for part in [
            TableStylePartKind::WholeTable,
            TableStylePartKind::OddRowBand,
            TableStylePartKind::EvenRowBand,
            TableStylePartKind::OddColumnBand,
            TableStylePartKind::EvenColumnBand,
            TableStylePartKind::FirstColumn,
            TableStylePartKind::LastColumn,
            TableStylePartKind::FirstRow,
            TableStylePartKind::LastRow,
            TableStylePartKind::SouthEastCell,
            TableStylePartKind::SouthWestCell,
            TableStylePartKind::NorthEastCell,
            TableStylePartKind::NorthWestCell,
        ] {
            assert_eq!(
                TableStylePartKind::from_xml_name(part.xml_name().as_bytes()),
                Some(part)
            );
        }
        assert_eq!(TableStylePartKind::from_xml_name(b"tbl"), None);
    }
}
