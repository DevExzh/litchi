//! Worksheet-level calculation properties (`CT_SheetCalcPr`).
//!
//! The worksheet `sheetCalcPr` element complements the workbook `calcPr`: its
//! `fullCalcOnLoad` flag forces a full recalculation of this sheet when the
//! workbook is opened. This module parses and serializes the element; it never
//! triggers a recalculation itself.

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};
use litchi_ooxml_common::mce::process_str;

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_DEPTH: usize = 256;

/// Namespace form used when serializing a `sheetCalcPr` fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetSheetCalculationPropertiesConformance {
    Transitional,
    Strict,
}

impl WorksheetSheetCalculationPropertiesConformance {
    fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }
}

/// Immutable worksheet `sheetCalcPr` settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct WorksheetSheetCalculationProperties {
    full_calc_on_load: bool,
}

impl WorksheetSheetCalculationProperties {
    pub fn new(full_calc_on_load: bool) -> Self {
        Self { full_calc_on_load }
    }

    /// Whether this sheet is fully recalculated the next time the workbook loads.
    pub fn full_calc_on_load(&self) -> bool {
        self.full_calc_on_load
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    SheetCalculationProperties,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Unbound,
    Main,
    Other,
}

/// Parses the direct worksheet `sheetCalcPr` child after applying shared MCE processing.
pub fn parse_worksheet_sheet_calculation_properties(
    xml: &[u8],
) -> Result<Option<WorksheetSheetCalculationProperties>> {
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    let processed = process_str(source)?;
    let mut reader = NsReader::from_reader(processed.as_bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut scopes = Vec::new();
    let mut properties: Option<WorksheetSheetCalculationProperties> = None;
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        let namespace = namespace_kind(resolved)?;
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if scopes.is_empty() && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                let next_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML depth overflow"))?;
                if next_depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                let scope = begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut properties,
                )?;
                if scopes.is_empty() {
                    root_seen = true;
                }
                depth = next_depth;
                scopes.push(scope);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if scopes.is_empty() {
                    if root_seen {
                        return Err(invalid("worksheet XML contains multiple roots"));
                    }
                    begin_element(&reader, &element, namespace, None, &mut properties)?;
                    root_seen = true;
                    root_closed = true;
                } else {
                    begin_element(
                        &reader,
                        &element,
                        namespace,
                        scopes.last().copied(),
                        &mut properties,
                    )?;
                }
            },
            Event::End(element) => {
                let scope = scopes
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet end element"))?;
                match scope {
                    Scope::Worksheet => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"worksheet"
                        {
                            return Err(invalid("mismatched worksheet end element"));
                        }
                        root_closed = true;
                    },
                    Scope::SheetCalculationProperties => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"sheetCalcPr"
                        {
                            return Err(invalid("mismatched sheetCalcPr end element"));
                        }
                    },
                    Scope::Other => {},
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("worksheet XML depth underflow"))?;
            },
            Event::Text(text) => {
                let whitespace = text.as_ref().iter().all(u8::is_ascii_whitespace);
                if scopes.is_empty() {
                    if !whitespace {
                        return Err(invalid("worksheet XML text is outside root"));
                    }
                } else if matches!(scopes.last(), Some(Scope::SheetCalculationProperties))
                    && !whitespace
                {
                    return Err(invalid("sheetCalcPr must be a leaf element"));
                }
            },
            Event::CData(_) => {
                if scopes.is_empty() {
                    return Err(invalid("worksheet XML CDATA is outside root"));
                }
                if matches!(scopes.last(), Some(Scope::SheetCalculationProperties)) {
                    return Err(invalid("sheetCalcPr must be a leaf element"));
                }
            },
            Event::GeneralRef(reference) => {
                if scopes.is_empty() {
                    return Err(invalid("worksheet XML entity is outside root"));
                }
                if matches!(scopes.last(), Some(Scope::SheetCalculationProperties)) {
                    return Err(invalid("sheetCalcPr must be a leaf element"));
                }
                let name = reference
                    .decode()
                    .map_err(|error| invalid(format!("invalid worksheet XML entity: {error}")))?;
                if reference
                    .resolve_char_ref()
                    .map_err(|error| invalid(format!("invalid worksheet XML entity: {error}")))?
                    .is_none()
                    && !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot")
                {
                    return Err(invalid("custom XML entities are rejected"));
                }
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || depth != 0 || !scopes.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    Ok(properties)
}

fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    properties: &mut Option<WorksheetSheetCalculationProperties>,
) -> Result<Scope> {
    let local = element.local_name();
    let local = local.as_ref();
    let main = namespace == NamespaceKind::Main;
    match parent {
        None => {
            if !main || local != b"worksheet" {
                return Err(invalid("expected SpreadsheetML worksheet root"));
            }
            Ok(Scope::Worksheet)
        },
        Some(Scope::Worksheet) => {
            if local != b"sheetCalcPr" {
                return Ok(Scope::Other);
            }
            if !main {
                return Err(invalid("spoofed sheetCalcPr element namespace"));
            }
            if properties.is_some() {
                return Err(invalid("duplicate worksheet sheetCalcPr element"));
            }
            *properties = Some(parse_sheet_calc_pr_attributes(reader, element)?);
            Ok(Scope::SheetCalculationProperties)
        },
        Some(Scope::SheetCalculationProperties) => {
            Err(invalid("sheetCalcPr must be a leaf element"))
        },
        Some(Scope::Other) => Ok(Scope::Other),
    }
}

fn parse_sheet_calc_pr_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<WorksheetSheetCalculationProperties> {
    let mut full_calc_on_load = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid sheetCalcPr attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound
            || local.as_ref() != b"fullCalcOnLoad"
        {
            return Err(invalid("unknown or spoofed sheetCalcPr attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid sheetCalcPr attribute value: {error}")))?;
        if full_calc_on_load.is_some() {
            return Err(invalid("duplicate sheetCalcPr fullCalcOnLoad attribute"));
        }
        full_calc_on_load = Some(match text.as_ref() {
            "true" | "1" => true,
            "false" | "0" => false,
            _ => return Err(invalid("sheetCalcPr fullCalcOnLoad must be an XML boolean")),
        });
    }
    Ok(WorksheetSheetCalculationProperties {
        full_calc_on_load: full_calc_on_load.unwrap_or(false),
    })
}

/// Serializes one canonical, namespace-complete `sheetCalcPr` fragment.
pub fn write_worksheet_sheet_calculation_properties(
    value: &WorksheetSheetCalculationProperties,
    conformance: WorksheetSheetCalculationPropertiesConformance,
) -> String {
    let mut xml = String::with_capacity(64 + conformance.main_namespace().len());
    xml.push_str("<sheetCalcPr xmlns=\"");
    xml.push_str(conformance.main_namespace());
    xml.push('"');
    if value.full_calc_on_load {
        xml.push_str(" fullCalcOnLoad=\"1\"");
    }
    xml.push_str("/>");
    xml
}

fn namespace_kind(result: ResolveResult<'_>) -> Result<NamespaceKind> {
    match result {
        ResolveResult::Unbound => Ok(NamespaceKind::Unbound),
        ResolveResult::Bound(namespace) if is_main_namespace(namespace.as_ref()) => {
            Ok(NamespaceKind::Main)
        },
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix {}",
            String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn is_main_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_MAIN.as_bytes() || namespace == STRICT_MAIN.as_bytes()
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetSheetCalculationProperties>> {
        parse_worksheet_sheet_calculation_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_full_calc_on_load_and_default() {
        let empty = format!(r#"<worksheet xmlns="{NS}"/>"#);
        assert!(
            parse_worksheet_sheet_calculation_properties(empty.as_bytes())
                .unwrap()
                .is_none()
        );
        assert!(
            parse(r#"<sheetCalcPr fullCalcOnLoad="1"/>"#)
                .unwrap()
                .unwrap()
                .full_calc_on_load()
        );
        assert!(
            !parse(r#"<sheetCalcPr fullCalcOnLoad="false"/>"#)
                .unwrap()
                .unwrap()
                .full_calc_on_load()
        );
        assert!(
            !parse("<sheetCalcPr/>")
                .unwrap()
                .unwrap()
                .full_calc_on_load()
        );
        assert!(parse("<sheetData/>").unwrap().is_none());
    }

    #[test]
    fn supports_strict_namespace() {
        let xml = concat!(
            r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">"#,
            r#"<sheetCalcPr fullCalcOnLoad="true"/></worksheet>"#,
        );
        assert!(
            parse_worksheet_sheet_calculation_properties(xml.as_bytes())
                .unwrap()
                .unwrap()
                .full_calc_on_load()
        );
    }

    #[test]
    fn rejects_bad_structure_and_attributes() {
        for child in [
            r#"<sheetCalcPr fullCalcOnLoad="yes"/>"#,
            r#"<sheetCalcPr mystery="1"/>"#,
            r#"<sheetCalcPr fullCalcOnLoad="1" fullCalcOnLoad="0"/>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
        assert!(parse("<sheetCalcPr/><sheetCalcPr/>").is_err());
    }

    #[test]
    fn rejects_leaf_content_and_malformed_document_state() {
        for child in [
            "<sheetCalcPr><extension/></sheetCalcPr>",
            "<sheetCalcPr>unexpected</sheetCalcPr>",
            "<sheetCalcPr><![CDATA[unexpected]]></sheetCalcPr>",
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }

        let mismatched_end = format!(r#"<worksheet xmlns="{NS}"><sheetCalcPr/></wrong>"#,);
        assert!(parse_worksheet_sheet_calculation_properties(mismatched_end.as_bytes()).is_err());

        let multiple_roots = format!(
            r#"<worksheet xmlns="{NS}"><sheetData/></worksheet><worksheet xmlns="{NS}"><sheetData/></worksheet>"#,
        );
        assert!(parse_worksheet_sheet_calculation_properties(multiple_roots.as_bytes()).is_err());
    }

    #[test]
    fn rejects_excessive_worksheet_nesting() {
        let mut xml = format!(r#"<worksheet xmlns="{NS}">"#);
        xml.push_str(&"<extension>".repeat(MAX_DEPTH));
        xml.push_str(&"</extension>".repeat(MAX_DEPTH));
        xml.push_str("</worksheet>");

        assert!(parse_worksheet_sheet_calculation_properties(xml.as_bytes()).is_err());
    }

    #[test]
    fn write_round_trips_through_the_reader() {
        for expected in [
            WorksheetSheetCalculationProperties::new(true),
            WorksheetSheetCalculationProperties::new(false),
        ] {
            for conformance in [
                WorksheetSheetCalculationPropertiesConformance::Transitional,
                WorksheetSheetCalculationPropertiesConformance::Strict,
            ] {
                let fragment = write_worksheet_sheet_calculation_properties(&expected, conformance);
                let document = format!(r#"<worksheet xmlns="{NS}">{fragment}</worksheet>"#);
                let parsed = parse_worksheet_sheet_calculation_properties(document.as_bytes())
                    .unwrap()
                    .unwrap();
                assert_eq!(parsed, expected);
            }
        }
    }

    #[test]
    fn reads_fixture_sheet_calc_pr() {
        let package = OpcPackage::from_bytes(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsx/FormatChoiceTests.xlsx"
        )))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        let value = parse_worksheet_sheet_calculation_properties(part.blob())
            .unwrap()
            .unwrap();
        assert!(value.full_calc_on_load());
    }
}
