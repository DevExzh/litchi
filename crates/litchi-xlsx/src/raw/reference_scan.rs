//! Conservative dependency scan for worksheet deletion.
//!
//! Known local formula and direct-name carriers are reported separately from
//! producer extensions whose semantics cannot be proven. Callers may delete a
//! sheet only when neither class is present in the retained package graph.

use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};
use crate::raw::formula;
use crate::raw::namespace::relationship_attribute_value;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const CHART: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const EXCEL_FORMULA: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const MAX_REFERENCE_DEPTH: usize = 256;
const MAX_CAPTURED_TEXT_BYTES: usize = 8 * 1024 * 1024;

/// The strongest dependency found in one retained XML part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Dependency {
    /// A modeled formula or direct sheet-name carrier names the target.
    Modeled,
    /// An unknown producer field can carry the target and cannot be interpreted.
    Unmodeled,
    /// The target occurs under markup-compatibility choice semantics.
    MarkupCompatibility,
}

/// One dependency and the corresponding index in the checked target slice.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Hit {
    pub(crate) target: usize,
    pub(crate) dependency: Dependency,
}

/// Borrowed logical sheet context. The native ID is admitted only to validate
/// the one modeled workbook carrier that stores it (`activeSheetId`).
#[derive(Debug, Clone, Copy)]
pub(crate) struct Sheet<'a> {
    pub(crate) name: &'a str,
    pub(crate) position: usize,
    pub(crate) native_id: u32,
    pub(crate) catalog: &'a [&'a str],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Mode {
    Formula,
    Suspicious,
    Other,
}

#[derive(Debug)]
struct Frame {
    mode: Mode,
    text: String,
    has_markup: bool,
    alternate: bool,
}

/// Scan one XML part without retaining its payload or allocating on the
/// ordinary no-formula path.
pub(crate) fn scan(content: &[u8], sheets: &[Sheet<'_>]) -> Result<Option<Hit>> {
    let mut reader = NsReader::from_reader(content);
    let mut stack = Vec::<Frame>::new();
    let mut captured_text_bytes = 0usize;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("sheet-dependency XML scan failed: {error}")))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_REFERENCE_DEPTH {
                    return Err(invalid(format!(
                        "sheet-dependency XML nesting exceeds {MAX_REFERENCE_DEPTH} levels"
                    )));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                if let Some(dependency) =
                    scan_attributes(&element, decoder, &resolver, &namespace, alternate, sheets)?
                {
                    return Ok(Some(dependency));
                }
                stack.push(Frame {
                    mode: mode(&element, &namespace)?,
                    text: String::new(),
                    has_markup: false,
                    alternate,
                });
            },
            Event::Empty(element) => {
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                if let Some(dependency) =
                    scan_attributes(&element, decoder, &resolver, &namespace, alternate, sheets)?
                {
                    return Ok(Some(dependency));
                }
            },
            Event::Text(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text.decode().map_err(|error| {
                        invalid(format!("sheet-dependency text is invalid: {error}"))
                    })?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::CData(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text.decode().map_err(|error| {
                        invalid(format!("sheet-dependency CDATA is invalid: {error}"))
                    })?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::GeneralRef(reference) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = decode_xml_reference(&reference)?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::Comment(_) | Event::PI(_) | Event::DocType(_) | Event::Decl(_) => {
                if let Some(parent) = stack.last_mut()
                    && parent.mode != Mode::Other
                {
                    parent.has_markup = true;
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| {
                    invalid("sheet-dependency XML scan found an unmatched closing tag")
                })?;
                captured_text_bytes = captured_text_bytes
                    .checked_sub(frame.text.len())
                    .ok_or_else(|| invalid("sheet-dependency text budget underflow"))?;
                if frame.mode != Mode::Other {
                    for (target, sheet) in sheets.iter().enumerate() {
                        if formula::depends_on_sheet(
                            &frame.text,
                            sheet.name,
                            sheet.position,
                            sheet.catalog,
                        ) {
                            return Ok(Some(Hit {
                                target,
                                dependency: classify(
                                    frame.alternate,
                                    frame.mode == Mode::Suspicious || frame.has_markup,
                                ),
                            }));
                        }
                    }
                }
                if frame.mode != Mode::Other
                    && !sheets.is_empty()
                    && formula::has_dynamic_reference(&frame.text)
                {
                    return Ok(Some(Hit {
                        target: 0,
                        dependency: classify(frame.alternate, true),
                    }));
                }
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("sheet-dependency XML scan ended inside an element"));
    }
    Ok(None)
}

fn scan_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    namespace: &ResolveResult<'_>,
    alternate: bool,
    sheets: &[Sheet<'_>],
) -> Result<Option<Hit>> {
    let name = element.name();
    let local_name = name.local_name();
    let element_local = std::str::from_utf8(local_name.as_ref()).map_err(|error| {
        invalid(format!(
            "sheet-dependency element name is not UTF-8: {error}"
        ))
    })?;
    let spreadsheet = is_namespace(namespace, &[SML, STRICT_SML]);
    let direct_carrier = spreadsheet && matches!(element_local, "worksheetSource" | "dataRef");
    let hyperlink = spreadsheet && element_local == "hyperlink";
    let object_extent = spreadsheet && element_local == "oleSize";
    let custom_view = spreadsheet && element_local == "customWorkbookView";
    let has_relationship = (direct_carrier || hyperlink)
        && relationship_attribute_value(element, b"id", decoder, resolver)?.is_some();

    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let attribute_name = attribute.key.local_name();
        let local = std::str::from_utf8(attribute_name.as_ref()).map_err(|error| {
            invalid(format!(
                "sheet-dependency attribute name is not UTF-8: {error}"
            ))
        })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let direct_name = direct_carrier && local == "sheet";
        if direct_name && has_relationship {
            continue;
        }
        if direct_name
            && let Some(target) = sheets
                .iter()
                .position(|sheet| crate::sheet::equivalent(&value, sheet.name))
        {
            return Ok(Some(Hit {
                target,
                dependency: classify(alternate, false),
            }));
        }
        if custom_view && local == "activeSheetId" {
            let Ok(native_id) = value.parse::<u32>() else {
                return Ok((!sheets.is_empty()).then_some(Hit {
                    target: 0,
                    dependency: classify(alternate, true),
                }));
            };
            if let Some(target) = sheets.iter().position(|sheet| native_id == sheet.native_id) {
                return Ok(Some(Hit {
                    target,
                    dependency: classify(alternate, false),
                }));
            }
        }
        let known_formula = (hyperlink && local == "location" && !has_relationship)
            || (object_extent && local == "ref");
        let suspicious_formula = formula_like(element_local) || formula_like(local);
        if known_formula || suspicious_formula {
            for (target, sheet) in sheets.iter().enumerate() {
                if formula::depends_on_sheet(&value, sheet.name, sheet.position, sheet.catalog) {
                    return Ok(Some(Hit {
                        target,
                        dependency: classify(alternate, !known_formula),
                    }));
                }
            }
        }
        if (known_formula || suspicious_formula)
            && !sheets.is_empty()
            && formula::has_dynamic_reference(&value)
        {
            return Ok(Some(Hit {
                target: 0,
                dependency: classify(alternate, true),
            }));
        }
        if !direct_name
            && sheet_name_like(local)
            && let Some(target) = sheets
                .iter()
                .position(|sheet| crate::sheet::equivalent(&value, sheet.name))
        {
            if spreadsheet && element_local == "sheetProtection" && local == "sheet" {
                continue;
            }
            return Ok(Some(Hit {
                target,
                dependency: classify(alternate, true),
            }));
        }
    }
    Ok(None)
}

fn classify(alternate: bool, unmodeled: bool) -> Dependency {
    if alternate {
        Dependency::MarkupCompatibility
    } else if unmodeled {
        Dependency::Unmodeled
    } else {
        Dependency::Modeled
    }
}

fn mode(element: &BytesStart<'_>, namespace: &ResolveResult<'_>) -> Result<Mode> {
    let name = element.name();
    let local_name = name.local_name();
    let local = std::str::from_utf8(local_name.as_ref()).map_err(|error| {
        invalid(format!(
            "sheet-dependency element name is not UTF-8: {error}"
        ))
    })?;
    if formula_element(local)
        && is_namespace(
            namespace,
            &[SML, STRICT_SML, CHART, STRICT_CHART, EXCEL_FORMULA],
        )
    {
        Ok(Mode::Formula)
    } else if matches!(local, "formula1" | "formula2") && is_namespace(namespace, &[X14]) {
        Ok(Mode::Other)
    } else if formula_like(local) {
        Ok(Mode::Suspicious)
    } else {
        Ok(Mode::Other)
    }
}

fn formula_like(local: &str) -> bool {
    local.eq_ignore_ascii_case("f")
        || contains_ascii_case_insensitive(local, "formula")
        || contains_ascii_case_insensitive(local, "fmla")
        || local.eq_ignore_ascii_case("refersto")
}

fn formula_element(local: &str) -> bool {
    matches!(
        local,
        "f" | "formula"
            | "formula1"
            | "formula2"
            | "definedName"
            | "calculatedColumnFormula"
            | "totalsRowFormula"
    )
}

fn sheet_name_like(local: &str) -> bool {
    local.eq_ignore_ascii_case("sheet")
        || local.eq_ignore_ascii_case("sheetName")
        || local.eq_ignore_ascii_case("worksheet")
}

fn contains_ascii_case_insensitive(value: &str, needle: &str) -> bool {
    value
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle.as_bytes()))
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[&[u8]]) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if expected.contains(value)
    )
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn push_text(stack: &mut [Frame], text: &str, captured_bytes: &mut usize) -> Result<()> {
    let receivers = stack
        .iter()
        .filter(|frame| frame.mode != Mode::Other)
        .count();
    let added = text
        .len()
        .checked_mul(receivers)
        .ok_or_else(|| invalid("sheet-dependency text budget overflow"))?;
    let next = captured_bytes
        .checked_add(added)
        .ok_or_else(|| invalid("sheet-dependency text budget overflow"))?;
    if next > MAX_CAPTURED_TEXT_BYTES {
        return Err(invalid(format!(
            "sheet-dependency text exceeds the {MAX_CAPTURED_TEXT_BYTES}-byte scan budget"
        )));
    }
    *captured_bytes = next;
    for frame in stack.iter_mut().filter(|frame| frame.mode != Mode::Other) {
        frame.text.push_str(text);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn sheet<'a>(catalog: &'a [&'a str]) -> Sheet<'a> {
        Sheet {
            name: "Middle",
            position: 1,
            native_id: 20,
            catalog,
        }
    }

    fn dependency(content: &[u8], catalog: &[&str]) -> Result<Option<Dependency>> {
        Ok(scan(content, &[sheet(catalog)])?.map(|hit| hit.dependency))
    }

    #[test]
    fn classifies_modeled_formula_and_direct_dependencies() {
        let catalog = ["One", "Middle", "Three"];
        let formula =
            format!(r#"<s:worksheet xmlns:s="{S}"><s:f>One:Three!A1</s:f></s:worksheet>"#);
        assert_eq!(
            dependency(formula.as_bytes(), &catalog).expect("formula scan"),
            Some(Dependency::Modeled)
        );
        for direct in [
            format!(
                r#"<s:root xmlns:s="{S}"><s:worksheetSource sheet="Middle" ref="A1"/></s:root>"#
            ),
            format!(r#"<s:workbook xmlns:s="{S}"><s:oleSize ref="Middle!A1:B2"/></s:workbook>"#),
            format!(
                r#"<s:workbook xmlns:s="{S}"><s:customWorkbookViews><s:customWorkbookView activeSheetId="20"/></s:customWorkbookViews></s:workbook>"#
            ),
        ] {
            assert_eq!(
                dependency(direct.as_bytes(), &catalog).expect("direct scan"),
                Some(Dependency::Modeled)
            );
        }
        let targets = [
            sheet(&catalog),
            Sheet {
                name: "Three",
                position: 2,
                native_id: 30,
                catalog: &catalog,
            },
        ];
        assert_eq!(
            scan(
                format!(r#"<s:worksheet xmlns:s="{S}"><s:f>Three!A1</s:f></s:worksheet>"#)
                    .as_bytes(),
                &targets,
            )
            .expect("multiple targets"),
            Some(Hit {
                target: 1,
                dependency: Dependency::Modeled,
            })
        );
    }

    #[test]
    fn ignores_relationship_backed_and_external_workbook_dependencies() {
        let catalog = ["One", "Middle", "Three"];
        let source = format!(
            r#"<s:root xmlns:s="{S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:f>[1]Middle!A1</s:f><s:worksheetSource sheet="Middle" r:id="external"/></s:root>"#
        );
        assert_eq!(
            dependency(source.as_bytes(), &catalog).expect("external scan"),
            None
        );
    }

    #[test]
    fn classifies_unknown_and_compatibility_carriers() {
        let catalog = ["One", "Middle", "Three"];
        assert_eq!(
            dependency(
                br#"<root><futureFormulaCache>Middle!A1</futureFormulaCache></root>"#,
                &catalog,
            )
            .expect("unknown scan"),
            Some(Dependency::Unmodeled)
        );
        assert_eq!(
            dependency(
                br#"<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Fallback><futureFormulaCache>Middle!A1</futureFormulaCache></mc:Fallback></mc:AlternateContent></root>"#,
                &catalog,
            )
            .expect("MCE scan"),
            Some(Dependency::MarkupCompatibility)
        );
        let dynamic =
            format!(r#"<s:worksheet xmlns:s="{S}"><s:f>INDIRECT("A1")</s:f></s:worksheet>"#);
        assert_eq!(
            dependency(dynamic.as_bytes(), &catalog).expect("dynamic scan"),
            Some(Dependency::Unmodeled)
        );
    }
}
