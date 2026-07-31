//! Lossless package-part surgery for local worksheet-name dependencies.
//!
//! Known formula and reference carriers are rewritten together. Unknown
//! formula-like carriers and markup-compatibility alternatives are rejected
//! before publication instead of being guessed at.

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, RenameBlock, Result, invalid};
use crate::raw::formula;
use crate::raw::namespace::relationship_attribute_value;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const SML: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const CHART: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const STRICT_CHART: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const EXCEL_FORMULA: &[u8] = b"http://schemas.microsoft.com/office/excel/2006/main";
const X14: &[u8] = b"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const EXTENDED_PROPERTIES: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const PROPERTY_TYPES: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const MAX_REFERENCE_DEPTH: usize = 256;
const MAX_CAPTURED_TEXT_BYTES: usize = 8 * 1024 * 1024;

/// One simultaneous source-to-final sheet spelling.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Rename<'a> {
    pub(crate) before: &'a str,
    pub(crate) after: &'a str,
}

/// Semantic error context retained across the physical part boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Context<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) part: &'a str,
    pub(crate) sheet_titles: &'a [&'a str],
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Mode {
    Formula,
    Suspicious,
    SheetTitle(usize),
    NamedRangeTitle,
    Other,
}

#[derive(Debug)]
struct Frame {
    mode: Mode,
    inner_start: usize,
    text: String,
    has_markup: bool,
    alternate: bool,
    titles: bool,
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug)]
struct Replacement {
    span: Span,
    bytes: Vec<u8>,
}

#[derive(Debug)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

/// Rewrite one XML part, returning `None` without allocating an output buffer
/// when no recognized local reference changes.
pub(crate) fn rewrite(
    content: &[u8],
    renames: &[Rename<'_>],
    context: Context<'_>,
) -> Result<Option<Vec<u8>>> {
    if renames.is_empty() {
        return Ok(None);
    }
    let formula_renames = renames
        .iter()
        .map(|rename| formula::Rename {
            before: rename.before,
            after: rename.after,
        })
        .collect::<Vec<_>>();
    let mut reader = NsReader::from_reader(content);
    let mut stack = Vec::<Frame>::new();
    let mut replacements = Vec::new();
    let mut saw_titles = false;
    let mut title_entries = 0usize;
    let mut captured_text_bytes = 0usize;

    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("sheet-reference XML scan failed: {error}")))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_REFERENCE_DEPTH {
                    return Err(invalid(format!(
                        "sheet-reference XML nesting exceeds {MAX_REFERENCE_DEPTH} levels"
                    )));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let name = element.name();
                let local_name = name.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|error| {
                    invalid(format!(
                        "sheet-reference element name is not UTF-8: {error}"
                    ))
                })?;
                let starts_titles =
                    local == "TitlesOfParts" && is_namespace(&namespace, &[EXTENDED_PROPERTIES]);
                if local == "TitlesOfParts" && !starts_titles && !context.sheet_titles.is_empty() {
                    return Err(unsupported(context, alternate));
                }
                if starts_titles && std::mem::replace(&mut saw_titles, true) {
                    return Err(unsupported(context, alternate));
                }
                let titles = stack.last().is_some_and(|frame| frame.titles) || starts_titles;
                let title_index =
                    (titles && local == "lpstr" && is_namespace(&namespace, &[PROPERTY_TYPES]))
                        .then(|| {
                            let index = title_entries;
                            title_entries = title_entries.saturating_add(1);
                            index
                        });
                let element_mode = mode(local, title_index, &namespace, context.sheet_titles.len());
                let spreadsheet = is_namespace(&namespace, &[SML, STRICT_SML]);
                attribute_replacement(
                    &element,
                    decoder,
                    &resolver,
                    local,
                    spreadsheet,
                    false,
                    alternate,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                    renames,
                    &formula_renames,
                    context,
                    &mut replacements,
                )?;
                stack.push(Frame {
                    mode: element_mode,
                    inner_start: event_end,
                    text: String::new(),
                    has_markup: false,
                    alternate,
                    titles,
                });
            },
            Event::Empty(element) => {
                if let Some(parent) = stack.last_mut() {
                    parent.has_markup = true;
                }
                let alternate = stack.last().is_some_and(|frame| frame.alternate)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let name = element.name();
                let local_name = name.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|error| {
                    invalid(format!(
                        "sheet-reference element name is not UTF-8: {error}"
                    ))
                })?;
                let starts_titles =
                    local == "TitlesOfParts" && is_namespace(&namespace, &[EXTENDED_PROPERTIES]);
                if local == "TitlesOfParts" && !starts_titles && !context.sheet_titles.is_empty() {
                    return Err(unsupported(context, alternate));
                }
                if starts_titles && std::mem::replace(&mut saw_titles, true) {
                    return Err(unsupported(context, alternate));
                }
                let titles = stack.last().is_some_and(|frame| frame.titles) || starts_titles;
                let title_index =
                    (titles && local == "lpstr" && is_namespace(&namespace, &[PROPERTY_TYPES]))
                        .then(|| {
                            let index = title_entries;
                            title_entries = title_entries.saturating_add(1);
                            index
                        });
                let element_mode = mode(local, title_index, &namespace, context.sheet_titles.len());
                let spreadsheet = is_namespace(&namespace, &[SML, STRICT_SML]);
                attribute_replacement(
                    &element,
                    decoder,
                    &resolver,
                    local,
                    spreadsheet,
                    true,
                    alternate,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                    renames,
                    &formula_renames,
                    context,
                    &mut replacements,
                )?;
                if matches!(element_mode, Mode::SheetTitle(_)) {
                    finish_frame(
                        Frame {
                            mode: element_mode,
                            inner_start: event_end,
                            text: String::new(),
                            has_markup: false,
                            alternate,
                            titles,
                        },
                        event_end,
                        renames,
                        &formula_renames,
                        context,
                        &mut replacements,
                    )?;
                }
            },
            Event::Text(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text
                        .decode()
                        .map_err(|error| invalid(format!("formula text is invalid: {error}")))?;
                    push_text(&mut stack, &text, &mut captured_text_bytes)?;
                }
            },
            Event::CData(text) => {
                if stack.iter().any(|frame| frame.mode != Mode::Other) {
                    let text = text
                        .decode()
                        .map_err(|error| invalid(format!("formula CDATA is invalid: {error}")))?;
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
                    invalid("sheet-reference XML scan found an unmatched closing tag")
                })?;
                captured_text_bytes = captured_text_bytes
                    .checked_sub(frame.text.len())
                    .ok_or_else(|| invalid("sheet-reference text budget underflow"))?;
                finish_frame(
                    frame,
                    event_start,
                    renames,
                    &formula_renames,
                    context,
                    &mut replacements,
                )?;
            },
            Event::Eof => break,
        }
    }
    if !stack.is_empty() {
        return Err(invalid("sheet-reference XML scan ended inside an element"));
    }
    if saw_titles && title_entries < context.sheet_titles.len() {
        return Err(unsupported(context, false));
    }
    if replacements.is_empty() {
        return Ok(None);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping sheet-reference replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.span.end - replacement.span.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("sheet-reference output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|error| invalid(format!("cannot reserve sheet-reference output: {error}")))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

#[allow(clippy::too_many_arguments)]
fn attribute_replacement(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    element_local: &str,
    spreadsheet: bool,
    empty: bool,
    alternate: bool,
    span: Span,
    renames: &[Rename<'_>],
    formula_renames: &[formula::Rename<'_>],
    context: Context<'_>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    let direct_carrier = spreadsheet && matches!(element_local, "worksheetSource" | "dataRef");
    let hyperlink = spreadsheet && element_local == "hyperlink";
    let has_relationship = (direct_carrier || hyperlink)
        && relationship_attribute_value(element, b"id", decoder, resolver)?.is_some();
    let mut changed = Vec::<(Box<str>, String)>::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("sheet-reference attribute is not UTF-8: {error}")))?;
        let attribute_local_name = attribute.key.local_name();
        let local = std::str::from_utf8(attribute_local_name.as_ref()).map_err(|error| {
            invalid(format!(
                "sheet-reference attribute local name is not UTF-8: {error}"
            ))
        })?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        let direct_name_attribute = direct_carrier && local == "sheet";
        let known_direct = direct_name_attribute && !has_relationship;
        let known_formula = hyperlink && local == "location" && !has_relationship;
        if direct_name_attribute && has_relationship {
            continue;
        }
        if known_direct {
            if let Some(after) = direct_name(&value, renames) {
                ensure_direct(alternate, context)?;
                changed.push((name.into(), after.to_owned()));
            }
            continue;
        }
        let suspicious_formula = formula_like(element_local) || formula_like(local);
        if known_formula || suspicious_formula {
            let result = formula::rename_sheets(&value, formula_renames);
            if known_formula {
                if let Some(text) = result.text {
                    ensure_direct(alternate, context)?;
                    changed.push((name.into(), text));
                }
            } else if result.matched {
                return Err(unsupported(context, alternate));
            }
        }
        if !known_direct && sheet_name_like(local) && direct_name(&value, renames).is_some() {
            if spreadsheet && element_local == "sheetProtection" && local == "sheet" {
                continue;
            }
            return Err(unsupported(context, alternate));
        }
    }
    if changed.is_empty() {
        return Ok(());
    }
    replacements.push(Replacement {
        span,
        bytes: write_tag(&tag(element, decoder)?, empty, &changed),
    });
    Ok(())
}

fn finish_frame(
    frame: Frame,
    inner_end: usize,
    renames: &[Rename<'_>],
    formula_renames: &[formula::Rename<'_>],
    context: Context<'_>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    if frame.mode == Mode::Other {
        return Ok(());
    }
    let (text, matched) = match frame.mode {
        Mode::Formula => {
            let result = formula::rename_sheets(&frame.text, formula_renames);
            (result.text, result.matched)
        },
        Mode::Suspicious => {
            let result = formula::rename_sheets(&frame.text, formula_renames);
            if result.matched {
                return Err(unsupported(context, frame.alternate));
            }
            (None, false)
        },
        Mode::SheetTitle(index) => {
            let expected = context.sheet_titles.get(index).ok_or_else(|| {
                invalid("extended-property sheet-title index escaped its checked catalog")
            })?;
            if !crate::sheet::equivalent(&frame.text, expected) {
                return Err(unsupported(context, frame.alternate));
            }
            let direct = direct_name(&frame.text, renames).map(str::to_owned);
            let matched = direct.is_some();
            (direct, matched)
        },
        Mode::NamedRangeTitle => {
            let result = formula::rename_sheets(&frame.text, formula_renames);
            (result.text, result.matched)
        },
        Mode::Other => (None, false),
    };
    if frame.has_markup {
        return if matched {
            Err(unsupported(context, frame.alternate))
        } else {
            Ok(())
        };
    }
    let Some(text) = text else {
        return Ok(());
    };
    ensure_direct(frame.alternate, context)?;
    replacements.push(Replacement {
        span: Span {
            start: frame.inner_start,
            end: inner_end,
        },
        bytes: escape_xml(&text).into_bytes(),
    });
    Ok(())
}

fn direct_name<'a>(value: &str, renames: &'a [Rename<'a>]) -> Option<&'a str> {
    renames
        .iter()
        .find(|rename| crate::sheet::equivalent(value, rename.before))
        .map(|rename| rename.after)
}

fn mode(
    local: &str,
    title_index: Option<usize>,
    namespace: &ResolveResult<'_>,
    sheet_titles: usize,
) -> Mode {
    if let Some(index) = title_index {
        if index < sheet_titles {
            Mode::SheetTitle(index)
        } else {
            Mode::NamedRangeTitle
        }
    } else if formula_element(local)
        && is_namespace(
            namespace,
            &[SML, STRICT_SML, CHART, STRICT_CHART, EXCEL_FORMULA],
        )
    {
        Mode::Formula
    } else if matches!(local, "formula1" | "formula2") && is_namespace(namespace, &[X14]) {
        // x14 uses these as structural wrappers around either an xm:f formula
        // or an x12ac:list literal. The payload's expanded name decides its
        // semantics, so concatenating wrapper children would be incorrect.
        Mode::Other
    } else if formula_like(local) {
        Mode::Suspicious
    } else {
        Mode::Other
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

fn ensure_direct(alternate: bool, context: Context<'_>) -> Result<()> {
    if alternate {
        Err(block(context, RenameBlock::MarkupCompatibility))
    } else {
        Ok(())
    }
}

fn block(context: Context<'_>, reason: RenameBlock) -> Error {
    Error::RenameBlocked {
        sheet: context.sheet.to_owned(),
        position: context.position,
        part: context.part.to_owned(),
        reason,
    }
}

fn unsupported(context: Context<'_>, alternate: bool) -> Error {
    block(
        context,
        if alternate {
            RenameBlock::MarkupCompatibility
        } else {
            RenameBlock::UnmodeledReference
        },
    )
}

fn push_text(stack: &mut [Frame], text: &str, captured_bytes: &mut usize) -> Result<()> {
    let receivers = stack
        .iter()
        .filter(|frame| frame.mode != Mode::Other)
        .count();
    let added = text
        .len()
        .checked_mul(receivers)
        .ok_or_else(|| invalid("sheet-reference text budget overflow"))?;
    let next = captured_bytes
        .checked_add(added)
        .ok_or_else(|| invalid("sheet-reference text budget overflow"))?;
    if next > MAX_CAPTURED_TEXT_BYTES {
        return Err(invalid(format!(
            "sheet-reference text exceeds the {MAX_CAPTURED_TEXT_BYTES}-byte edit budget"
        )));
    }
    *captured_bytes = next;
    for frame in stack.iter_mut().filter(|frame| frame.mode != Mode::Other) {
        frame.text.push_str(text);
    }
    Ok(())
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| {
            invalid(format!(
                "sheet-reference element name is not UTF-8: {error}"
            ))
        })?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("sheet-reference attribute is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

fn write_tag(tag: &Tag, empty: bool, changed: &[(Box<str>, String)]) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        let value = changed
            .iter()
            .find(|(name, _)| name.as_ref() == attribute.name.as_ref())
            .map_or(attribute.value.as_ref(), |(_, value)| value.as_str());
        output.extend_from_slice(escape_xml(value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    output
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("sheet-reference XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn context<'a>(part: &'a str) -> Context<'a> {
        Context {
            sheet: "Data",
            position: 0,
            part,
            sheet_titles: &["Data"],
        }
    }

    #[test]
    fn rewrites_formula_hyperlink_pivot_and_title_dependencies() {
        let source = format!(
            r#"<root xmlns:s="{S}" xmlns:ep="{}" xmlns:vt="{}"><s:f>Data!A1</s:f><s:formula1>'Data'!Named</s:formula1><s:hyperlink location="Data!$A$1" keep="x"/><s:worksheetSource sheet="data" ref="A1"/><s:dataRef sheet="Data" ref="A1"/><ep:TitlesOfParts><vt:vector><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data</vt:lpstr><vt:lpstr>Data!Print_Area</vt:lpstr></vt:vector></ep:TitlesOfParts></root>"#,
            std::str::from_utf8(EXTENDED_PROPERTIES).expect("namespace"),
            std::str::from_utf8(PROPERTY_TYPES).expect("namespace")
        );
        let changed = rewrite(
            source.as_bytes(),
            &[Rename {
                before: "Data",
                after: "Input 2026",
            }],
            context("/xl/test.xml"),
        )
        .expect("rewrite")
        .expect("changed output");
        let changed = std::str::from_utf8(&changed).expect("UTF-8");
        assert!(changed.contains("<s:f>&apos;Input 2026&apos;!A1</s:f>"));
        assert!(changed.contains("<s:formula1>&apos;Input 2026&apos;!Named</s:formula1>"));
        assert!(changed.contains("location=\"&apos;Input 2026&apos;!$A$1\""));
        assert_eq!(changed.matches("sheet=\"Input 2026\"").count(), 2);
        assert!(changed.contains("<vt:lpstr>Input 2026</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>Data</vt:lpstr>"));
        assert!(changed.contains("<vt:lpstr>&apos;Input 2026&apos;!Print_Area</vt:lpstr>"));
    }

    #[test]
    fn leaves_external_and_inert_values_exact() {
        let source = format!(
            r#"<root xmlns:s="{S}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><s:f>[1]Data!A1</s:f><value>Data!A1</value><s:f>Other!A1</s:f><s:worksheetSource sheet="Data" ref="A1" r:id="external-book"/><s:dataRef sheet="Data" ref="A1" r:id="external-data"/><s:hyperlink ref="A1" location="Data!A1" r:id="external-link"/></root>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &[Rename {
                before: "Data",
                after: "Input",
            }],
            context("/xl/test.xml"),
        )
        .expect("scan");
        assert!(changed.is_none());
    }

    #[test]
    fn blocks_unmodeled_formula_like_and_compatibility_references() {
        let rename = [Rename {
            before: "Data",
            after: "Input",
        }];
        let unknown = rewrite(
            br#"<root><futureFormulaCache>Data!A1</futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        );
        assert!(matches!(
            unknown,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let alternate = rewrite(
            br#"<root xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Fallback><f>Data!A1</f></mc:Fallback></mc:AlternateContent></root>"#,
            &rename,
            context("/xl/alternate.xml"),
        );
        assert!(matches!(
            alternate,
            Err(Error::RenameBlocked {
                reason: RenameBlock::MarkupCompatibility,
                ..
            })
        ));

        let nested = rewrite(
            br#"<root><futureFormulaCache><value>Data!A1</value></futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        );
        assert!(matches!(
            nested,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let unrelated_nested = rewrite(
            br#"<root><futureFormulaCache><value>42</value></futureFormulaCache></root>"#,
            &rename,
            context("/xl/future.xml"),
        )
        .expect("unrelated nested extension");
        assert!(unrelated_nested.is_none());

        let malformed_known = rewrite(
            format!(r#"<s:f xmlns:s="{S}"><future/>Data!A1</s:f>"#).as_bytes(),
            &rename,
            context("/xl/malformed.xml"),
        );
        assert!(matches!(
            malformed_known,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));
    }

    #[test]
    fn classifies_namespaces_wrappers_and_vml_conservatively() {
        let rename = [Rename {
            before: "Data",
            after: "Input",
        }];
        let spoofed = rewrite(
            br#"<root xmlns:x="urn:not-a-formula"><x:f>Data!A1</x:f></root>"#,
            &rename,
            context("/xl/spoofed.xml"),
        );
        assert!(matches!(
            spoofed,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));
        let spoofed_titles = rewrite(
            br#"<TitlesOfParts><lpstr>Data</lpstr></TitlesOfParts>"#,
            &rename,
            context("/docProps/app.xml"),
        );
        assert!(matches!(
            spoofed_titles,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let vml = rewrite(
            br#"<xml xmlns:x="urn:schemas-microsoft-com:office:excel"><x:FmlaRange>Data!A1:A3</x:FmlaRange></xml>"#,
            &rename,
            context("/xl/drawings/vmlDrawing1.vml"),
        );
        assert!(matches!(
            vml,
            Err(Error::RenameBlocked {
                reason: RenameBlock::UnmodeledReference,
                ..
            })
        ));

        let x14 = format!(
            r#"<root xmlns:x14="{}" xmlns:xm="{}"><x14:formula1><xm:f>Data!A1</xm:f></x14:formula1></root>"#,
            std::str::from_utf8(X14).expect("namespace"),
            std::str::from_utf8(EXCEL_FORMULA).expect("namespace")
        );
        let changed = rewrite(
            x14.as_bytes(),
            &rename,
            context("/xl/worksheets/sheet1.xml"),
        )
        .expect("x14 scan")
        .expect("xm formula changed");
        assert!(
            std::str::from_utf8(&changed)
                .expect("UTF-8")
                .contains("<xm:f>Input!A1</xm:f>")
        );

        let numeric = [Rename {
            before: "1",
            after: "One",
        }];
        assert!(
            rewrite(
                format!(r#"<sheetProtection xmlns="{S}" sheet="1"/>"#).as_bytes(),
                &numeric,
                context("/xl/worksheets/sheet1.xml"),
            )
            .expect("boolean sheet attribute")
            .is_none()
        );

        let deep = format!(
            "{}Data!A1{}",
            "<futureFormula>".repeat(MAX_REFERENCE_DEPTH + 1),
            "</futureFormula>".repeat(MAX_REFERENCE_DEPTH + 1)
        );
        assert!(matches!(
            rewrite(deep.as_bytes(), &rename, context("/xl/deep.xml")),
            Err(Error::Invalid(message)) if message.contains("nesting")
        ));
    }
}
