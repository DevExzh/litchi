use super::model::{PlaceholderSpec, SlideLayoutKind};
use crate::{Error, Result};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fmt::Write as FmtWrite;

pub(super) const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(super) const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const STRICT_SLIDE_MASTER_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
pub(super) const STRICT_SLIDE_LAYOUT_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";

/// Shape ID 1 is reserved for the group-shape root of every shape tree.
pub(super) const FIRST_SHAPE_ID: u32 = 2;
/// Bounded-input ceiling for every part this module parses or patches.
pub(super) const MAX_PART_XML_BYTES: usize = 8 * 1024 * 1024;
/// Bounded-input ceiling for XML node counts while scanning.
pub(super) const MAX_SCAN_NODES: usize = 100_000;
/// Bounded-input ceiling for XML nesting depth while scanning.
pub(super) const MAX_SCAN_DEPTH: usize = 128;
/// Bounded ceiling for authored layout names.
pub(super) const MAX_NAME_CHARS: usize = 256;
/// Bounded ceiling for placeholder shapes authored in a single operation.
pub(super) const MAX_PLACEHOLDERS_PER_OPERATION: usize = 64;
/// Indentation step between the nine paragraph levels, in EMUs.
pub(super) const LEVEL_MARGIN_STEP_EMU: u32 = 457_200;
/// Default body font size for generated text-style levels, in hundredths of a point.
pub(super) const LEVEL_FONT_SIZE_HUNDREDTHS: u32 = 1800;

pub(super) const XML_DECL: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
pub(super) const SP_TREE_HEADER: &str = "<p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr>";
pub(super) const COLOR_MAP: &str = "<p:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>";

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

// ============================================================================

/// Serialize a new slide master part with default text styles.
pub(super) fn master_xml() -> String {
    let mut xml = String::with_capacity(8192);
    xml.push_str(XML_DECL);
    xml.push_str("<p:sldMaster xmlns:a=\"");
    xml.push_str(A_NS);
    xml.push_str("\" xmlns:r=\"");
    xml.push_str(R_NS);
    xml.push_str("\" xmlns:p=\"");
    xml.push_str(P_NS);
    xml.push_str("\"><p:cSld>");
    xml.push_str(SP_TREE_HEADER);
    xml.push_str("</p:spTree></p:cSld>");
    xml.push_str(COLOR_MAP);
    xml.push_str("<p:sldLayoutIdLst/>");
    xml.push_str("<p:txStyles><p:titleStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:titleStyle><p:bodyStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:bodyStyle><p:otherStyle>");
    push_text_style_levels(&mut xml);
    xml.push_str("</p:otherStyle></p:txStyles></p:sldMaster>");
    xml
}

/// Write the nine paragraph levels shared by all generated text styles.
pub(super) fn push_text_style_levels(xml: &mut String) {
    for level in 1..=9u32 {
        let margin = (level - 1) * LEVEL_MARGIN_STEP_EMU;
        let _result = write!(
            xml,
            "<a:lvl{level}pPr marL=\"{margin}\" algn=\"l\" defTabSz=\"457200\" rtl=\"0\" eaLnBrk=\"1\" latinLnBrk=\"0\" hangingPunct=\"1\"><a:defRPr sz=\"{LEVEL_FONT_SIZE_HUNDREDTHS}\" kern=\"1200\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/><a:cs typeface=\"+mn-cs\"/></a:defRPr></a:lvl{level}pPr>"
        );
    }
}

/// Serialize a new slide layout part.
pub(super) fn layout_xml(
    kind: SlideLayoutKind,
    name: &str,
    placeholders: &[PlaceholderSpec],
) -> Result<String> {
    let mut xml = String::with_capacity(2048);
    xml.push_str(XML_DECL);
    let _result = write!(
        xml,
        "<p:sldLayout xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\" type=\"{}\" matchingName=\"{}\"><p:cSld name=\"{}\">",
        kind.as_str(),
        escape_xml(name),
        escape_xml(name)
    );
    xml.push_str(SP_TREE_HEADER);
    for (offset, spec) in placeholders.iter().enumerate() {
        let shape_id = FIRST_SHAPE_ID + offset as u32;
        xml.push_str(&placeholder_shape_xml(shape_id, spec, false));
    }
    xml.push_str(
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
    );
    if xml.len() > MAX_PART_XML_BYTES {
        return Err(invalid(
            "generated slide layout exceeds the part size limit",
        ));
    }
    Ok(xml)
}

/// Serialize one placeholder shape.
///
/// When `declare_namespaces` is set the shape carries its own `xmlns`
/// declarations so it can be patched into a part with unknown prefix
/// bindings.
pub(super) fn placeholder_shape_xml(
    shape_id: u32,
    spec: &PlaceholderSpec,
    declare_namespaces: bool,
) -> String {
    let name = spec
        .name
        .clone()
        .unwrap_or_else(|| format!("{} Placeholder {shape_id}", spec.kind.label()));
    let mut xml = String::with_capacity(512);
    xml.push_str("<p:sp");
    if declare_namespaces {
        let _result = write!(xml, " xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"");
    }
    let _result = write!(
        xml,
        "><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"{}\"",
        escape_xml(&name),
        spec.kind.as_str()
    );
    if let Some(index) = spec.index {
        let _result = write!(xml, " idx=\"{index}\"");
    }
    xml.push_str("/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>");
    if let Some(text) = &spec.text {
        let _result = write!(xml, "<a:r><a:t>{}</a:t></a:r>", escape_xml(text));
    }
    xml.push_str("<a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>");
    xml
}

// ============================================================================
// Bounded XML scanning and patching
// ============================================================================

pub(super) const SPTREE_DEPTH: usize = 3;
/// Depth of `p:sp` shapes inside the shape tree.
pub(super) const SHAPE_DEPTH: usize = 4;

/// Byte span of an XML element.
#[derive(Debug, Clone, Copy)]
pub(super) struct ElementSpan {
    /// Offset of the `<` that opens the element.
    pub(super) start: usize,
    /// Offset one past the `>` that closes the element.
    pub(super) end: usize,
    /// Offset of the `</` that opens the closing tag (equals `start` for empty elements).
    pub(super) close_start: usize,
    /// Whether the element uses the self-closing form.
    pub(super) empty: bool,
}

/// Where a missing ID list should be created.
pub(super) enum IdListAnchor {
    /// `p:sldMasterIdLst` heads the `CT_Presentation` sequence.
    AfterRootStart,
    /// `p:sldLayoutIdLst` follows `p:clrMap` in the `CT_SlideMaster` sequence.
    AfterElement(&'static str),
}

pub(super) fn check_size(xml: &[u8]) -> Result<()> {
    if xml.len() > MAX_PART_XML_BYTES {
        return Err(invalid("part XML exceeds 8 MiB"));
    }
    Ok(())
}

pub(super) fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

/// Find the first element with `target` as local name at exactly `depth`.
pub(super) fn scan_element_span(
    xml: &[u8],
    target: &str,
    depth: usize,
) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut stack: Vec<(usize, String)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                let local =
                    String::from_utf8_lossy(local_name(element.name().as_ref())).into_owned();
                stack.push((before, local));
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if stack.len() + 1 == depth
                    && local_name(element.name().as_ref()) == target.as_bytes()
                {
                    return Ok(Some(ElementSpan {
                        start: before,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: true,
                    }));
                }
            },
            Ok(Event::End(element)) => {
                let (start, local) = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element in part XML"))?;
                if stack.len() + 1 == depth && local == target {
                    return Ok(Some(ElementSpan {
                        start,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: false,
                    }));
                }
                if local_name(element.name().as_ref()) != local.as_bytes() {
                    return Err(invalid("mismatched closing element in part XML"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated part XML"));
    }
    Ok(None)
}

/// Insert `entry` into the ID list element `list_local`, creating the list at
/// the schema-correct position when it is missing.
pub(super) fn insert_id_list_entry(
    xml: &[u8],
    list_local: &str,
    entry: &str,
    anchor: IdListAnchor,
) -> Result<Vec<u8>> {
    if let Some(span) = scan_element_span(xml, list_local, 2)? {
        if span.empty {
            let wrapped = format!(
                "<p:{list_local} xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\">{entry}</p:{list_local}>"
            );
            return replace_span(xml, &span, wrapped.as_bytes());
        }
        return insert_bytes(xml, span.close_start, entry.as_bytes());
    }
    let wrapped =
        format!("<p:{list_local} xmlns:p=\"{P_NS}\" xmlns:r=\"{R_NS}\">{entry}</p:{list_local}>");
    let offset = match anchor {
        IdListAnchor::AfterRootStart => root_start_end(xml)?,
        IdListAnchor::AfterElement(anchor_local) => {
            let span = scan_element_span(xml, anchor_local, 2)?.ok_or_else(|| {
                invalid(format!("part XML is missing its '{anchor_local}' anchor"))
            })?;
            span.end
        },
    };
    insert_bytes(xml, offset, wrapped.as_bytes())
}

/// Offset one past the root element's start tag.
pub(super) fn root_start_end(xml: &[u8]) -> Result<usize> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(_) | Event::Empty(_)) => {
                return Ok(reader.buffer_position() as usize);
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => return Err(invalid("part XML has no root element")),
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
}

/// Remove the ID-list entry whose `r:id` matches `relationship_id`.
pub(super) fn remove_id_list_entry(
    xml: &[u8],
    entry_local: &str,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == entry_local.as_bytes()
                    && element_relationship_id(&element)?.as_deref() == Some(relationship_id)
                {
                    let span = ElementSpan {
                        start: before,
                        end: reader.buffer_position() as usize,
                        close_start: before,
                        empty: true,
                    };
                    return replace_span(xml, &span, b"");
                }
            },
            Ok(Event::Start(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == entry_local.as_bytes()
                    && element_relationship_id(&element)?.as_deref() == Some(relationship_id)
                {
                    // Consume events up to the matching closing tag so entries
                    // with extension children are removed whole.
                    let mut depth = 1usize;
                    loop {
                        match reader.read_event() {
                            Ok(Event::Start(_)) => depth += 1,
                            Ok(Event::End(_)) => {
                                depth -= 1;
                                if depth == 0 {
                                    let span = ElementSpan {
                                        start: before,
                                        end: reader.buffer_position() as usize,
                                        close_start: before,
                                        empty: false,
                                    };
                                    return replace_span(xml, &span, b"");
                                }
                            },
                            Ok(Event::Eof) => {
                                return Err(invalid("unterminated ID-list entry"));
                            },
                            Err(error) => return Err(Error::Xml(error.to_string())),
                            _ => {},
                        }
                    }
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    Err(invalid(format!(
        "ID list has no entry for relationship '{relationship_id}'"
    )))
}

/// Read the relationship-namespace `id` attribute of an element.
pub(super) fn element_relationship_id(
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<String>> {
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?;
        if name.rsplit_once(':').map(|(_, local)| local) == Some("id") && name.contains(':') {
            let value = std::str::from_utf8(attribute.value.as_ref())
                .map_err(|error| Error::Xml(error.to_string()))?;
            return Ok(Some(value.to_owned()));
        }
    }
    Ok(None)
}

/// Find the direct `p:sp` child of the shape tree whose `p:ph` matches
/// `kind` and `index`.
pub(super) fn find_placeholder_span(
    xml: &[u8],
    kind: &str,
    index: u32,
) -> Result<Option<ElementSpan>> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut depth = 0usize;
    let mut shape_start = None;
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                depth += 1;
                if nodes > MAX_SCAN_NODES || depth > MAX_SCAN_DEPTH {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                let local = local_name(element.name().as_ref()).to_vec();
                if depth == SHAPE_DEPTH && local == b"sp" {
                    shape_start = Some(before);
                } else if local == b"ph"
                    && shape_start.is_some()
                    && placeholder_matches(&element, kind, index)?
                {
                    let start = shape_start.ok_or_else(|| invalid("missing placeholder shape"))?;
                    return Ok(Some(ElementSpan {
                        start,
                        end: shape_end(xml, start)?,
                        close_start: start,
                        empty: false,
                    }));
                }
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == b"ph"
                    && shape_start.is_some()
                    && placeholder_matches(&element, kind, index)?
                {
                    let start = shape_start.ok_or_else(|| invalid("missing placeholder shape"))?;
                    return Ok(Some(ElementSpan {
                        start,
                        end: shape_end(xml, start)?,
                        close_start: start,
                        empty: false,
                    }));
                }
            },
            Ok(Event::End(_)) => {
                if depth == SHAPE_DEPTH {
                    shape_start = None;
                }
                if depth == 0 {
                    return Err(invalid("unexpected closing element in part XML"));
                }
                depth -= 1;
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("unterminated part XML"));
    }
    Ok(None)
}

/// Whether a `p:ph` element matches the requested type and index.
pub(super) fn placeholder_matches(
    element: &quick_xml::events::BytesStart<'_>,
    kind: &str,
    index: u32,
) -> Result<bool> {
    let mut ph_type = None;
    let mut ph_index = 0u32;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        match attribute.key.as_ref() {
            b"type" => {
                ph_type = Some(
                    std::str::from_utf8(attribute.value.as_ref())
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .to_owned(),
                );
            },
            b"idx" => {
                let value = std::str::from_utf8(attribute.value.as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?;
                ph_index = value
                    .parse::<u32>()
                    .map_err(|_err| invalid(format!("invalid placeholder index '{value}'")))?;
            },
            _ => {},
        }
    }
    // ECMA defaults: type "obj", idx 0.
    Ok(ph_type.as_deref().unwrap_or("obj") == kind && ph_index == index)
}

/// Compute the end offset of the `p:sp` element starting at `start`.
pub(super) fn shape_end(xml: &[u8], start: usize) -> Result<usize> {
    let mut reader = Reader::from_reader(&xml[start..]);
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(_)) => depth += 1,
            Ok(Event::Empty(_)) if depth == 0 => {
                return Ok(start + reader.buffer_position() as usize);
            },
            Ok(Event::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    return Ok(start + reader.buffer_position() as usize);
                }
            },
            Ok(Event::Eof) => return Err(invalid("unterminated placeholder shape")),
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
}

/// Extract the `p:cNvPr/@id` shape ID from a shape byte range.
pub(super) fn shape_id_within(bytes: &[u8]) -> Result<u32> {
    let mut reader = Reader::from_reader(bytes);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element))
                if local_name(element.name().as_ref()) == b"cNvPr" =>
            {
                for attribute in element.attributes().with_checks(true) {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    if attribute.key.as_ref() == b"id" {
                        let value = std::str::from_utf8(attribute.value.as_ref())
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        return value
                            .parse::<u32>()
                            .map_err(|_err| invalid("invalid shape ID in placeholder"));
                    }
                }
                return Err(invalid("placeholder shape has no shape ID"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    Err(invalid("placeholder shape has no non-visual properties"))
}

/// Allocate the next free shape ID for a part (max existing + 1, starting at 2).
pub(super) fn next_shape_id(xml: &[u8]) -> Result<u32> {
    check_size(xml)?;
    let mut reader = Reader::from_reader(xml);
    let mut max_id = FIRST_SHAPE_ID - 1;
    let mut nodes = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(element) | Event::Empty(element)) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES {
                    return Err(invalid("part XML resource limit exceeded"));
                }
                if local_name(element.name().as_ref()) == b"cNvPr" {
                    for attribute in element.attributes().with_checks(true) {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        if attribute.key.as_ref() == b"id" {
                            let value = std::str::from_utf8(attribute.value.as_ref())
                                .map_err(|error| Error::Xml(error.to_string()))?;
                            let id = value
                                .parse::<u32>()
                                .map_err(|_err| invalid(format!("invalid shape ID '{value}'")))?;
                            max_id = max_id.max(id);
                        }
                    }
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(Error::Xml(error.to_string())),
            _ => {},
        }
    }
    max_id
        .checked_add(1)
        .ok_or_else(|| invalid("shape ID overflow"))
}

pub(super) fn replace_span(xml: &[u8], span: &ElementSpan, replacement: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len() + replacement.len());
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[span.end..]);
    check_size(&output)?;
    Ok(output)
}

pub(super) fn insert_bytes(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len() + value.len());
    output.extend_from_slice(&xml[..offset]);
    output.extend_from_slice(value);
    output.extend_from_slice(&xml[offset..]);
    check_size(&output)?;
    Ok(output)
}
