use super::model::{PlaceholderSpec, SlideLayoutKind};
use crate::shape::{PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderTypeExtension};
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use quick_xml::Reader;
use quick_xml::events::Event;
use std::fmt::{self, Write as FmtWrite};

pub(super) const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(super) const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const P232_NS: &str = "http://schemas.microsoft.com/office/powerpoint/2023/02/main";
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

#[derive(Default)]
struct LengthWriter {
    length: usize,
}

impl LengthWriter {
    fn add(&mut self, length: usize) -> fmt::Result {
        self.length = self.length.checked_add(length).ok_or(fmt::Error)?;
        Ok(())
    }
}

impl FmtWrite for LengthWriter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.add(value.len())
    }
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

/// Check the XML 1.0 fifth-edition `Char` production before any authored
/// value is copied into an output buffer. Rust strings cannot contain UTF-16
/// surrogate code points, so the scalar-value check covers the remaining
/// ranges directly.
pub(super) fn validate_xml10_chars(value: &str, resource: &'static str) -> Result<()> {
    if let Some(character) = value.chars().find(|character| {
        let code = u32::from(*character);
        !matches!(code, 0x09 | 0x0A | 0x0D)
            && !(0x20..=0xD7FF).contains(&code)
            && !(0xE000..=0xFFFD).contains(&code)
            && !(0x10000..=0x10FFFF).contains(&code)
    }) {
        return Err(Error::Invalid(format!(
            "{resource} contains XML 1.0-invalid character U+{:04X}",
            u32::from(character)
        )));
    }
    Ok(())
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    validate_xml10_chars(value, "authored XML text")?;
    value.chars().try_fold(0usize, |length, character| {
        let extra = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length.checked_add(extra).ok_or(Error::Limit {
            resource: "generated slide layout escaped XML bytes",
            limit: MAX_PART_XML_BYTES,
        })
    })
}

fn escape_xml_fallible(value: &str) -> Result<String> {
    let length = escaped_xml_len(value)?;
    if length > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "generated slide layout escaped XML bytes",
            limit: MAX_PART_XML_BYTES,
        });
    }
    let mut escaped = String::new();
    escaped
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "generated slide layout escaped XML bytes",
            source,
        })?;
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            '\t' => escaped.push_str("&#x9;"),
            '\n' => escaped.push_str("&#xA;"),
            '\r' => escaped.push_str("&#xD;"),
            _ => escaped.push(character),
        }
    }
    Ok(escaped)
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
    validate_layout_inputs(name, placeholders)?;
    let escaped_name = escape_xml_fallible(name)?;

    let mut length = LengthWriter::default();
    write_layout_header(
        &mut length,
        kind,
        &escaped_name,
        placeholders
            .iter()
            .any(|placeholder| placeholder.type_extension.is_some()),
    )
    .map_err(|_| Error::Limit {
        resource: "generated slide layout bytes",
        limit: MAX_PART_XML_BYTES,
    })?;
    length.write_str(SP_TREE_HEADER).map_err(|_| Error::Limit {
        resource: "generated slide layout bytes",
        limit: MAX_PART_XML_BYTES,
    })?;
    for (offset, spec) in placeholders.iter().enumerate() {
        length
            .add(placeholder_shape_xml_len(
                FIRST_SHAPE_ID + offset as u32,
                spec,
                false,
            )?)
            .map_err(|_| Error::Limit {
                resource: "generated slide layout bytes",
                limit: MAX_PART_XML_BYTES,
            })?;
    }
    length
        .write_str(
            "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
        )
        .map_err(|_| Error::Limit {
            resource: "generated slide layout bytes",
            limit: MAX_PART_XML_BYTES,
        })?;
    if length.length > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "generated slide layout bytes",
            limit: MAX_PART_XML_BYTES,
        });
    }

    let mut xml = String::new();
    xml.try_reserve_exact(length.length)
        .map_err(|source| Error::Allocation {
            resource: "generated slide layout bytes",
            source,
        })?;
    write_layout_header(
        &mut xml,
        kind,
        &escaped_name,
        placeholders
            .iter()
            .any(|placeholder| placeholder.type_extension.is_some()),
    )
    .map_err(|_| invalid("generated slide layout formatting failed"))?;
    xml.push_str(SP_TREE_HEADER);
    for (offset, spec) in placeholders.iter().enumerate() {
        let shape_id = FIRST_SHAPE_ID + offset as u32;
        xml.push_str(&placeholder_shape_xml(shape_id, spec, false)?);
    }
    xml.push_str(
        "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>",
    );
    if xml.len() != length.length {
        return Err(invalid("generated slide layout length preflight disagreed"));
    }
    Ok(xml)
}

fn write_layout_header<W: FmtWrite>(
    output: &mut W,
    kind: SlideLayoutKind,
    escaped_name: &str,
    declares_p232: bool,
) -> fmt::Result {
    output.write_str(XML_DECL)?;
    write!(
        output,
        "<p:sldLayout xmlns:a=\"{A_NS}\" xmlns:r=\"{R_NS}\" xmlns:p=\"{P_NS}\""
    )?;
    if declares_p232 {
        write!(output, " xmlns:p232=\"{P232_NS}\"")?;
    }
    write!(
        output,
        " type=\"{}\" matchingName=\"{}\"><p:cSld name=\"{}\">",
        kind.as_str(),
        escaped_name,
        escaped_name
    )
}

/// Check every user supplied string before a generated layout buffer is allocated.
pub(crate) fn validate_layout_inputs(name: &str, placeholders: &[PlaceholderSpec]) -> Result<()> {
    validate_xml10_chars(name, "slide layout name")?;
    let _name_length = escaped_xml_len(name)?;
    for spec in placeholders {
        if let Some(value) = spec.name.as_deref() {
            validate_xml10_chars(value, "placeholder name")?;
            let _ = escaped_xml_len(value)?;
        }
        if let Some(value) = spec.text.as_deref() {
            validate_xml10_chars(value, "placeholder text")?;
            let _ = escaped_xml_len(value)?;
        }
    }
    Ok(())
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
) -> Result<String> {
    let (escaped_name, escaped_text) = escaped_placeholder_values(shape_id, spec)?;
    let length = placeholder_shape_len_from_escaped(
        shape_id,
        spec,
        declare_namespaces,
        &escaped_name,
        escaped_text.as_deref(),
    )?;
    let mut xml = String::new();
    xml.try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "generated placeholder shape bytes",
            source,
        })?;
    write_placeholder_shape(
        &mut xml,
        shape_id,
        spec,
        declare_namespaces,
        &escaped_name,
        escaped_text.as_deref(),
    )
    .map_err(|_| invalid("generated placeholder shape formatting failed"))?;
    if xml.len() != length {
        return Err(invalid(
            "generated placeholder shape length preflight disagreed",
        ));
    }
    Ok(xml)
}

fn placeholder_shape_xml_len(
    shape_id: u32,
    spec: &PlaceholderSpec,
    declare_namespaces: bool,
) -> Result<usize> {
    let (escaped_name, escaped_text) = escaped_placeholder_values(shape_id, spec)?;
    placeholder_shape_len_from_escaped(
        shape_id,
        spec,
        declare_namespaces,
        &escaped_name,
        escaped_text.as_deref(),
    )
}

fn escaped_placeholder_values(
    shape_id: u32,
    spec: &PlaceholderSpec,
) -> Result<(String, Option<String>)> {
    if let Some(name) = spec.name.as_deref() {
        validate_xml10_chars(name, "placeholder name")?;
    }
    if let Some(text) = spec.text.as_deref() {
        validate_xml10_chars(text, "placeholder text")?;
    }
    let name = spec
        .name
        .clone()
        .unwrap_or_else(|| format!("{} Placeholder {shape_id}", spec.kind.label()));
    let escaped_name = escape_xml_fallible(&name)?;
    let escaped_text = spec.text.as_deref().map(escape_xml_fallible).transpose()?;
    Ok((escaped_name, escaped_text))
}

fn placeholder_shape_len_from_escaped(
    shape_id: u32,
    spec: &PlaceholderSpec,
    declare_namespaces: bool,
    escaped_name: &str,
    escaped_text: Option<&str>,
) -> Result<usize> {
    let mut length = LengthWriter::default();
    write_placeholder_shape(
        &mut length,
        shape_id,
        spec,
        declare_namespaces,
        escaped_name,
        escaped_text,
    )
    .map_err(|_| Error::Limit {
        resource: "generated placeholder shape bytes",
        limit: MAX_PART_XML_BYTES,
    })?;
    if length.length > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "generated placeholder shape bytes",
            limit: MAX_PART_XML_BYTES,
        });
    }
    Ok(length.length)
}

fn write_placeholder_shape<W: FmtWrite>(
    output: &mut W,
    shape_id: u32,
    spec: &PlaceholderSpec,
    declare_namespaces: bool,
    escaped_name: &str,
    escaped_text: Option<&str>,
) -> fmt::Result {
    output.write_str("<p:sp")?;
    if declare_namespaces {
        write!(output, " xmlns:p=\"{P_NS}\" xmlns:a=\"{A_NS}\"")?;
        if spec.type_extension.is_some() {
            write!(output, " xmlns:p232=\"{P232_NS}\"")?;
        }
    }
    write!(
        output,
        "><p:nvSpPr><p:cNvPr id=\"{shape_id}\" name=\"{escaped_name}\"/><p:cNvSpPr><a:spLocks noGrp=\"1\"/></p:cNvSpPr><p:nvPr><p:ph type=\"{}\"",
        spec.kind.as_str()
    )?;
    if let Some(index) = spec.index {
        write!(output, " idx=\"{index}\"")?;
    }
    if let Some(extension) = spec.type_extension {
        write!(
            output,
            "><p:extLst><p:ext uri=\"{PLACEHOLDER_TYPE_EXTENSION_URI}\"><p232:phTypeExt><p232:type>"
        )?;
        match extension {
            PlaceholderTypeExtension::Cameo => output.write_str("<p232:cameo/>")?,
            PlaceholderTypeExtension::Unknown => output.write_str("<p232:unknown/>")?,
        }
        output.write_str("</p232:type></p232:phTypeExt></p:ext></p:extLst></p:ph>")?;
    } else {
        output.write_str("/>")?;
    }
    output.write_str("</p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>")?;
    if let Some(text) = escaped_text {
        write!(output, "<a:r><a:t>{text}</a:t></a:r>")?;
    }
    output.write_str("<a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>")
}

// ============================================================================
// Bounded XML scanning and patching
// ============================================================================

pub(super) const SPTREE_DEPTH: usize = 3;

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
    let origin = ReaderOrigin::of(xml);
    let mut stack: Vec<(usize, String)> = Vec::new();
    let mut nodes = 0usize;
    loop {
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML source position exceeds usize"))?;
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
                        end: origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("XML source position exceeds usize"))?,
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
                        end: origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("XML source position exceeds usize"))?,
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
    let origin = ReaderOrigin::of(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(_) | Event::Empty(_)) => {
                return origin
                    .offset(reader.buffer_position())
                    .ok_or_else(|| invalid("XML source position exceeds usize"));
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
    let origin = ReaderOrigin::of(xml);
    let mut nodes = 0usize;
    loop {
        let before = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("XML source position exceeds usize"))?;
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
                        end: origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("XML source position exceeds usize"))?,
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
                                        end: origin.offset(reader.buffer_position()).ok_or_else(
                                            || invalid("XML source position exceeds usize"),
                                        )?,
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
    if span.start > span.end || span.end > xml.len() {
        return Err(invalid("replacement span is outside the owner XML"));
    }
    let target_len = xml
        .len()
        .checked_sub(span.end - span.start)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(Error::Limit {
            resource: "generated slide layout bytes",
            limit: MAX_PART_XML_BYTES,
        })?;
    if target_len > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "generated slide layout bytes",
            limit: MAX_PART_XML_BYTES,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(target_len)
        .map_err(|source| Error::Allocation {
            resource: "generated slide layout bytes",
            source,
        })?;
    output.extend_from_slice(&xml[..span.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[span.end..]);
    Ok(output)
}

pub(super) fn insert_bytes(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    if offset > xml.len() {
        return Err(invalid("insertion offset is outside the owner XML"));
    }
    let target_len = xml.len().checked_add(value.len()).ok_or(Error::Limit {
        resource: "generated slide layout bytes",
        limit: MAX_PART_XML_BYTES,
    })?;
    if target_len > MAX_PART_XML_BYTES {
        return Err(Error::Limit {
            resource: "generated slide layout bytes",
            limit: MAX_PART_XML_BYTES,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(target_len)
        .map_err(|source| Error::Allocation {
            resource: "generated slide layout bytes",
            source,
        })?;
    output.extend_from_slice(&xml[..offset]);
    output.extend_from_slice(value);
    output.extend_from_slice(&xml[offset..]);
    Ok(output)
}
