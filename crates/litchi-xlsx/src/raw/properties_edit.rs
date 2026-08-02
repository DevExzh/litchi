//! Lossless extended-properties synchronization for worksheet structure.

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Result, allocation, invalid};

const EXTENDED: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const TYPES: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
const MAX_DEPTH: usize = 64;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Properties,
    Titles,
    TitleVector,
    Title,
    HeadingPairs,
    HeadingVector,
    Variant,
    Label,
    Count,
    Other,
}

#[derive(Debug)]
struct Frame {
    kind: Kind,
    start: usize,
    inner_start: usize,
    tag: Option<Tag>,
    text: String,
    markup: bool,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug, Clone)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug)]
struct Vector {
    start: usize,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    size: Option<usize>,
    base_type: Option<Box<str>>,
}

#[derive(Debug)]
struct Title {
    start: usize,
    end: usize,
    tag: Tag,
    text: String,
}

#[derive(Debug, Default)]
struct Variant {
    label: Option<String>,
    count: Option<(usize, usize, usize)>,
}

#[derive(Debug)]
struct Replacement {
    start: usize,
    end: usize,
    bytes: Vec<u8>,
}

/// One final worksheet-title slot. Existing entries retain their complete XML
/// bytes; new entries synthesize only the required title element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Sheet<'a> {
    Existing(usize),
    New(&'a str),
}

/// Insert new worksheet titles after the existing worksheet title prefix.
///
/// Producers may omit these optional properties. An unrecognized or stale
/// layout is preserved byte-exact instead of being guessed at.
#[cfg(test)]
pub(crate) fn append_sheets(
    content: &[u8],
    existing: &[&str],
    added: &[&str],
) -> Result<Option<Vec<u8>>> {
    if added.is_empty() {
        return Ok(None);
    }
    let mut order = Vec::new();
    order
        .try_reserve_exact(existing.len().saturating_add(added.len()))
        .map_err(|source| allocation("appended property order", source))?;
    order.extend((0..existing.len()).map(Sheet::Existing));
    order.extend(added.iter().copied().map(Sheet::New));
    arrange_sheets(content, existing, &order)
}

/// Synchronize a standard worksheet-title prefix with one checked final order.
///
/// Named-range titles after the worksheet prefix remain byte-exact. Stale or
/// producer-specific optional layouts are preserved rather than guessed.
pub(crate) fn arrange_sheets(
    content: &[u8],
    existing: &[&str],
    order: &[Sheet<'_>],
) -> Result<Option<Vec<u8>>> {
    if order.len() < existing.len() {
        return Err(invalid(
            "extended-properties worksheet order omits existing sheets",
        ));
    }
    let mut seen = Vec::new();
    seen.try_reserve_exact(existing.len())
        .map_err(|source| allocation("property-order validation", source))?;
    seen.resize(existing.len(), false);
    let mut new_count = 0usize;
    let mut identity = order.len() == existing.len();
    for (position, entry) in order.iter().copied().enumerate() {
        match entry {
            Sheet::Existing(index) => {
                let slot = seen.get_mut(index).ok_or_else(|| {
                    invalid("extended-properties order has an unknown existing sheet")
                })?;
                if std::mem::replace(slot, true) {
                    return Err(invalid(
                        "extended-properties order repeats an existing sheet",
                    ));
                }
                identity &= position == index;
            },
            Sheet::New(_) => {
                new_count = new_count
                    .checked_add(1)
                    .ok_or_else(|| invalid("extended-properties new-sheet count overflow"))?;
                identity = false;
            },
        }
    }
    if seen.contains(&false) {
        return Err(invalid("extended-properties order omits an existing sheet"));
    }
    if identity {
        return Ok(None);
    }

    let (vector, titles, variants) = scan(content)?;
    let Some(vector) = vector else {
        return Ok(None);
    };
    if vector
        .base_type
        .as_deref()
        .is_some_and(|value| value != "lpstr")
        || vector.size.is_some_and(|size| size != titles.len())
        || titles.len() < existing.len()
        || !titles
            .iter()
            .take(existing.len())
            .zip(existing)
            .all(|(actual, expected)| actual.text == *expected)
    {
        return Ok(None);
    }

    let final_size = titles
        .len()
        .checked_add(new_count)
        .ok_or_else(|| invalid("extended-properties title count overflow"))?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve(3)
        .map_err(|source| allocation("property replacements", source))?;
    replacements.push(Replacement {
        start: vector.start,
        end: vector.tag_end,
        bytes: rewrite_tag(&vector.tag, "size", &final_size.to_string()),
    });

    let title_name = titles.first().map_or_else(
        || sibling_name(&vector.tag.name, "lpstr"),
        |title| title.tag.name.to_string(),
    );
    let prefix_start = titles
        .first()
        .map_or(vector.close_start, |title| title.start);
    let prefix_end = titles
        .get(existing.len())
        .map_or(vector.close_start, |title| title.start);
    let mut arranged = Vec::new();
    let prefix_len = prefix_end
        .checked_sub(prefix_start)
        .ok_or_else(|| invalid("extended-properties title prefix is inverted"))?;
    arranged
        .try_reserve_exact(prefix_len)
        .map_err(|source| allocation("arranged sheet titles", source))?;
    for entry in order {
        match entry {
            Sheet::Existing(index) => {
                let title = titles.get(*index).ok_or_else(|| {
                    invalid("extended-properties existing title disappeared during reorder")
                })?;
                let title_len = title
                    .end
                    .checked_sub(title.start)
                    .ok_or_else(|| invalid("extended-properties title span is inverted"))?;
                arranged
                    .try_reserve(title_len)
                    .map_err(|source| allocation("existing sheet title", source))?;
                arranged.extend_from_slice(&content[title.start..title.end]);
            },
            Sheet::New(title) => {
                let escaped = escape_xml(title);
                let required = title_name
                    .len()
                    .checked_mul(2)
                    .and_then(|size| size.checked_add(5))
                    .and_then(|size| size.checked_add(escaped.len()))
                    .ok_or_else(|| invalid("new extended-properties title size overflow"))?;
                arranged
                    .try_reserve(required)
                    .map_err(|source| allocation("new sheet title", source))?;
                arranged.extend_from_slice(b"<");
                arranged.extend_from_slice(title_name.as_bytes());
                arranged.extend_from_slice(b">");
                arranged.extend_from_slice(escaped.as_bytes());
                arranged.extend_from_slice(b"</");
                arranged.extend_from_slice(title_name.as_bytes());
                arranged.extend_from_slice(b">");
            },
        }
    }
    replacements.push(Replacement {
        start: prefix_start,
        end: prefix_end,
        bytes: arranged,
    });

    if let Some(count) = variants.windows(2).find_map(|pair| {
        let label = pair[0].label.as_deref()?;
        matches!(label, "Worksheet" | "Worksheets")
            .then_some(pair[1].count)
            .flatten()
    }) && count.2 == existing.len()
    {
        replacements.push(Replacement {
            start: count.0,
            end: count.1,
            bytes: order.len().to_string().into_bytes(),
        });
    }

    replacements.sort_unstable_by_key(|replacement| (replacement.start, replacement.end));
    if replacements
        .windows(2)
        .any(|pair| pair[0].end > pair[1].start)
    {
        return Err(invalid("overlapping extended-properties replacements"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.end - replacement.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("extended-properties output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("extended properties", source))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

/// Remove selected worksheet titles from a standard title prefix.
///
/// Optional or producer-specific layouts are left byte-exact. The workbook
/// catalog remains authoritative, so stale metadata never blocks deletion.
pub(crate) fn remove_sheets(
    content: &[u8],
    existing: &[&str],
    removed: &[usize],
) -> Result<Option<Vec<u8>>> {
    if removed.is_empty() {
        return Ok(None);
    }
    if removed.windows(2).any(|pair| pair[0] >= pair[1])
        || removed.iter().any(|position| *position >= existing.len())
    {
        return Err(invalid(
            "extended-properties sheet removals must be unique, sorted, and in range",
        ));
    }
    let (vector, titles, variants) = scan(content)?;
    let Some(vector) = vector else {
        return Ok(None);
    };
    if vector
        .base_type
        .as_deref()
        .is_some_and(|value| value != "lpstr")
        || vector.size.is_some_and(|size| size != titles.len())
        || titles.len() < existing.len()
        || !titles
            .iter()
            .take(existing.len())
            .zip(existing)
            .all(|(actual, expected)| actual.text == *expected)
    {
        return Ok(None);
    }

    let final_size = titles
        .len()
        .checked_sub(removed.len())
        .ok_or_else(|| invalid("extended-properties title count underflow"))?;
    let mut replacements = Vec::new();
    replacements
        .try_reserve(removed.len().saturating_add(2))
        .map_err(|source| allocation("property removals", source))?;
    replacements.push(Replacement {
        start: vector.start,
        end: vector.tag_end,
        bytes: rewrite_tag(&vector.tag, "size", &final_size.to_string()),
    });
    for position in removed {
        let title = titles
            .get(*position)
            .ok_or_else(|| invalid("extended-properties removal escaped checked titles"))?;
        replacements.push(Replacement {
            start: title.start,
            end: title.end,
            bytes: Vec::new(),
        });
    }
    if let Some(count) = variants.windows(2).find_map(|pair| {
        let label = pair[0].label.as_deref()?;
        matches!(label, "Worksheet" | "Worksheets")
            .then_some(pair[1].count)
            .flatten()
    }) && count.2 == existing.len()
    {
        replacements.push(Replacement {
            start: count.0,
            end: count.1,
            bytes: existing
                .len()
                .checked_sub(removed.len())
                .ok_or_else(|| invalid("extended-properties worksheet count underflow"))?
                .to_string()
                .into_bytes(),
        });
    }

    replacements.sort_unstable_by_key(|replacement| (replacement.start, replacement.end));
    if replacements
        .windows(2)
        .any(|pair| pair[0].end > pair[1].start)
    {
        return Err(invalid("overlapping extended-properties removals"));
    }
    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            size.checked_sub(replacement.end - replacement.start)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("extended-properties removal size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("extended properties", source))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(Some(output))
}

fn scan(content: &[u8]) -> Result<(Option<Vector>, Vec<Title>, Vec<Variant>)> {
    let mut reader = NsReader::from_reader(content);
    let mut stack = Vec::<Frame>::new();
    let mut title_vector = None;
    let mut titles = Vec::new();
    let mut variants = Vec::new();
    let mut pending_variant = None::<Variant>;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("extended-properties scan failed: {error}")))?
            .into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid(format!(
                        "extended-properties nesting exceeds {MAX_DEPTH} levels"
                    )));
                }
                if let Some(parent) = stack.last_mut() {
                    parent.markup = true;
                }
                let parent = stack.last().map(|frame| frame.kind);
                let kind = kind(parent, &namespace, &element);
                if kind == Kind::Variant && pending_variant.replace(Variant::default()).is_some() {
                    return Err(invalid("nested extended-properties variants"));
                }
                let element_tag = matches!(kind, Kind::TitleVector | Kind::Title)
                    .then(|| tag(&element, decoder))
                    .transpose()?;
                stack.push(Frame {
                    kind,
                    start,
                    inner_start: end,
                    tag: element_tag,
                    text: String::new(),
                    markup: false,
                });
            },
            Event::Empty(element) => {
                if let Some(parent) = stack.last_mut() {
                    parent.markup = true;
                }
                let parent = stack.last().map(|frame| frame.kind);
                match kind(parent, &namespace, &element) {
                    Kind::TitleVector => return Ok((None, titles, variants)),
                    Kind::Title => titles.push(Title {
                        start,
                        end,
                        tag: tag(&element, decoder)?,
                        text: String::new(),
                    }),
                    _ => {},
                }
            },
            Event::Text(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, Kind::Title | Kind::Label | Kind::Count)
                {
                    frame.text.push_str(
                        &text
                            .decode()
                            .map_err(|error| invalid(format!("property text decode: {error}")))?,
                    );
                }
            },
            Event::CData(text) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, Kind::Title | Kind::Label | Kind::Count)
                {
                    frame.text.push_str(
                        &text
                            .decode()
                            .map_err(|error| invalid(format!("property CDATA decode: {error}")))?,
                    );
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(frame) = stack.last_mut()
                    && matches!(frame.kind, Kind::Title | Kind::Label | Kind::Count)
                {
                    frame.text.push_str(&decode_xml_reference(&reference)?);
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("extended-properties closing tag is unmatched"))?;
                match frame.kind {
                    Kind::TitleVector => {
                        if title_vector.is_some() {
                            return Err(invalid("duplicate TitlesOfParts vector"));
                        }
                        let tag = frame
                            .tag
                            .ok_or_else(|| invalid("title vector lost its start tag"))?;
                        title_vector = Some(Vector {
                            start: frame.start,
                            tag_end: frame.inner_start,
                            close_start: start,
                            size: optional_usize(&tag, "size")?,
                            base_type: optional_value(&tag, "baseType"),
                            tag,
                        });
                    },
                    Kind::Title if !frame.markup => titles.push(Title {
                        start: frame.start,
                        end,
                        tag: frame
                            .tag
                            .ok_or_else(|| invalid("title lost its start tag"))?,
                        text: frame.text,
                    }),
                    Kind::Label if !frame.markup => {
                        if let Some(variant) = pending_variant.as_mut() {
                            variant.label = Some(frame.text);
                        }
                    },
                    Kind::Count if !frame.markup => {
                        if let Some(variant) = pending_variant.as_mut() {
                            let value = frame.text.parse::<usize>().map_err(|_| {
                                invalid("invalid worksheet count in extended properties")
                            })?;
                            variant.count = Some((frame.inner_start, start, value));
                        }
                    },
                    Kind::Variant => variants.push(
                        pending_variant
                            .take()
                            .ok_or_else(|| invalid("extended-properties variant lost state"))?,
                    ),
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("extended-properties XML ended inside an element"));
    }
    Ok((title_vector, titles, variants))
}

fn kind(parent: Option<Kind>, namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> Kind {
    let local = element.name().local_name();
    let local = local.as_ref();
    if parent.is_none() && expanded(namespace, EXTENDED) && local == b"Properties" {
        Kind::Properties
    } else if parent == Some(Kind::Properties)
        && expanded(namespace, EXTENDED)
        && local == b"TitlesOfParts"
    {
        Kind::Titles
    } else if parent == Some(Kind::Titles) && expanded(namespace, TYPES) && local == b"vector" {
        Kind::TitleVector
    } else if parent == Some(Kind::TitleVector) && expanded(namespace, TYPES) && local == b"lpstr" {
        Kind::Title
    } else if parent == Some(Kind::Properties)
        && expanded(namespace, EXTENDED)
        && local == b"HeadingPairs"
    {
        Kind::HeadingPairs
    } else if parent == Some(Kind::HeadingPairs) && expanded(namespace, TYPES) && local == b"vector"
    {
        Kind::HeadingVector
    } else if parent == Some(Kind::HeadingVector)
        && expanded(namespace, TYPES)
        && local == b"variant"
    {
        Kind::Variant
    } else if parent == Some(Kind::Variant) && expanded(namespace, TYPES) && local == b"lpstr" {
        Kind::Label
    } else if parent == Some(Kind::Variant) && expanded(namespace, TYPES) && local == b"i4" {
        Kind::Count
    } else {
        Kind::Other
    }
}

fn expanded(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn optional_value(tag: &Tag, name: &str) -> Option<Box<str>> {
    tag.attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| attribute.value.clone())
}

fn optional_usize(tag: &Tag, name: &str) -> Result<Option<usize>> {
    optional_value(tag, name)
        .map(|value| {
            value
                .parse::<usize>()
                .map_err(|_| invalid(format!("invalid extended-properties {name} '{value}'")))
        })
        .transpose()
}

fn rewrite_tag(tag: &Tag, name: &str, value: &str) -> Vec<u8> {
    let mut output = Vec::new();
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if attribute.name.as_ref() == name {
            continue;
        }
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(&attribute.value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    output.extend_from_slice(b" ");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    output.extend_from_slice(escape_xml(value).as_bytes());
    output.extend_from_slice(b"\">");
    output
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| invalid(format!("property element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("property attribute is not UTF-8: {error}")))?
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

fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("extended-properties byte position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn appends_sheet_titles_and_updates_standard_heading_count() {
        let source = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr><vt:lpstr>One!Print_Area</vt:lpstr></vt:vector></TitlesOfParts><Company>keep</Company></Properties>"#;
        let output = append_sheets(source, &["One", "Two"], &["A&B", "Four"])
            .expect("rewrite")
            .expect("changed");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<vt:i4>4</vt:i4>"));
        assert!(text.contains("baseType=\"lpstr\""));
        assert!(text.contains("size=\"5\""));
        assert!(text.contains(
            "<vt:lpstr>One</vt:lpstr><vt:lpstr>Two</vt:lpstr><vt:lpstr>A&amp;B</vt:lpstr><vt:lpstr>Four</vt:lpstr><vt:lpstr>One!Print_Area</vt:lpstr>"
        ));
        assert!(text.contains("<Company>keep</Company>"));
    }

    #[test]
    fn stale_title_prefix_is_preserved_instead_of_guessed() {
        let source = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Stale</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#;
        assert!(
            append_sheets(source, &["Actual"], &["New"])
                .expect("scan")
                .is_none()
        );
    }

    #[test]
    fn arranges_sheet_prefix_without_touching_named_range_titles() {
        let source = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr data-keep="one">Data</vt:lpstr><vt:lpstr>Calc</vt:lpstr><vt:lpstr data-range="exact">Data!Print_Area</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#;
        let output = arrange_sheets(
            source,
            &["Data", "Calc"],
            &[
                Sheet::Existing(1),
                Sheet::New("Middle & More"),
                Sheet::Existing(0),
            ],
        )
        .expect("rewrite")
        .expect("changed");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<vt:i4>3</vt:i4>"));
        assert!(text.contains("size=\"4\""));
        assert!(text.contains(concat!(
            "<vt:lpstr>Calc</vt:lpstr>",
            "<vt:lpstr>Middle &amp; More</vt:lpstr>",
            "<vt:lpstr data-keep=\"one\">Data</vt:lpstr>",
            "<vt:lpstr data-range=\"exact\">Data!Print_Area</vt:lpstr>"
        )));
    }

    #[test]
    fn removes_sheet_titles_and_updates_standard_heading_count() {
        let source = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>One</vt:lpstr><vt:lpstr>Middle</vt:lpstr><vt:lpstr>Three</vt:lpstr><vt:lpstr>Named Range</vt:lpstr></vt:vector></TitlesOfParts><Company>keep</Company></Properties>"#;
        let output = remove_sheets(source, &["One", "Middle", "Three"], &[1])
            .expect("rewrite")
            .expect("changed");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains("<vt:i4>2</vt:i4>"));
        assert!(text.contains("size=\"3\""));
        assert!(text.contains(
            "<vt:lpstr>One</vt:lpstr><vt:lpstr>Three</vt:lpstr><vt:lpstr>Named Range</vt:lpstr>"
        ));
        assert!(text.contains("<Company>keep</Company>"));
    }

    #[test]
    fn preserves_stale_sheet_titles_during_removal() {
        let source = br#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Stale</vt:lpstr></vt:vector></TitlesOfParts></Properties>"#;
        assert!(
            remove_sheets(source, &["Actual", "Two"], &[1])
                .expect("scan")
                .is_none()
        );
    }
}
