//! Namespace-aware parsing and lossless XML replacement for list alignments.

use super::model::{Alignment, FollowedBy, Kind, Length, Style, Styles};
use super::{
    FO, MAX_DEPTH, MAX_ENTRIES, MAX_LEVEL, MAX_TOTAL, MAX_VALUE, MAX_XML, OFFICE, STYLE, TEXT, bad,
};
use litchi_core::Result;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum Namespace {
    Office,
    Style,
    Text,
    Fo,
    Other,
}

fn namespace(result: ResolveResult<'_>) -> Namespace {
    match result {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => Namespace::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => Namespace::Style,
        ResolveResult::Bound(value) if value.as_ref() == TEXT => Namespace::Text,
        ResolveResult::Bound(value) if value.as_ref() == FO => Namespace::Fo,
        _ => Namespace::Other,
    }
}

fn element(reader: &NsReader<&[u8]>, qualified_name: QName<'_>) -> (Namespace, Vec<u8>) {
    let (resolved_namespace, local_name) = reader.resolver().resolve_element(qualified_name);
    (namespace(resolved_namespace), local_name.as_ref().to_vec())
}

fn attributes(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &BytesStart<'_>,
) -> Result<Vec<(Namespace, Vec<u8>, String)>> {
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute =
            attribute.map_err(|error| bad(format!("invalid list alignment attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved_namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let key = (namespace(resolved_namespace), local_name.as_ref().to_vec());
        if !seen.insert(key.clone()) {
            return Err(bad("duplicate list alignment attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| bad(format!("invalid list alignment value: {error}")))?
            .into_owned();
        if value.len() > MAX_VALUE {
            return Err(bad("list alignment value too large"));
        }
        attributes.push((key.0, key.1, value));
    }
    Ok(attributes)
}

fn take(
    attributes: &mut Vec<(Namespace, Vec<u8>, String)>,
    namespace: Namespace,
    local_name: &[u8],
) -> Option<String> {
    attributes
        .iter()
        .position(|attribute| attribute.0 == namespace && attribute.1 == local_name)
        .map(|index| attributes.remove(index).2)
}

fn parse_alignment(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    element: &BytesStart<'_>,
) -> Result<Alignment> {
    let mut attributes = attributes(reader, version, element)?;
    let followed_by = take(&mut attributes, Namespace::Text, b"label-followed-by")
        .ok_or_else(|| bad("missing text:label-followed-by"))?;
    let alignment = Alignment {
        label_followed_by: FollowedBy::parse(&followed_by)?,
        list_tab_stop_position: take(&mut attributes, Namespace::Text, b"list-tab-stop-position")
            .map(Length::new)
            .transpose()?,
        text_indent: take(&mut attributes, Namespace::Fo, b"text-indent")
            .map(Length::new)
            .transpose()?,
        margin_left: take(&mut attributes, Namespace::Fo, b"margin-left")
            .map(Length::new)
            .transpose()?,
    };
    if !attributes.is_empty() {
        return Err(bad("unknown list-level-label-alignment attribute"));
    }
    alignment.validate()?;
    Ok(alignment)
}

/// Parse every modern list-level label alignment in styles or flat-document XML.
pub fn parse(xml: &str) -> Result<Styles> {
    if xml.len() > MAX_XML {
        return Err(bad("XML too large"));
    }
    if !xml.contains("list-level-label-alignment") {
        return Ok(Default::default());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Namespace, Vec<u8>)> = Vec::new();
    let mut list: Option<(usize, String, Kind, HashSet<u16>)> = None;
    let mut level: Option<(usize, u16, bool, bool)> = None;
    let mut entries = Vec::new();
    let mut total = 0usize;
    let mut open_alignment = None;
    loop {
        match reader.read_event() {
            Ok(Event::Decl(declaration)) => {
                version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::Start(element_start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("XML nesting too deep"));
                }
                if open_alignment.is_some() {
                    return Err(bad("list-level-label-alignment must be empty"));
                }
                let current = element(&reader, element_start.name());
                let parent = stack.last();
                let direct_list = parent.is_some_and(|parent| {
                    parent.0 == Namespace::Office
                        && matches!(parent.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Namespace::Text
                    && matches!(current.1.as_slice(), b"list-style" | b"outline-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct_list {
                    let mut attributes = attributes(&reader, version, &element_start)?;
                    let name = take(&mut attributes, Namespace::Style, b"name")
                        .ok_or_else(|| bad("list style missing style:name"))?;
                    if name.is_empty() {
                        return Err(bad("empty list style name"));
                    }
                    let kind = if current.1 == b"outline-style" {
                        Kind::Outline
                    } else {
                        Kind::List
                    };
                    list = Some((depth, name, kind, HashSet::new()));
                    continue;
                }
                if let Some((list_depth, _, _, seen)) = list.as_mut()
                    && depth == *list_depth + 1
                    && current.0 == Namespace::Text
                    && matches!(
                        current.1.as_slice(),
                        b"list-level-style-number"
                            | b"list-level-style-bullet"
                            | b"list-level-style-image"
                            | b"outline-level-style"
                    )
                {
                    let mut attributes = attributes(&reader, version, &element_start)?;
                    let number = take(&mut attributes, Namespace::Text, b"level")
                        .ok_or_else(|| bad("list level missing text:level"))?
                        .parse::<u16>()
                        .map_err(|_error| bad("invalid text:level"))?;
                    if !(1..=MAX_LEVEL).contains(&number) || !seen.insert(number) {
                        return Err(bad("invalid or duplicate list level"));
                    }
                    level = Some((depth, number, false, false));
                    continue;
                }
                if let Some((level_depth, level_number, properties_seen, alignment_seen)) =
                    level.as_mut()
                {
                    if depth == *level_depth + 1
                        && current.0 == Namespace::Style
                        && current.1 == b"list-level-properties"
                    {
                        if *properties_seen {
                            return Err(bad("duplicate list-level-properties"));
                        }
                        *properties_seen = true;
                        let mut attributes = attributes(&reader, version, &element_start)?;
                        if take(
                            &mut attributes,
                            Namespace::Text,
                            b"list-level-position-and-space-mode",
                        )
                        .as_deref()
                            != Some("label-alignment")
                        {
                            return Err(bad(
                                "label alignment requires label-alignment position mode",
                            ));
                        }
                    } else if depth == *level_depth + 2
                        && current.0 == Namespace::Style
                        && current.1 == b"list-level-label-alignment"
                    {
                        if !*properties_seen || *alignment_seen {
                            return Err(bad("invalid or duplicate list-level-label-alignment"));
                        }
                        *alignment_seen = true;
                        let alignment = parse_alignment(&reader, version, &element_start)?;
                        let (_, list_name, kind, _) = list
                            .as_ref()
                            .ok_or_else(|| bad("list-level alignment has no parent list"))?;
                        let style =
                            Style::new_in(*kind, list_name.clone(), *level_number, alignment)?;
                        total +=
                            style.list_style_name.len() + style.alignment.to_xml_fragment()?.len();
                        if entries.len() >= MAX_ENTRIES || total > MAX_TOTAL {
                            return Err(bad("too many list alignments"));
                        }
                        entries.push(style);
                        open_alignment = Some(depth);
                    } else if current.1 == b"list-level-label-alignment" {
                        return Err(bad(
                            "list-level-label-alignment has invalid parent or namespace",
                        ));
                    }
                } else if current.1 == b"list-level-label-alignment" {
                    return Err(bad("list-level-label-alignment has invalid parent"));
                }
            },
            Ok(Event::Empty(element_empty)) => {
                let current = element(&reader, element_empty.name());
                let depth = stack.len() + 1;
                if let Some((level_depth, level_number, properties_seen, alignment_seen)) =
                    level.as_mut()
                {
                    if depth == *level_depth + 1
                        && current.0 == Namespace::Style
                        && current.1 == b"list-level-properties"
                    {
                        if *properties_seen {
                            return Err(bad("duplicate list-level-properties"));
                        }
                        *properties_seen = true;
                        return Err(bad(
                            "empty list-level-properties cannot contain label alignment",
                        ));
                    }
                    if depth == *level_depth + 2
                        && current.0 == Namespace::Style
                        && current.1 == b"list-level-label-alignment"
                    {
                        if !*properties_seen || *alignment_seen {
                            return Err(bad("invalid or duplicate list-level-label-alignment"));
                        }
                        *alignment_seen = true;
                        let alignment = parse_alignment(&reader, version, &element_empty)?;
                        let (_, list_name, kind, _) = list
                            .as_ref()
                            .ok_or_else(|| bad("list-level alignment has no parent list"))?;
                        let style =
                            Style::new_in(*kind, list_name.clone(), *level_number, alignment)?;
                        total +=
                            style.list_style_name.len() + style.alignment.to_xml_fragment()?.len();
                        if entries.len() >= MAX_ENTRIES || total > MAX_TOTAL {
                            return Err(bad("too many list alignments"));
                        }
                        entries.push(style);
                    } else if current.1 == b"list-level-label-alignment" {
                        return Err(bad(
                            "list-level-label-alignment has invalid parent or namespace",
                        ));
                    }
                } else if current.1 == b"list-level-label-alignment" {
                    return Err(bad("list-level-label-alignment has invalid parent"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if open_alignment == Some(depth) {
                    open_alignment = None;
                }
                if level.as_ref().is_some_and(|value| value.0 == depth) {
                    level = None;
                }
                if list.as_ref().is_some_and(|value| value.0 == depth) {
                    list = None;
                }
                stack.pop();
            },
            Ok(Event::Text(text)) if open_alignment.is_some() => {
                let bytes: &[u8] = text.as_ref();
                if !bytes.is_empty() {
                    return Err(bad("list-level-label-alignment must be empty"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid XML: {error}"))),
        }
    }
    Ok(Styles { levels: entries })
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML boundary"))
}

/// Replace one existing alignment element, preserving every unrelated byte.
pub(crate) fn set_xml(xml: &str, item: &Style) -> Result<String> {
    item.validate()?;
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(Namespace, Vec<u8>)> = Vec::new();
    let mut list: Option<(usize, bool)> = None;
    let mut level: Option<(usize, bool)> = None;
    let mut found = None;
    loop {
        match reader.read_event() {
            Ok(Event::Decl(declaration)) => {
                version = declaration
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?;
            },
            Ok(Event::Start(element_start)) => {
                let end = reader.buffer_position() as usize;
                let current = element(&reader, element_start.name());
                let parent = stack.last();
                let direct_list = parent.is_some_and(|parent| {
                    parent.0 == Namespace::Office
                        && matches!(parent.1.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == Namespace::Text
                    && matches!(current.1.as_slice(), b"list-style" | b"outline-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct_list {
                    let mut attributes = attributes(&reader, version, &element_start)?;
                    list = Some((
                        depth,
                        take(&mut attributes, Namespace::Style, b"name").as_deref()
                            == Some(&item.list_style_name)
                            && (if current.1 == b"outline-style" {
                                Kind::Outline
                            } else {
                                Kind::List
                            }) == item.list_style_kind,
                    ));
                } else if list == Some((depth - 1, true))
                    && current.0 == Namespace::Text
                    && matches!(
                        current.1.as_slice(),
                        b"list-level-style-number"
                            | b"list-level-style-bullet"
                            | b"list-level-style-image"
                            | b"outline-level-style"
                    )
                {
                    let mut attributes = attributes(&reader, version, &element_start)?;
                    level = Some((
                        depth,
                        take(&mut attributes, Namespace::Text, b"level")
                            .and_then(|value| value.parse().ok())
                            == Some(item.level),
                    ));
                } else if level
                    .is_some_and(|(level_depth, enabled)| enabled && depth == level_depth + 2)
                    && current.0 == Namespace::Style
                    && current.1 == b"list-level-label-alignment"
                {
                    let start = event_start(xml, end)?;
                    found = Some((start, 0usize, depth));
                }
            },
            Ok(Event::Empty(element_empty)) => {
                let end = reader.buffer_position() as usize;
                let current = element(&reader, element_empty.name());
                let depth = stack.len() + 1;
                if level.is_some_and(|(level_depth, enabled)| enabled && depth == level_depth + 2)
                    && current.0 == Namespace::Style
                    && current.1 == b"list-level-label-alignment"
                {
                    if found.is_some() {
                        return Err(bad("duplicate target alignment"));
                    }
                    found = Some((event_start(xml, end)?, end, 0));
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if let Some((start, 0, target_depth)) = found
                    && depth == target_depth
                {
                    found = Some((start, end, 0));
                }
                if level.as_ref().is_some_and(|value| value.0 == depth) {
                    level = None;
                }
                if list.as_ref().is_some_and(|value| value.0 == depth) {
                    list = None;
                }
                stack.pop();
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid XML: {error}"))),
        }
    }
    let (start, end, _) = found
        .filter(|value| value.1 > 0)
        .ok_or_else(|| bad("target list-level-label-alignment does not exist"))?;
    Ok(format!(
        "{}{}{}",
        &xml[..start],
        item.alignment.to_xml_fragment()?,
        &xml[end..]
    ))
}
