//! New list definitions in `numbering.xml` (record 0251e, ported from the
//! L-DOCX-W10 draft): one abstract definition and one numbering instance per
//! list, spliced into the authored part without rewriting anything else.
//!
//! The definition is the one OfficeCLI's `listStyle` creates
//! (`WordHandler.StyleList.cs`, `ApplyListStyle`): a hybrid multi-level
//! definition of nine levels, bullets cycling through `•`, `◦`, `▪`, numbers
//! cycling through decimal, lower letter and lower Roman as `%N.`, each
//! level starting at 1, left-justified and indented by 720 twips per level
//! with a 360-twip hanging indent.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const BULLETS: [&str; 3] = ["\u{2022}", "\u{25E6}", "\u{25AA}"];
const LEVELS: u8 = 9;

/// The template of a new list definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ListKind {
    /// A bulleted list: `•`, `◦`, `▪` by level.
    Bullet,
    /// A numbered list: `1.`, `a.`, `i.` by level.
    Decimal,
}

/// The identifiers of a list definition a numbering edit added.
///
/// A paragraph joins the list through its numbering properties: `w:numId`
/// is [`Self::num_id`] (see [`crate::properties::Numbering`] and
/// [`crate::document::Edit::format_paragraph_at`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AddedList {
    pub(crate) kind: ListKind,
    pub(crate) abstract_num_id: u32,
    pub(crate) num_id: u32,
    pub(crate) start: Option<(u8, u32)>,
}

impl AddedList {
    /// The template the definition was made from.
    #[must_use]
    pub const fn kind(&self) -> ListKind {
        self.kind
    }

    /// The `w:abstractNumId` of the new abstract definition.
    #[must_use]
    pub const fn abstract_num_id(&self) -> u32 {
        self.abstract_num_id
    }

    /// The `w:numId` of the new numbering instance, which paragraphs name.
    #[must_use]
    pub const fn num_id(&self) -> u32 {
        self.num_id
    }

    /// The level whose numbering starts at a value other than 1, and that
    /// value.
    #[must_use]
    pub const fn start(&self) -> Option<(u8, u32)> {
        self.start
    }
}

/// The empty numbering part a document without one starts from.
pub(crate) fn empty_numbering_xml(strict: bool) -> Vec<u8> {
    let namespace = if strict { STRICT } else { TRANSITIONAL };
    let mut xml =
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w=""#
            .to_vec();
    xml.extend_from_slice(namespace);
    xml.extend_from_slice(br#""/>"#);
    xml
}

/// Where the root and its definitions sit.
struct Layout {
    /// The root's qualified-name prefix.
    prefix: Vec<u8>,
    /// Whether the root is in the Strict namespace.
    strict: bool,
    /// An empty root (`<w:numbering/>`): its start tag.
    empty_root: Option<Range<usize>>,
    /// The offset of the root's end tag.
    root_end: usize,
    last_abstract_end: Option<usize>,
    first_instance_start: Option<usize>,
    last_instance_end: Option<usize>,
    cleanup_start: Option<usize>,
    /// Every `w:abstractNumId` and `w:numId` written on a definition or
    /// instance, at any depth.
    abstract_ids: Vec<u32>,
    instance_ids: Vec<u32>,
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}

fn offset(origin: ReaderOrigin, position: u64) -> Result<usize> {
    origin
        .offset(position)
        .ok_or_else(|| invalid("numbering XML offset exceeds usize"))
}

fn scan(xml: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut layout = Layout {
        prefix: Vec::new(),
        strict: false,
        empty_root: None,
        root_end: 0,
        last_abstract_end: None,
        first_instance_start: None,
        last_instance_end: None,
        cleanup_start: None,
        abstract_ids: Vec::new(),
        instance_ids: Vec::new(),
    };
    let mut root_seen = false;
    // Open elements: whether each is a direct child of the root that the
    // layout tracks (`abstractNum`, `num`).
    let mut stack: Vec<Option<&'static str>> = Vec::new();
    loop {
        let start = offset(origin, reader.buffer_position())?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let word = match &namespace {
            ResolveResult::Bound(Namespace(value)) => {
                if *value == TRANSITIONAL {
                    Some(false)
                } else if *value == STRICT {
                    Some(true)
                } else {
                    None
                }
            },
            _ => None,
        };
        let opened = match event {
            Event::Start(element) => Some((element, false)),
            Event::Empty(element) => Some((element, true)),
            Event::End(_) => {
                let end = offset(origin, reader.buffer_position())?;
                let closed = stack
                    .pop()
                    .ok_or_else(|| invalid("numbering XML is unbalanced"))?;
                match (stack.len(), closed) {
                    (0, _) => {
                        layout.root_end = start;
                        break;
                    },
                    (1, Some("abstractNum")) => layout.last_abstract_end = Some(end),
                    (1, Some("num")) => layout.last_instance_end = Some(end),
                    _ => {},
                }
                continue;
            },
            Event::Eof => return Err(invalid("numbering XML has no root element")),
            Event::DocType(_) => return Err(invalid("DTD is not allowed in numbering XML")),
            _ => continue,
        };
        let Some((element, empty)) = opened else {
            continue;
        };
        let end = offset(origin, reader.buffer_position())?;
        let local = element.local_name();
        let local = local.as_ref();
        if !root_seen {
            let Some(strict) = word.filter(|_| local == b"numbering") else {
                return Err(invalid("numbering XML root is not w:numbering"));
            };
            root_seen = true;
            layout.strict = strict;
            layout.prefix = element
                .name()
                .prefix()
                .map(|prefix| prefix.as_ref().to_vec())
                .unwrap_or_default();
            if layout.prefix.is_empty() {
                return Err(invalid(
                    "numbering XML binds WordprocessingML as its default namespace; new definitions need a prefix for their attributes",
                ));
            }
            if empty {
                layout.empty_root = Some(start..end);
                layout.root_end = end;
                break;
            }
            stack.push(None);
            continue;
        }
        let depth = stack.len();
        let tracked = match (word.is_some(), local) {
            (true, b"abstractNum") => {
                if let Some(id) = id_attribute(&element, b"abstractNumId", &reader)? {
                    layout.abstract_ids.push(id);
                }
                Some("abstractNum")
            },
            (true, b"num") => {
                if let Some(id) = id_attribute(&element, b"numId", &reader)? {
                    layout.instance_ids.push(id);
                }
                if depth == 1 && layout.first_instance_start.is_none() {
                    layout.first_instance_start = Some(start);
                }
                Some("num")
            },
            (true, b"numIdMacAtCleanup") if depth == 1 => {
                layout.cleanup_start.get_or_insert(start);
                None
            },
            _ => None,
        };
        if empty {
            match (depth, tracked) {
                (1, Some("abstractNum")) => layout.last_abstract_end = Some(end),
                (1, Some("num")) => layout.last_instance_end = Some(end),
                _ => {},
            }
        } else {
            stack.push(tracked);
        }
    }
    Ok(layout)
}

/// The decimal value of a WordprocessingML attribute `local` on `element`.
fn id_attribute(
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    reader: &NsReader<&[u8]>,
) -> Result<Option<u32>> {
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != local {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        let in_word = matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if value == TRANSITIONAL || value == STRICT
        );
        if !in_word {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?;
        return Ok(value.trim().parse().ok());
    }
    Ok(None)
}

/// The identifiers the next list definition takes: one more than the
/// largest abstract definition and instance identifier in `xml`.
pub(crate) fn next_ids(xml: &[u8]) -> Result<(u32, u32)> {
    let layout = scan(xml)?;
    let abstract_id = match layout.abstract_ids.iter().copied().max() {
        Some(max) => max
            .checked_add(1)
            .ok_or_else(|| invalid("abstract numbering identifiers are exhausted"))?,
        None => 0,
    };
    let instance_id = layout
        .instance_ids
        .iter()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| invalid("numbering instance identifiers are exhausted"))?;
    Ok((abstract_id, instance_id))
}

/// Splice `lists` into `xml`: each abstract definition after the last
/// existing one (or before the first instance), each instance after the last
/// existing one (or before `w:numIdMacAtCleanup`), as `CT_Numbering`'s
/// sequence orders them. Every other byte is copied.
pub(crate) fn insert_lists(xml: &[u8], lists: &[AddedList]) -> Result<Vec<u8>> {
    if lists.is_empty() {
        return Ok(xml.to_vec());
    }
    let layout = scan(xml)?;
    let mut definitions = Vec::new();
    let mut instances = Vec::new();
    for list in lists {
        render_definition(&mut definitions, &layout.prefix, layout.strict, list);
        render_instance(&mut instances, &layout.prefix, list);
    }
    let mut output = Vec::with_capacity(xml.len() + definitions.len() + instances.len() + 32);
    if let Some(root) = &layout.empty_root {
        // `<w:numbering .../>`: open the root and close it after the lists.
        let tag = &xml[root.clone()];
        output.extend_from_slice(&xml[..root.start]);
        output.extend_from_slice(&tag[..tag.len() - 2]);
        output.push(b'>');
        output.extend_from_slice(&definitions);
        output.extend_from_slice(&instances);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&layout.prefix);
        output.extend_from_slice(b":numbering>");
        output.extend_from_slice(&xml[root.end..]);
        return Ok(output);
    }
    let definitions_at = layout
        .last_abstract_end
        .or(layout.first_instance_start)
        .or(layout.cleanup_start)
        .unwrap_or(layout.root_end);
    let instances_at = layout
        .last_instance_end
        .or(layout.cleanup_start)
        .unwrap_or(layout.root_end)
        .max(definitions_at);
    output.extend_from_slice(&xml[..definitions_at]);
    output.extend_from_slice(&definitions);
    output.extend_from_slice(&xml[definitions_at..instances_at]);
    output.extend_from_slice(&instances);
    output.extend_from_slice(&xml[instances_at..]);
    Ok(output)
}

fn push_element(
    output: &mut Vec<u8>,
    prefix: &[u8],
    local: &str,
    attributes: &[(&str, &str)],
    empty: bool,
) {
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.push(b':');
    output.extend_from_slice(local.as_bytes());
    for (name, value) in attributes {
        output.push(b' ');
        output.extend_from_slice(prefix);
        output.push(b':');
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(value.as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
}

fn push_end(output: &mut Vec<u8>, prefix: &[u8], local: &str) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.push(b':');
    output.extend_from_slice(local.as_bytes());
    output.push(b'>');
}

fn render_definition(output: &mut Vec<u8>, prefix: &[u8], strict: bool, list: &AddedList) {
    let id = list.abstract_num_id.to_string();
    push_element(
        output,
        prefix,
        "abstractNum",
        &[("abstractNumId", &id)],
        false,
    );
    push_element(
        output,
        prefix,
        "multiLevelType",
        &[("val", "hybridMultilevel")],
        true,
    );
    for level in 0..LEVELS {
        let index = level.to_string();
        push_element(output, prefix, "lvl", &[("ilvl", &index)], false);
        let start = match list.start {
            Some((start_level, value)) if start_level == level => value.to_string(),
            _ => "1".to_owned(),
        };
        push_element(output, prefix, "start", &[("val", &start)], true);
        let text = match list.kind {
            ListKind::Bullet => {
                push_element(output, prefix, "numFmt", &[("val", "bullet")], true);
                BULLETS[usize::from(level) % BULLETS.len()].to_owned()
            },
            ListKind::Decimal => {
                let format = match level % 3 {
                    0 => "decimal",
                    1 => "lowerLetter",
                    _ => "lowerRoman",
                };
                push_element(output, prefix, "numFmt", &[("val", format)], true);
                format!("%{}.", level + 1)
            },
        };
        push_element(output, prefix, "lvlText", &[("val", &text)], true);
        push_element(
            output,
            prefix,
            "lvlJc",
            &[("val", if strict { "start" } else { "left" })],
            true,
        );
        push_element(output, prefix, "pPr", &[], false);
        let indent = (u32::from(level) + 1).saturating_mul(720).to_string();
        push_element(
            output,
            prefix,
            "ind",
            &[
                (if strict { "start" } else { "left" }, &indent),
                ("hanging", "360"),
            ],
            true,
        );
        push_end(output, prefix, "pPr");
        push_end(output, prefix, "lvl");
    }
    push_end(output, prefix, "abstractNum");
}

fn render_instance(output: &mut Vec<u8>, prefix: &[u8], list: &AddedList) {
    let id = list.num_id.to_string();
    let abstract_id = list.abstract_num_id.to_string();
    push_element(output, prefix, "num", &[("numId", &id)], false);
    push_element(
        output,
        prefix,
        "abstractNumId",
        &[("val", &abstract_id)],
        true,
    );
    push_end(output, prefix, "num");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn lists_are_spliced_in_schema_order_with_fresh_identifiers() {
        let xml = br#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="4"/><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num><w:numIdMacAtCleanup w:val="1"/></w:numbering>"#;
        assert_eq!(next_ids(xml).unwrap(), (5, 8));
        let list = AddedList {
            kind: ListKind::Decimal,
            abstract_num_id: 5,
            num_id: 8,
            start: Some((0, 3)),
        };
        let spliced = String::from_utf8(insert_lists(xml, &[list]).unwrap()).unwrap();
        let definition = spliced
            .find(r#"<w:abstractNum w:abstractNumId="5">"#)
            .unwrap();
        let instance = spliced.find(r#"<w:num w:numId="8">"#).unwrap();
        assert!(spliced.find(r#"<w:num w:numId="7">"#).unwrap() > definition);
        assert!(instance < spliced.find("<w:numIdMacAtCleanup").unwrap());
        assert!(spliced.contains(r#"<w:start w:val="3"/>"#));
        assert_eq!(next_ids(spliced.as_bytes()).unwrap(), (6, 9));

        let empty = empty_numbering_xml(false);
        assert_eq!(next_ids(&empty).unwrap(), (0, 1));
        let list = AddedList {
            kind: ListKind::Bullet,
            abstract_num_id: 0,
            num_id: 1,
            start: None,
        };
        let opened = String::from_utf8(insert_lists(&empty, &[list]).unwrap()).unwrap();
        assert!(opened.ends_with("</w:num></w:numbering>"), "{opened}");
        assert!(opened.contains("\u{2022}"));
    }
}
