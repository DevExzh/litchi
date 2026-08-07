//! Canonical index-mark fragments and byte-preserving inline mutation.

use super::MAX_MARKS;
use super::{
    TextAlphabeticalMarkMetadata, TextIndexMark, TextIndexMarkFragments, TextIndexMarkKind,
    parse_text_index_marks,
};
use crate::elements::xml::{TEXT_NAMESPACE, is_bound, namespaced_attribute};
use crate::index::TextIndexAttribute;
use crate::{TextBibliographyType, bibliography_configuration::Field};
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

pub(super) const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_FRAGMENT_BYTES: usize = 4 * 1024 * 1024;
const MAX_DEPTH: usize = 4_096;

impl TextIndexMark {
    pub fn toc_point(value: impl Into<String>, outline_level: Option<u16>) -> Result<Self> {
        point_mark(
            TextIndexMarkKind::TableOfContents,
            value.into(),
            toc_attributes(outline_level)?,
        )
    }

    pub fn toc_range(id: impl Into<String>, outline_level: Option<u16>) -> Result<Self> {
        range_mark(
            TextIndexMarkKind::TableOfContents,
            id.into(),
            toc_attributes(outline_level)?,
        )
    }

    pub fn user_point(
        value: impl Into<String>,
        index_name: impl Into<String>,
        outline_level: Option<u16>,
    ) -> Result<Self> {
        point_mark(
            TextIndexMarkKind::User,
            value.into(),
            user_attributes(index_name.into(), outline_level)?,
        )
    }

    pub fn user_range(
        id: impl Into<String>,
        index_name: impl Into<String>,
        outline_level: Option<u16>,
    ) -> Result<Self> {
        range_mark(
            TextIndexMarkKind::User,
            id.into(),
            user_attributes(index_name.into(), outline_level)?,
        )
    }

    pub fn alphabetical_point(
        value: impl Into<String>,
        metadata: TextAlphabeticalMarkMetadata,
    ) -> Result<Self> {
        point_mark(
            TextIndexMarkKind::Alphabetical,
            value.into(),
            alphabetical_attributes(metadata),
        )
    }

    pub fn alphabetical_range(
        id: impl Into<String>,
        metadata: TextAlphabeticalMarkMetadata,
    ) -> Result<Self> {
        range_mark(
            TextIndexMarkKind::Alphabetical,
            id.into(),
            alphabetical_attributes(metadata),
        )
    }

    pub fn bibliography_point(
        bibliography_type: TextBibliographyType,
        visible_text: impl Into<String>,
        fields: Vec<(Field, String)>,
    ) -> Result<Self> {
        if fields.len() > 64 {
            return invalid("bibliography mark has too many fields");
        }
        let mut names = HashSet::new();
        let mut attributes = vec![attribute("bibliography-type", bibliography_type.as_str())];
        for (field, value) in fields {
            if !names.insert(field.as_str()) {
                return invalid(format!("duplicate bibliography field {}", field.as_str()));
            }
            checked_string(&value, "bibliography field")?;
            attributes.push(attribute(field.as_str(), value));
        }
        point_mark(
            TextIndexMarkKind::Bibliography,
            visible_text.into(),
            attributes,
        )
    }

    pub fn to_xml_fragments(&self) -> Result<TextIndexMarkFragments> {
        validate_mark(self)?;
        if self.range {
            let id = self
                .id
                .as_deref()
                .ok_or_else(|| Error::InvalidFormat("range mark has no ID".to_string()))?;
            let mut start_attributes = self.attributes.clone();
            set_attribute(&mut start_attributes, "id", id.to_string());
            let start = empty_fragment(start_name(self.kind)?, &start_attributes);
            let end = empty_fragment(end_name(self.kind)?, &[attribute("id", id)]);
            if start.len() + end.len() > MAX_FRAGMENT_BYTES {
                return invalid("index mark fragments exceed 4 MiB");
            }
            Ok(TextIndexMarkFragments::Range { start, end })
        } else {
            let mut attributes = self.attributes.clone();
            if self.kind == TextIndexMarkKind::Bibliography {
                let mut fragment = start_fragment("bibliography-mark", &attributes);
                escape_text(&self.value, &mut fragment);
                fragment.push_str("</text:bibliography-mark>");
                if fragment.len() > MAX_FRAGMENT_BYTES {
                    return invalid("bibliography mark fragment exceeds 4 MiB");
                }
                Ok(TextIndexMarkFragments::Point(fragment))
            } else {
                set_attribute(&mut attributes, "string-value", self.value.clone());
                let fragment = empty_fragment(point_name(self.kind), &attributes);
                if fragment.len() > MAX_FRAGMENT_BYTES {
                    return invalid("index mark fragment exceeds 4 MiB");
                }
                Ok(TextIndexMarkFragments::Point(fragment))
            }
        }
    }
}

pub fn insert_text_index_mark_xml(
    xml: &str,
    paragraph_index: usize,
    mark: &TextIndexMark,
) -> Result<String> {
    validated_marks(xml)?;
    let fragments = mark.to_xml_fragments()?;
    let scan = scan_xml(xml)?;
    let paragraph = scan.paragraphs.get(paragraph_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    let output = match (&paragraph.site, fragments) {
        (
            ParagraphSite::Paired {
                open_end: _,
                close_start,
            },
            TextIndexMarkFragments::Point(fragment),
        ) => splice_one(xml, *close_start, *close_start, &fragment),
        (
            ParagraphSite::Paired {
                open_end,
                close_start,
            },
            TextIndexMarkFragments::Range { start, end },
        ) => {
            let mut output = String::with_capacity(xml.len() + start.len() + end.len());
            output.push_str(&xml[..*open_end]);
            output.push_str(&start);
            output.push_str(&xml[*open_end..*close_start]);
            output.push_str(&end);
            output.push_str(&xml[*close_start..]);
            output
        },
        (ParagraphSite::Empty { start, end, qname }, TextIndexMarkFragments::Point(fragment)) => {
            expand_empty(xml, *start, *end, qname, &fragment)
        },
        (
            ParagraphSite::Empty { start, end, qname },
            TextIndexMarkFragments::Range {
                start: range_start,
                end: range_end,
            },
        ) => expand_empty(xml, *start, *end, qname, &(range_start + &range_end)),
    };
    validated_marks(&output)?;
    Ok(output)
}

pub fn replace_text_index_mark_xml(
    xml: &str,
    mark_index: usize,
    replacement: &TextIndexMark,
) -> Result<String> {
    let current = validated_marks(xml)?;
    let old = current
        .get(mark_index)
        .ok_or_else(|| Error::InvalidFormat(format!("index mark {mark_index} is out of bounds")))?;
    if old.is_range() != replacement.is_range() {
        return invalid("point and range index marks cannot replace each other");
    }
    let scan = scan_xml(xml)?;
    let location = scan.marks.get(mark_index).ok_or_else(|| {
        Error::InvalidFormat("index mark parser/scanner order mismatch".to_string())
    })?;
    let fragments = replacement.to_xml_fragments()?;
    let output = match (location, fragments) {
        (MarkLocation::Point { span, .. }, TextIndexMarkFragments::Point(fragment)) => {
            splice_one(xml, span.start, span.end, &fragment)
        },
        (
            MarkLocation::Range { start, end },
            TextIndexMarkFragments::Range {
                start: start_fragment,
                end: end_fragment,
            },
        ) => {
            let mut output = String::with_capacity(
                xml.len() - (start.len() + end.len()) + start_fragment.len() + end_fragment.len(),
            );
            output.push_str(&xml[..start.start]);
            output.push_str(&start_fragment);
            output.push_str(&xml[start.end..end.start]);
            output.push_str(&end_fragment);
            output.push_str(&xml[end.end..]);
            output
        },
        _ => return invalid("index mark parser/scanner shape mismatch"),
    };
    validated_marks(&output)?;
    Ok(output)
}

pub fn remove_text_index_mark_xml(xml: &str, mark_index: usize) -> Result<String> {
    validated_marks(xml)?
        .get(mark_index)
        .ok_or_else(|| Error::InvalidFormat(format!("index mark {mark_index} is out of bounds")))?;
    let scan = scan_xml(xml)?;
    let location = scan.marks.get(mark_index).ok_or_else(|| {
        Error::InvalidFormat("index mark parser/scanner order mismatch".to_string())
    })?;
    let output = match location {
        MarkLocation::Point { span, inner } => {
            let replacement = inner.map_or("", |span| &xml[span.start..span.end]);
            splice_one(xml, span.start, span.end, replacement)
        },
        MarkLocation::Range { start, end } => {
            let mut output = String::with_capacity(xml.len() - start.len() - end.len());
            output.push_str(&xml[..start.start]);
            output.push_str(&xml[start.end..end.start]);
            output.push_str(&xml[end.end..]);
            output
        },
    };
    validated_marks(&output)?;
    Ok(output)
}

fn point_mark(
    kind: TextIndexMarkKind,
    value: String,
    attributes: Vec<TextIndexAttribute>,
) -> Result<TextIndexMark> {
    checked_string(&value, "index mark value")?;
    let mark = TextIndexMark {
        kind,
        id: None,
        value,
        range: false,
        attributes,
    };
    validate_mark(&mark)?;
    Ok(mark)
}

fn range_mark(
    kind: TextIndexMarkKind,
    id: String,
    attributes: Vec<TextIndexAttribute>,
) -> Result<TextIndexMark> {
    if kind == TextIndexMarkKind::Bibliography {
        return invalid("bibliography marks cannot be ranges");
    }
    required(&id, "index range ID")?;
    checked_string(&id, "index range ID")?;
    let mark = TextIndexMark {
        kind,
        id: Some(id),
        value: String::new(),
        range: true,
        attributes,
    };
    validate_mark(&mark)?;
    Ok(mark)
}

fn toc_attributes(outline_level: Option<u16>) -> Result<Vec<TextIndexAttribute>> {
    let mut attributes = Vec::new();
    optional_outline(&mut attributes, outline_level)?;
    Ok(attributes)
}

fn user_attributes(
    index_name: String,
    outline_level: Option<u16>,
) -> Result<Vec<TextIndexAttribute>> {
    checked_string(&index_name, "user index name")?;
    let mut attributes = vec![attribute("index-name", index_name)];
    optional_outline(&mut attributes, outline_level)?;
    Ok(attributes)
}

fn alphabetical_attributes(metadata: TextAlphabeticalMarkMetadata) -> Vec<TextIndexAttribute> {
    let mut attributes = Vec::new();
    for (name, value) in [
        ("key1", metadata.key1),
        ("key2", metadata.key2),
        ("string-value-phonetic", metadata.string_value_phonetic),
        ("key1-phonetic", metadata.key1_phonetic),
        ("key2-phonetic", metadata.key2_phonetic),
    ] {
        if let Some(value) = value {
            attributes.push(attribute(name, value));
        }
    }
    if let Some(value) = metadata.main_entry {
        attributes.push(attribute("main-entry", value.to_string()));
    }
    attributes
}

fn optional_outline(attributes: &mut Vec<TextIndexAttribute>, value: Option<u16>) -> Result<()> {
    if let Some(value) = value {
        if value == 0 {
            return invalid("index mark outline level must be positive");
        }
        attributes.push(attribute("outline-level", value.to_string()));
    }
    Ok(())
}

fn validate_mark(mark: &TextIndexMark) -> Result<()> {
    checked_string(&mark.value, "index mark value")?;
    if mark.range && mark.kind == TextIndexMarkKind::Bibliography {
        return invalid("bibliography marks cannot be ranges");
    }
    let allowed: &[&str] = match mark.kind {
        TextIndexMarkKind::TableOfContents => &["outline-level", "string-value", "id"],
        TextIndexMarkKind::User => &["index-name", "outline-level", "string-value", "id"],
        TextIndexMarkKind::Alphabetical => &[
            "key1",
            "key2",
            "string-value-phonetic",
            "key1-phonetic",
            "key2-phonetic",
            "main-entry",
            "string-value",
            "id",
        ],
        TextIndexMarkKind::Bibliography => &[
            "bibliography-type",
            "identifier",
            "address",
            "annote",
            "author",
            "booktitle",
            "chapter",
            "edition",
            "editor",
            "howpublished",
            "institution",
            "journal",
            "month",
            "note",
            "number",
            "organizations",
            "pages",
            "publisher",
            "school",
            "series",
            "title",
            "report-type",
            "volume",
            "year",
            "url",
            "custom1",
            "custom2",
            "custom3",
            "custom4",
            "custom5",
            "isbn",
            "issn",
        ],
    };
    let mut names = HashSet::new();
    for attribute in &mark.attributes {
        if attribute.namespace_uri.as_deref() != Some(TEXT)
            || !allowed.contains(&attribute.local_name.as_str())
        {
            return invalid(format!(
                "unexpected index mark attribute {}",
                attribute.local_name
            ));
        }
        if !names.insert(attribute.local_name.as_str()) {
            return invalid("duplicate index mark attribute");
        }
        checked_string(&attribute.value, "index mark attribute")?;
    }
    if mark.kind == TextIndexMarkKind::User && mark.attribute(Some(TEXT), "index-name").is_none() {
        return invalid("user index mark requires text:index-name");
    }
    if let Some(level) = mark.attribute(Some(TEXT), "outline-level")
        && level
            .parse::<u64>()
            .ok()
            .as_ref()
            .is_none_or(|value| *value <= 0)
    {
        return invalid("index mark outline level must be positive");
    }
    if let Some(value) = mark.attribute(Some(TEXT), "main-entry")
        && !matches!(value, "true" | "false" | "1" | "0")
    {
        return invalid("text:main-entry is not an XML boolean");
    }
    if mark.kind == TextIndexMarkKind::Bibliography {
        let kind = mark
            .attribute(Some(TEXT), "bibliography-type")
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "bibliography mark requires text:bibliography-type".to_string(),
                )
            })?;
        if !super::is_bibliography_type(kind) {
            return invalid("invalid bibliography mark type");
        }
    }
    if mark.range {
        required(mark.id.as_deref().unwrap_or(""), "index range ID")?;
    }
    Ok(())
}

pub(super) fn validated_marks(xml: &str) -> Result<Vec<TextIndexMark>> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("index-mark XML exceeds 64 MiB");
    }
    let marks = parse_text_index_marks(xml)?;
    for mark in &marks {
        validate_mark(mark)?;
    }
    Ok(marks)
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}
impl Span {
    fn len(self) -> usize {
        self.end - self.start
    }
}

enum ParagraphSite {
    Paired {
        open_end: usize,
        close_start: usize,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
    },
}
struct ParagraphLocation {
    site: ParagraphSite,
}
enum MarkLocation {
    Point { span: Span, inner: Option<Span> },
    Range { start: Span, end: Span },
}
struct Scan {
    paragraphs: Vec<ParagraphLocation>,
    marks: Vec<MarkLocation>,
}
struct OpenElement {
    local: Vec<u8>,
    depth: usize,
    start: usize,
    open_end: usize,
    order: Option<usize>,
    key: Option<(TextIndexMarkKind, String)>,
}

fn scan_xml(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("index-mark XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut paragraphs = Vec::new();
    let mut open_paragraphs = Vec::<(usize, usize)>::new();
    let mut marks: Vec<Option<MarkLocation>> = Vec::new();
    let mut open_ranges = HashMap::<(TextIndexMarkKind, String), usize>::new();
    let mut open_elements = Vec::<OpenElement>::new();
    let mut mark_count = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid index-mark XML while scanning: {error}"))
            })?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        drop(namespace);
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                if text_element && matches!(local.as_slice(), b"p" | b"h") {
                    open_paragraphs.push((depth, end));
                }
                let (order, key) = if text_element {
                    start_mark_element(
                        &reader,
                        element,
                        &local,
                        Span { start, end },
                        &mut marks,
                        &mut open_ranges,
                        &mut mark_count,
                        false,
                    )?
                } else {
                    (None, None)
                };
                open_elements.push(OpenElement {
                    local,
                    depth,
                    start,
                    open_end: end,
                    order,
                    key,
                });
                depth += 1;
                if depth > MAX_DEPTH {
                    return invalid("index-mark XML nesting exceeds 4096");
                }
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if text_element && matches!(local, b"p" | b"h") {
                    let qname = std::str::from_utf8(element.name().as_ref())
                        .map_err(|_| Error::InvalidFormat("non-UTF-8 paragraph name".to_string()))?
                        .to_string();
                    paragraphs.push(ParagraphLocation {
                        site: ParagraphSite::Empty { start, end, qname },
                    });
                }
                if text_element {
                    start_mark_element(
                        &reader,
                        element,
                        local,
                        Span { start, end },
                        &mut marks,
                        &mut open_ranges,
                        &mut mark_count,
                        true,
                    )?;
                }
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("index-mark XML stack underflow".to_string())
                })?;
                let open = open_elements.pop().ok_or_else(|| {
                    Error::InvalidFormat("index-mark XML stack mismatch".to_string())
                })?;
                if open.depth != depth || open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("index-mark XML has mismatched elements");
                }
                if text_element && matches!(open.local.as_slice(), b"p" | b"h") {
                    let (_, open_end) = open_paragraphs.pop().ok_or_else(|| {
                        Error::InvalidFormat("paragraph stack mismatch".to_string())
                    })?;
                    paragraphs.push(ParagraphLocation {
                        site: ParagraphSite::Paired {
                            open_end,
                            close_start: start,
                        },
                    });
                }
                finish_mark_element(open, Span { start, end }, &mut marks, &mut open_ranges)?;
            },
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in mutable index-mark XML");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || !open_ranges.is_empty() {
        return invalid("incomplete index-mark XML scan");
    }
    paragraphs.sort_by_key(|paragraph| match paragraph.site {
        ParagraphSite::Paired { open_end, .. } => open_end,
        ParagraphSite::Empty { start, .. } => start,
    });
    Ok(Scan {
        paragraphs,
        marks: marks
            .into_iter()
            .map(|mark| {
                mark.ok_or_else(|| Error::InvalidFormat("incomplete range mark scan".to_string()))
            })
            .collect::<Result<_>>()?,
    })
}

#[allow(clippy::too_many_arguments)]
#[allow(clippy::type_complexity)]
fn start_mark_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    span: Span,
    marks: &mut Vec<Option<MarkLocation>>,
    open_ranges: &mut HashMap<(TextIndexMarkKind, String), usize>,
    mark_count: &mut usize,
    empty: bool,
) -> Result<(Option<usize>, Option<(TextIndexMarkKind, String)>)> {
    if let Some(kind) = super::start_kind(local) {
        if *mark_count >= MAX_MARKS {
            return invalid("too many index marks");
        }
        let id = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"id", "index range")?
            .ok_or_else(|| Error::InvalidFormat("index range start has no ID".to_string()))?;
        let key = (kind, id);
        let order = *mark_count;
        *mark_count += 1;
        marks.push(None);
        if open_ranges.insert(key.clone(), order).is_some() {
            return invalid("duplicate open index range");
        }
        if empty {
            marks[order] = Some(MarkLocation::Range {
                start: span,
                end: Span { start: 0, end: 0 },
            });
            return Ok((None, Some(key)));
        }
        return Ok((Some(order), Some(key)));
    }
    if let Some(kind) = super::end_kind(local) {
        let id = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"id", "index range")?
            .ok_or_else(|| Error::InvalidFormat("index range end has no ID".to_string()))?;
        let key = (kind, id);
        if !empty {
            return Ok((None, Some(key)));
        }
        close_range_location(key, span, marks, open_ranges)?;
        return Ok((None, None));
    }
    if super::point_kind(local).is_some() {
        if *mark_count >= MAX_MARKS {
            return invalid("too many index marks");
        }
        let order = *mark_count;
        *mark_count += 1;
        marks.push(None);
        if empty {
            marks[order] = Some(MarkLocation::Point { span, inner: None });
            return Ok((None, None));
        }
        return Ok((Some(order), None));
    }
    Ok((None, None))
}

fn finish_mark_element(
    open: OpenElement,
    closing: Span,
    marks: &mut [Option<MarkLocation>],
    open_ranges: &mut HashMap<(TextIndexMarkKind, String), usize>,
) -> Result<()> {
    if let Some(key) = open.key {
        if super::start_kind(&open.local).is_some() {
            let order = open_ranges
                .get(&key)
                .copied()
                .ok_or_else(|| Error::InvalidFormat("missing open range".to_string()))?;
            marks[order] = Some(MarkLocation::Range {
                start: Span {
                    start: open.start,
                    end: closing.end,
                },
                end: Span { start: 0, end: 0 },
            });
        } else if super::end_kind(&open.local).is_some() {
            close_range_location(
                key,
                Span {
                    start: open.start,
                    end: closing.end,
                },
                marks,
                open_ranges,
            )?;
        }
    } else if let Some(order) = open.order {
        if super::point_kind(&open.local) != Some(TextIndexMarkKind::Bibliography)
            && open.open_end != closing.start
        {
            return invalid("non-bibliography index marks cannot contain content");
        }
        marks[order] = Some(MarkLocation::Point {
            span: Span {
                start: open.start,
                end: closing.end,
            },
            inner: Some(Span {
                start: open.open_end,
                end: closing.start,
            }),
        });
    }
    Ok(())
}

fn close_range_location(
    key: (TextIndexMarkKind, String),
    end: Span,
    marks: &mut [Option<MarkLocation>],
    open_ranges: &mut HashMap<(TextIndexMarkKind, String), usize>,
) -> Result<()> {
    let order = open_ranges
        .remove(&key)
        .ok_or_else(|| Error::InvalidFormat("index range end has no start".to_string()))?;
    match marks[order].take() {
        Some(MarkLocation::Range { start, .. }) => {
            marks[order] = Some(MarkLocation::Range { start, end });
        },
        None => return invalid("range start marker scan is incomplete"),
        _ => return invalid("range marker scan shape mismatch"),
    }
    Ok(())
}

fn point_name(kind: TextIndexMarkKind) -> &'static str {
    match kind {
        TextIndexMarkKind::TableOfContents => "toc-mark",
        TextIndexMarkKind::User => "user-index-mark",
        TextIndexMarkKind::Alphabetical => "alphabetical-index-mark",
        TextIndexMarkKind::Bibliography => "bibliography-mark",
    }
}
fn start_name(kind: TextIndexMarkKind) -> Result<&'static str> {
    match kind {
        TextIndexMarkKind::TableOfContents => Ok("toc-mark-start"),
        TextIndexMarkKind::User => Ok("user-index-mark-start"),
        TextIndexMarkKind::Alphabetical => Ok("alphabetical-index-mark-start"),
        TextIndexMarkKind::Bibliography => invalid("bibliography marks cannot be ranges"),
    }
}
fn end_name(kind: TextIndexMarkKind) -> Result<&'static str> {
    match kind {
        TextIndexMarkKind::TableOfContents => Ok("toc-mark-end"),
        TextIndexMarkKind::User => Ok("user-index-mark-end"),
        TextIndexMarkKind::Alphabetical => Ok("alphabetical-index-mark-end"),
        TextIndexMarkKind::Bibliography => invalid("bibliography marks cannot be ranges"),
    }
}

fn attribute(local: &str, value: impl Into<String>) -> TextIndexAttribute {
    TextIndexAttribute {
        namespace_uri: Some(TEXT.to_string()),
        local_name: local.to_string(),
        value: value.into(),
    }
}
fn set_attribute(attributes: &mut Vec<TextIndexAttribute>, local: &str, value: String) {
    if let Some(attribute) = attributes.iter_mut().find(|attribute| {
        attribute.namespace_uri.as_deref() == Some(TEXT) && attribute.local_name == local
    }) {
        attribute.value = value;
    } else {
        attributes.push(attribute(local, value));
    }
}
fn required(value: &str, context: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{context} cannot be empty"))
    } else {
        Ok(())
    }
}
fn checked_string(value: &str, context: &str) -> Result<()> {
    if value.len() > MAX_FRAGMENT_BYTES {
        invalid(format!("{context} exceeds 4 MiB"))
    } else {
        Ok(())
    }
}

fn empty_fragment(name: &str, attributes: &[TextIndexAttribute]) -> String {
    let mut output = start_fragment(name, attributes);
    output.truncate(output.len() - 1);
    output.push_str("/>");
    output
}
fn start_fragment(name: &str, attributes: &[TextIndexAttribute]) -> String {
    let mut output = format!("<text:{name} xmlns:text=\"{TEXT}\"");
    let mut attributes: Vec<_> = attributes.iter().collect();
    attributes.sort_by_key(|attribute| attribute.local_name.as_str());
    for attribute in attributes {
        output.push_str(" text:");
        output.push_str(&attribute.local_name);
        output.push_str("=\"");
        escape_attr(&attribute.value, &mut output);
        output.push('"');
    }
    output.push('>');
    output
}
fn escape_text(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}
fn escape_attr(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}
fn splice_one(xml: &str, start: usize, end: usize, replacement: &str) -> String {
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(replacement);
    output.push_str(&xml[end..]);
    output
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> String {
    let raw = &xml[start..end];
    let slash = raw.rfind("/>").expect("quick-xml empty element");
    let mut output = String::with_capacity(xml.len() + content.len() + qname.len() + 3);
    output.push_str(&xml[..start]);
    output.push_str(&raw[..slash]);
    output.push('>');
    output.push_str(content);
    output.push_str("</");
    output.push_str(qname);
    output.push('>');
    output.push_str(&xml[end..]);
    output
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
