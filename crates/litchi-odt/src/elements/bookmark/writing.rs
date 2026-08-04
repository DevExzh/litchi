use super::{Bookmark, BookmarkParser, BookmarkRange};
use crate::elements::xml::{TEXT_NAMESPACE, is_bound, namespaced_attribute};
use crate::expanded_attributes;
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;
use std::ops::Range;

const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 4_096;
const MAX_BOOKMARKS: usize = 1_000_000;
const MAX_NAME_BYTES: usize = 65_536;
const TEXT_NAMESPACE_STRING: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// A complete point or range bookmark target.
#[derive(Debug, Clone)]
pub enum BookmarkTarget {
    Point(Bookmark),
    Range(BookmarkRange),
}

impl BookmarkTarget {
    pub fn point(name: impl AsRef<str>) -> Self {
        Self::Point(Bookmark::new(name.as_ref()))
    }

    pub fn range(name: impl Into<String>) -> Self {
        Self::Range(BookmarkRange::new(name.into()))
    }

    pub fn name(&self) -> &str {
        match self {
            Self::Point(bookmark) => bookmark.name().expect("typed bookmark has text:name"),
            Self::Range(range) => &range.name,
        }
    }

    pub fn is_range(&self) -> bool {
        matches!(self, Self::Range(_))
    }

    pub fn to_xml_fragments(&self) -> Result<BookmarkFragments> {
        validate_name(self.name())?;
        let name = escape_attribute(self.name());
        Ok(match self {
            Self::Point(_) => {
                BookmarkFragments::Point(format!(r#"<text:bookmark text:name="{name}"/>"#))
            },
            Self::Range(_) => BookmarkFragments::Range {
                start: format!(r#"<text:bookmark-start text:name="{name}"/>"#),
                end: format!(r#"<text:bookmark-end text:name="{name}"/>"#),
            },
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BookmarkFragments {
    Point(String),
    Range { start: String, end: String },
}

pub fn parse_bookmark_targets(xml: &str) -> Result<Vec<BookmarkTarget>> {
    let scan = scan_locations(xml)?;
    let ranges = BookmarkParser::parse_bookmark_ranges(xml)?;
    Ok(scan
        .records
        .into_iter()
        .map(|record| match record.kind {
            TargetKind::Point => BookmarkTarget::point(record.name),
            TargetKind::Range => {
                let range = ranges
                    .iter()
                    .find(|range| range.name == record.name)
                    .cloned()
                    .unwrap_or_else(|| BookmarkRange::new(record.name));
                BookmarkTarget::Range(range)
            },
        })
        .collect())
}

pub fn insert_bookmark_xml(
    xml: &str,
    paragraph_index: usize,
    target: &BookmarkTarget,
) -> Result<String> {
    let scan = scan_locations(xml)?;
    ensure_unique(&scan.records, None, target.name())?;
    let paragraph = scan.paragraphs.get(paragraph_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "bookmark paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    let fragments = bind_if_needed(xml, target.to_xml_fragments()?);
    match (&paragraph.site, fragments) {
        (ParagraphSite::Paired { close_start, .. }, BookmarkFragments::Point(fragment)) => {
            apply_edits(xml, vec![(*close_start..*close_start, fragment)])
        },
        (
            ParagraphSite::Paired {
                open_end,
                close_start,
            },
            BookmarkFragments::Range { start, end },
        ) => apply_edits(
            xml,
            vec![
                (*close_start..*close_start, end),
                (*open_end..*open_end, start),
            ],
        ),
        (ParagraphSite::Empty { start, end, qname }, BookmarkFragments::Point(fragment)) => {
            expand_empty(xml, *start, *end, qname, &fragment)
        },
        (
            ParagraphSite::Empty { start, end, qname },
            BookmarkFragments::Range {
                start: open,
                end: close,
            },
        ) => expand_empty(xml, *start, *end, qname, &(open + &close)),
    }
}

pub fn replace_bookmark_xml(
    xml: &str,
    ordinal: usize,
    replacement: &BookmarkTarget,
) -> Result<String> {
    let scan = scan_locations(xml)?;
    let location = scan.locations.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat(format!("bookmark ordinal {ordinal} is out of bounds"))
    })?;
    ensure_unique(&scan.records, Some(ordinal), replacement.name())?;
    let fragments = bind_if_needed(xml, replacement.to_xml_fragments()?);
    match (location, fragments) {
        (TargetLocation::Point(span), BookmarkFragments::Point(fragment)) => {
            apply_edits(xml, vec![(span.clone(), fragment)])
        },
        (TargetLocation::Point(span), BookmarkFragments::Range { start, end }) => {
            apply_edits(xml, vec![(span.clone(), start + &end)])
        },
        (TargetLocation::Range { start, end }, BookmarkFragments::Point(fragment)) => apply_edits(
            xml,
            vec![(end.clone(), String::new()), (start.clone(), fragment)],
        ),
        (
            TargetLocation::Range { start, end },
            BookmarkFragments::Range {
                start: open,
                end: close,
            },
        ) => apply_edits(xml, vec![(end.clone(), close), (start.clone(), open)]),
    }
}

pub fn remove_bookmark_xml(xml: &str, ordinal: usize) -> Result<String> {
    let scan = scan_locations(xml)?;
    let location = scan.locations.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat(format!("bookmark ordinal {ordinal} is out of bounds"))
    })?;
    match location {
        TargetLocation::Point(span) => apply_edits(xml, vec![(span.clone(), String::new())]),
        TargetLocation::Range { start, end } => apply_edits(
            xml,
            vec![(end.clone(), String::new()), (start.clone(), String::new())],
        ),
    }
}

pub(super) fn validate_bookmark_xml(xml: &str) -> Result<()> {
    scan_locations(xml).map(|_| ())
}

#[derive(Clone)]
struct TargetRecord {
    name: String,
    kind: TargetKind,
}

#[derive(Clone, Copy)]
enum TargetKind {
    Point,
    Range,
}

enum TargetLocation {
    Point(Range<usize>),
    Range {
        start: Range<usize>,
        end: Range<usize>,
    },
}
struct ParagraphLocation {
    site: ParagraphSite,
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
struct Scan {
    paragraphs: Vec<ParagraphLocation>,
    records: Vec<TargetRecord>,
    locations: Vec<TargetLocation>,
}
struct Span {
    start: usize,
    end: usize,
}
enum MarkerKind {
    Point(usize),
    Start(String, usize),
    End(String),
}
struct OpenElement {
    local: Vec<u8>,
    start: usize,
    open_end: usize,
    paragraph: bool,
    marker: Option<MarkerKind>,
}

fn scan_locations(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML_BYTES {
        return invalid(format!("bookmark XML exceeds {MAX_XML_BYTES} bytes"));
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous_end = 0usize;
    let mut depth = 0usize;
    let mut open_elements = Vec::new();
    let mut open_ranges = HashMap::<String, usize>::new();
    let mut identities = HashMap::<String, usize>::new();
    let mut paragraphs = Vec::new();
    let mut records = Vec::new();
    let mut locations = Vec::<Option<TargetLocation>>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid bookmark XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        let span = Span {
            start: previous_end,
            end: event_end,
        };
        match event {
            Event::Start(ref element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("bookmark depth overflow".to_string()))?;
                if depth > MAX_DEPTH {
                    return invalid("bookmark nesting exceeds 4096 levels");
                }
                let local_name = element.local_name();
                let local = local_name.as_ref();
                let marker = if text_element {
                    start_marker(
                        &reader,
                        element,
                        local,
                        &span,
                        false,
                        &mut records,
                        &mut locations,
                        &mut identities,
                        &mut open_ranges,
                    )?
                } else {
                    None
                };
                open_elements.push(OpenElement {
                    local: local.to_vec(),
                    start: span.start,
                    open_end: span.end,
                    paragraph: text_element && matches!(local, b"p" | b"h"),
                    marker,
                });
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if text_element && matches!(local, b"p" | b"h") {
                    let qname = std::str::from_utf8(element.name().as_ref())
                        .map_err(|_| {
                            Error::InvalidFormat("non-UTF-8 bookmark paragraph name".to_string())
                        })?
                        .to_string();
                    paragraphs.push(ParagraphLocation {
                        site: ParagraphSite::Empty {
                            start: span.start,
                            end: span.end,
                            qname,
                        },
                    });
                }
                if text_element {
                    start_marker(
                        &reader,
                        element,
                        local,
                        &span,
                        true,
                        &mut records,
                        &mut locations,
                        &mut identities,
                        &mut open_ranges,
                    )?;
                }
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("bookmark XML stack underflow".to_string())
                })?;
                let open = open_elements.pop().ok_or_else(|| {
                    Error::InvalidFormat("bookmark XML stack mismatch".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("bookmark XML has mismatched elements");
                }
                if open.paragraph {
                    paragraphs.push(ParagraphLocation {
                        site: ParagraphSite::Paired {
                            open_end: open.open_end,
                            close_start: span.start,
                        },
                    });
                }
                finish_marker(open, &span, &mut locations, &mut open_ranges)?;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in bookmark XML"),
            Event::Eof => break,
            _ => {},
        }
        previous_end = event_end;
        buffer.clear();
    }
    if depth != 0 || !open_elements.is_empty() || !open_ranges.is_empty() {
        return invalid("incomplete bookmark XML");
    }
    paragraphs.sort_by_key(|paragraph| match paragraph.site {
        ParagraphSite::Paired { open_end, .. } => open_end,
        ParagraphSite::Empty { start, .. } => start,
    });
    let locations = locations
        .into_iter()
        .map(|value| {
            value.ok_or_else(|| Error::InvalidFormat("incomplete bookmark range".to_string()))
        })
        .collect::<Result<_>>()?;
    Ok(Scan {
        paragraphs,
        records,
        locations,
    })
}

#[allow(clippy::too_many_arguments)]
fn start_marker(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    span: &Span,
    empty: bool,
    records: &mut Vec<TargetRecord>,
    locations: &mut Vec<Option<TargetLocation>>,
    identities: &mut HashMap<String, usize>,
    open_ranges: &mut HashMap<String, usize>,
) -> Result<Option<MarkerKind>> {
    let kind = match local {
        b"bookmark" => 0,
        b"bookmark-start" => 1,
        b"bookmark-end" => 2,
        _ => return Ok(None),
    };
    validate_attributes(reader, element)?;
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "bookmark")?
        .ok_or_else(|| Error::InvalidFormat("bookmark requires text:name".to_string()))?;
    validate_name(&name)?;
    if kind != 2 {
        if records.len() >= MAX_BOOKMARKS {
            return invalid("too many bookmarks");
        }
        if identities.insert(name.clone(), records.len()).is_some() {
            return invalid(format!("duplicate bookmark identity '{name}'"));
        }
        let order = records.len();
        records.push(TargetRecord {
            name: name.clone(),
            kind: if kind == 0 {
                TargetKind::Point
            } else {
                TargetKind::Range
            },
        });
        if kind == 0 {
            locations.push(if empty {
                Some(TargetLocation::Point(span.start..span.end))
            } else {
                None
            });
            return Ok(if empty {
                None
            } else {
                Some(MarkerKind::Point(order))
            });
        }
        locations.push(if empty {
            Some(TargetLocation::Range {
                start: span.start..span.end,
                end: 0..0,
            })
        } else {
            None
        });
        if open_ranges.insert(name.clone(), order).is_some() {
            return invalid(format!("duplicate open bookmark range '{name}'"));
        }
        return Ok(if empty {
            None
        } else {
            Some(MarkerKind::Start(name, order))
        });
    }
    if empty {
        close_range(&name, span.start..span.end, locations, open_ranges)?;
        Ok(None)
    } else {
        Ok(Some(MarkerKind::End(name)))
    }
}

fn validate_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attributes = expanded_attributes(reader, element, "bookmark")?;
    if attributes.len() != 1
        || attributes[0].namespace_uri.as_deref() != Some(TEXT_NAMESPACE_STRING)
        || attributes[0].local_name != "name"
    {
        return invalid("bookmarks allow only text:name");
    }
    Ok(())
}

fn finish_marker(
    open: OpenElement,
    closing: &Span,
    locations: &mut [Option<TargetLocation>],
    open_ranges: &mut HashMap<String, usize>,
) -> Result<()> {
    let Some(marker) = open.marker else {
        return Ok(());
    };
    if open.open_end != closing.start {
        return invalid("bookmark marker elements must be empty");
    }
    let span = open.start..closing.end;
    match marker {
        MarkerKind::Point(order) => locations[order] = Some(TargetLocation::Point(span)),
        MarkerKind::Start(name, order) => {
            if open_ranges.get(&name).copied() != Some(order) {
                return invalid("bookmark range identity changed during scan");
            }
            locations[order] = Some(TargetLocation::Range {
                start: span,
                end: 0..0,
            });
        },
        MarkerKind::End(name) => close_range(&name, span, locations, open_ranges)?,
    }
    Ok(())
}

fn close_range(
    name: &str,
    end: Range<usize>,
    locations: &mut [Option<TargetLocation>],
    open_ranges: &mut HashMap<String, usize>,
) -> Result<()> {
    let order = open_ranges
        .remove(name)
        .ok_or_else(|| Error::InvalidFormat(format!("bookmark-end has no open range '{name}'")))?;
    match locations.get_mut(order).and_then(Option::as_mut) {
        Some(TargetLocation::Range { end: target, .. }) => *target = end,
        _ => return invalid("bookmark range start is incomplete"),
    }
    Ok(())
}

fn validate_name(name: &str) -> Result<()> {
    if name.len() > MAX_NAME_BYTES {
        return invalid(format!("bookmark name exceeds {MAX_NAME_BYTES} bytes"));
    }
    if name.chars().any(|character| !matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) { return invalid("bookmark name contains a forbidden XML character"); }
    Ok(())
}

fn ensure_unique(records: &[TargetRecord], except: Option<usize>, name: &str) -> Result<()> {
    validate_name(name)?;
    if records
        .iter()
        .enumerate()
        .any(|(index, record)| Some(index) != except && record.name == name)
    {
        return invalid(format!("duplicate bookmark identity '{name}'"));
    }
    Ok(())
}

fn escape_attribute(value: &str) -> String {
    let mut output = String::new();
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            _ => output.push(character),
        }
    }
    output
}
fn bind_if_needed(xml: &str, fragments: BookmarkFragments) -> BookmarkFragments {
    if xml.contains(&format!(r#"xmlns:text="{TEXT_NAMESPACE_STRING}""#))
        || xml.contains(&format!("xmlns:text='{TEXT_NAMESPACE_STRING}'"))
    {
        return fragments;
    }
    let bind = |fragment: String| {
        fragment.replacen(
            " text:name=",
            &format!(r#" xmlns:text="{TEXT_NAMESPACE_STRING}" text:name="#),
            1,
        )
    };
    match fragments {
        BookmarkFragments::Point(fragment) => BookmarkFragments::Point(bind(fragment)),
        BookmarkFragments::Range { start, end } => BookmarkFragments::Range {
            start: bind(start),
            end: bind(end),
        },
    }
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid empty bookmark paragraph span".to_string()))?;
    let open = source.strip_suffix("/>").ok_or_else(|| {
        Error::InvalidFormat("empty bookmark paragraph does not end with />".to_string())
    })?;
    apply_edits(
        xml,
        vec![(start..end, format!("{open}>{content}</{qname}>"))],
    )
}
fn apply_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0.start));
    let mut output = xml.to_string();
    let mut previous = xml.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > previous || span.end > output.len() {
            return invalid("overlapping or invalid bookmark mutation spans");
        }
        output.replace_range(span.clone(), &replacement);
        previous = span.start;
    }
    Ok(output)
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;
    const TEXT: &str = TEXT_NAMESPACE_STRING;
    fn wrapped(body: &str) -> String {
        format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:text="{TEXT}">{body}</o:text>"#
        )
    }

    #[test]
    fn canonical_fragments_and_lossless_mutations_round_trip() {
        let point = BookmarkTarget::point("p<&\"");
        assert_eq!(
            point.to_xml_fragments().unwrap(),
            BookmarkFragments::Point(
                r#"<text:bookmark text:name="p&lt;&amp;&quot;"/>"#.to_string()
            )
        );
        let source =
            wrapped(r#"<t:p t:style-name="P">A&amp;<t:span>B</t:span><!--keep--></t:p><t:p/>"#);
        let ranged = insert_bookmark_xml(&source, 0, &BookmarkTarget::range("r")).unwrap();
        assert!(ranged.contains(r#"<text:bookmark-start text:name="r"/>A&amp;<t:span>B</t:span><!--keep--><text:bookmark-end text:name="r"/>"#));
        let point = replace_bookmark_xml(&ranged, 0, &BookmarkTarget::point("p")).unwrap();
        let removed = remove_bookmark_xml(&point, 0).unwrap();
        assert_eq!(removed, source);
        assert!(
            insert_bookmark_xml(&source, 1, &BookmarkTarget::point("empty"))
                .unwrap()
                .contains(r#"<t:p><text:bookmark text:name="empty"/></t:p>"#)
        );
    }

    #[test]
    fn hostile_identity_namespace_content_and_resources_are_rejected() {
        for body in [
            r#"<t:p><t:bookmark t:name="x" u:extra="1"/></t:p>"#,
            r#"<t:p><t:bookmark u:name="x"/></t:p>"#,
            r#"<t:p><t:bookmark t:name="x"/><t:bookmark-start t:name="x"/><t:bookmark-end t:name="x"/></t:p>"#,
            r#"<t:p><t:bookmark-start t:name="x"/></t:p>"#,
            r#"<t:p><t:bookmark t:name="x">bad</t:bookmark></t:p>"#,
            r#"<!DOCTYPE x><t:p/>"#,
        ] {
            let xml = format!(
                r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:u="urn:hostile">{body}</o:text>"#
            );
            assert!(parse_bookmark_targets(&xml).is_err(), "accepted {body}");
        }
        assert!(
            BookmarkTarget::point("x".repeat(MAX_NAME_BYTES + 1))
                .to_xml_fragments()
                .is_err()
        );
        assert!(
            BookmarkTarget::point("bad\0name")
                .to_xml_fragments()
                .is_err()
        );
    }

    #[test]
    fn libreoffice_odfpy_odfdo_and_crossed_ranges_round_trip() {
        let libreoffice = include_str!(
            "../../../../../test-data/libreoffice-core/sw/qa/extras/odfexport/data/CrossRefHeadingBookmark.fodt"
        );
        let targets = parse_bookmark_targets(libreoffice).unwrap();
        assert_eq!(targets.len(), 2);
        assert!(targets.iter().all(BookmarkTarget::is_range));
        let replaced =
            replace_bookmark_xml(libreoffice, 1, &BookmarkTarget::range("odfdo-range")).unwrap();
        assert!(replaced.contains(r#"<text:bookmark-ref text:reference-format="number" text:ref-name="__RefHeading__1673_25705824">1.1</text:bookmark-ref>"#));
        assert_eq!(parse_bookmark_targets(&replaced).unwrap().len(), 2);
        let odfpy = wrapped(
            r#"<t:h><t:bookmark-start t:name="ClassName"/>ClassName<t:bookmark-end t:name="ClassName"/></t:h>"#,
        );
        assert_eq!(parse_bookmark_targets(&odfpy).unwrap().len(), 1);
    }

    #[test]
    fn builder_and_mutable_document_round_trip_targets_without_evaluation() {
        let mut builder = crate::Builder::new();
        builder.add_paragraph("payload").unwrap();
        builder
            .add_bookmark_target(0, &BookmarkTarget::range("built"))
            .unwrap();
        let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
        let targets = document.bookmark_ranges().unwrap();
        assert_eq!(targets.len(), 1);
        assert_eq!(targets[0].name, "built");

        let mut mutable = crate::MutableDocument::from_document(document).unwrap();
        let old = mutable
            .replace_bookmark_target(0, &BookmarkTarget::point("changed"))
            .unwrap();
        assert!(old.is_range());
        mutable
            .insert_bookmark_target(0, &BookmarkTarget::point("second"))
            .unwrap();
        assert_eq!(mutable.bookmark_targets().unwrap().len(), 2);
        let removed = mutable.remove_bookmark_target(0).unwrap();
        assert_eq!(removed.name(), "changed");
    }
}
