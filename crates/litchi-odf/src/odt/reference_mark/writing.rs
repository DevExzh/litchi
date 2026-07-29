use super::{ReferenceMark, parse_reference_marks, validate_reference_name};
use crate::elements::xml::{TEXT_NAMESPACE, is_bound, namespaced_attribute};
use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;
use std::collections::HashMap;
use std::ops::Range;

const MAX_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 4_096;
const MAX_MARKS: usize = 1_000_000;

/// Canonical XML emitted for a point or range reference target.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ReferenceMarkFragments {
    Point(String),
    Range { start: String, end: String },
}

impl ReferenceMark {
    /// Serialize this target without evaluating any reference field.
    pub fn to_xml_fragments(&self) -> Result<ReferenceMarkFragments> {
        validate_reference_name(&self.name)?;
        let name = escape_attribute(&self.name);
        if self.range {
            Ok(ReferenceMarkFragments::Range {
                start: format!(r#"<text:reference-mark-start text:name="{name}"/>"#),
                end: format!(r#"<text:reference-mark-end text:name="{name}"/>"#),
            })
        } else {
            Ok(ReferenceMarkFragments::Point(format!(
                r#"<text:reference-mark text:name="{name}"/>"#
            )))
        }
    }
}

/// Insert a point at the end of a paragraph, or wrap a paragraph's inline content in a range.
pub fn insert_reference_mark_xml(
    xml: &str,
    paragraph_index: usize,
    mark: &ReferenceMark,
) -> Result<String> {
    let (existing, scan) = validated_scan(xml)?;
    ensure_unique_name(&existing, None, mark.name())?;
    let paragraph = scan.paragraphs.get(paragraph_index).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "reference-mark paragraph index {paragraph_index} is out of bounds"
        ))
    })?;
    let fragments = bind_fragments_if_needed(xml, mark.to_xml_fragments()?);
    match (&paragraph.site, fragments) {
        (
            ParagraphSite::Paired {
                open_end: _,
                close_start,
            },
            ReferenceMarkFragments::Point(fragment),
        ) => insert_at(xml, *close_start, &fragment),
        (
            ParagraphSite::Paired {
                open_end,
                close_start,
            },
            ReferenceMarkFragments::Range { start, end },
        ) => apply_edits(
            xml,
            vec![
                (*close_start..*close_start, end),
                (*open_end..*open_end, start),
            ],
        ),
        (ParagraphSite::Empty { start, end, qname }, ReferenceMarkFragments::Point(fragment)) => {
            expand_empty(xml, *start, *end, qname, &fragment)
        },
        (
            ParagraphSite::Empty { start, end, qname },
            ReferenceMarkFragments::Range {
                start: open,
                end: close,
            },
        ) => expand_empty(xml, *start, *end, qname, &(open + &close)),
    }
}

/// Replace the reference target at a document-order ordinal and return the previous target.
pub fn replace_reference_mark_xml(
    xml: &str,
    ordinal: usize,
    replacement: &ReferenceMark,
) -> Result<String> {
    let (existing, scan) = validated_scan(xml)?;
    let old = existing.get(ordinal).cloned().ok_or_else(|| {
        Error::InvalidFormat(format!("reference-mark ordinal {ordinal} is out of bounds"))
    })?;
    ensure_unique_name(&existing, Some(ordinal), replacement.name())?;
    let location = scan.marks.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat("reference-mark semantic and lexical scans disagree".to_string())
    })?;
    let fragments = bind_fragments_if_needed(xml, replacement.to_xml_fragments()?);
    let output = match (location, fragments) {
        (MarkLocation::Point(span), ReferenceMarkFragments::Point(fragment)) => {
            apply_edits(xml, vec![(span.clone(), fragment)])?
        },
        (MarkLocation::Point(span), ReferenceMarkFragments::Range { start, end }) => {
            apply_edits(xml, vec![(span.clone(), start + &end)])?
        },
        (MarkLocation::Range { start, end }, ReferenceMarkFragments::Point(fragment)) => {
            apply_edits(
                xml,
                vec![(end.clone(), String::new()), (start.clone(), fragment)],
            )?
        },
        (
            MarkLocation::Range { start, end },
            ReferenceMarkFragments::Range {
                start: open,
                end: close,
            },
        ) => apply_edits(xml, vec![(end.clone(), close), (start.clone(), open)])?,
    };
    let _ = old;
    Ok(output)
}

/// Remove marker elements while retaining all text and markup enclosed by a range.
pub fn remove_reference_mark_xml(xml: &str, ordinal: usize) -> Result<String> {
    let (existing, scan) = validated_scan(xml)?;
    let old = existing.get(ordinal).cloned().ok_or_else(|| {
        Error::InvalidFormat(format!("reference-mark ordinal {ordinal} is out of bounds"))
    })?;
    let location = scan.marks.get(ordinal).ok_or_else(|| {
        Error::InvalidFormat("reference-mark semantic and lexical scans disagree".to_string())
    })?;
    let edits = match location {
        MarkLocation::Point(span) => vec![(span.clone(), String::new())],
        MarkLocation::Range { start, end } => {
            vec![(end.clone(), String::new()), (start.clone(), String::new())]
        },
    };
    let _ = old;
    apply_edits(xml, edits)
}

fn ensure_unique_name(marks: &[ReferenceMark], except: Option<usize>, name: &str) -> Result<()> {
    validate_reference_name(name)?;
    if marks
        .iter()
        .enumerate()
        .any(|(index, mark)| Some(index) != except && mark.name() == name)
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate reference-mark identity '{name}'"
        )));
    }
    Ok(())
}

fn escape_attribute(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '"' => escaped.push_str("&quot;"),
            '\r' => escaped.push_str("&#13;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

fn bind_fragments_if_needed(
    xml: &str,
    fragments: ReferenceMarkFragments,
) -> ReferenceMarkFragments {
    let double = format!(r#"xmlns:text="{TEXT_NAMESPACE_STRING}""#);
    let single = format!("xmlns:text='{TEXT_NAMESPACE_STRING}'");
    if xml.contains(&double) || xml.contains(&single) {
        return fragments;
    }
    match fragments {
        ReferenceMarkFragments::Point(fragment) => {
            ReferenceMarkFragments::Point(bind_fragment(fragment))
        },
        ReferenceMarkFragments::Range { start, end } => ReferenceMarkFragments::Range {
            start: bind_fragment(start),
            end: bind_fragment(end),
        },
    }
}

const TEXT_NAMESPACE_STRING: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn bind_fragment(fragment: String) -> String {
    fragment.replacen(
        " text:name=",
        &format!(r#" xmlns:text="{TEXT_NAMESPACE_STRING}" text:name="#),
        1,
    )
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
}

enum MarkLocation {
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
    marks: Vec<MarkLocation>,
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

fn validated_scan(xml: &str) -> Result<(Vec<ReferenceMark>, Scan)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "reference-mark XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }
    let marks = parse_reference_marks(xml)?;
    let scan = scan_locations(xml)?;
    if marks.len() != scan.marks.len() {
        return Err(Error::InvalidFormat(
            "reference-mark semantic and lexical scans disagree".to_string(),
        ));
    }
    Ok((marks, scan))
}

fn scan_locations(xml: &str) -> Result<Scan> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous_end = 0usize;
    let mut depth = 0usize;
    let mut mark_count = 0usize;
    let mut open_elements = Vec::<OpenElement>::new();
    let mut open_ranges = HashMap::<String, usize>::new();
    let mut paragraphs = Vec::new();
    let mut marks = Vec::<Option<MarkLocation>>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid mutable reference-mark XML: {error}"))
            })?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        drop(namespace);
        let event_end = reader.buffer_position() as usize;
        let span = Span {
            start: previous_end,
            end: event_end,
        };
        match event {
            Event::Start(ref element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("reference-mark XML depth overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return invalid("reference-mark XML nesting exceeds 4096");
                }
                let local_name = element.local_name();
                let local = local_name.as_ref();
                let paragraph = text_element && matches!(local, b"p" | b"h");
                let marker = if text_element {
                    marker_start(
                        &reader,
                        element,
                        local,
                        &span,
                        &mut marks,
                        &mut open_ranges,
                        &mut mark_count,
                        false,
                    )?
                } else {
                    None
                };
                open_elements.push(OpenElement {
                    local: local.to_vec(),
                    start: span.start,
                    open_end: span.end,
                    paragraph,
                    marker,
                });
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if text_element && matches!(local, b"p" | b"h") {
                    let qname = std::str::from_utf8(element.name().as_ref())
                        .map_err(|_| {
                            Error::InvalidFormat(
                                "non-UTF-8 reference-mark paragraph name".to_string(),
                            )
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
                    marker_start(
                        &reader,
                        element,
                        local,
                        &span,
                        &mut marks,
                        &mut open_ranges,
                        &mut mark_count,
                        true,
                    )?;
                }
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("reference-mark XML stack underflow".to_string())
                })?;
                let open = open_elements.pop().ok_or_else(|| {
                    Error::InvalidFormat("reference-mark XML stack mismatch".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("reference-mark XML has mismatched elements");
                }
                if open.paragraph {
                    paragraphs.push(ParagraphLocation {
                        site: ParagraphSite::Paired {
                            open_end: open.open_end,
                            close_start: span.start,
                        },
                    });
                }
                finish_marker(open, &span, &mut marks, &mut open_ranges)?;
            },
            Event::DocType(_) => {
                return invalid("DOCTYPE is not allowed in mutable reference-mark XML");
            },
            Event::Eof => break,
            _ => {},
        }
        previous_end = event_end;
        buffer.clear();
    }
    if depth != 0 || !open_elements.is_empty() || !open_ranges.is_empty() {
        return invalid("incomplete mutable reference-mark XML");
    }
    paragraphs.sort_by_key(|paragraph| match paragraph.site {
        ParagraphSite::Paired { open_end, .. } => open_end,
        ParagraphSite::Empty { start, .. } => start,
    });
    let marks = marks
        .into_iter()
        .map(|mark| {
            mark.ok_or_else(|| {
                Error::InvalidFormat("incomplete reference-mark range scan".to_string())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(Scan { paragraphs, marks })
}

#[allow(clippy::too_many_arguments)]
fn marker_start(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    span: &Span,
    marks: &mut Vec<Option<MarkLocation>>,
    open_ranges: &mut HashMap<String, usize>,
    mark_count: &mut usize,
    empty: bool,
) -> Result<Option<MarkerKind>> {
    let kind = match local {
        b"reference-mark" => 0,
        b"reference-mark-start" => 1,
        b"reference-mark-end" => 2,
        _ => return Ok(None),
    };
    let name = namespaced_attribute(reader, element, TEXT_NAMESPACE, b"name", "reference mark")?
        .ok_or_else(|| Error::InvalidFormat("reference mark requires text:name".to_string()))?;
    match kind {
        0 => {
            capacity(*mark_count)?;
            let order = *mark_count;
            *mark_count += 1;
            marks.push(None);
            if empty {
                marks[order] = Some(MarkLocation::Point(span.start..span.end));
                Ok(None)
            } else {
                Ok(Some(MarkerKind::Point(order)))
            }
        },
        1 => {
            capacity(*mark_count)?;
            let order = *mark_count;
            *mark_count += 1;
            marks.push(if empty {
                Some(MarkLocation::Range {
                    start: span.start..span.end,
                    end: 0..0,
                })
            } else {
                None
            });
            if open_ranges.insert(name.clone(), order).is_some() {
                return invalid("duplicate open mutable reference-mark range");
            }
            if empty {
                Ok(None)
            } else {
                Ok(Some(MarkerKind::Start(name, order)))
            }
        },
        _ => {
            if empty {
                close_range(&name, span.start..span.end, marks, open_ranges)?;
                Ok(None)
            } else {
                Ok(Some(MarkerKind::End(name)))
            }
        },
    }
}

fn finish_marker(
    open: OpenElement,
    closing: &Span,
    marks: &mut [Option<MarkLocation>],
    open_ranges: &mut HashMap<String, usize>,
) -> Result<()> {
    let Some(marker) = open.marker else {
        return Ok(());
    };
    if open.open_end != closing.start {
        return invalid("reference-mark elements must be empty");
    }
    let span = open.start..closing.end;
    match marker {
        MarkerKind::Point(order) => marks[order] = Some(MarkLocation::Point(span)),
        MarkerKind::Start(name, order) => {
            if open_ranges.get(&name).copied() != Some(order) {
                return invalid("reference-mark range identity changed during scan");
            }
            marks[order] = Some(MarkLocation::Range {
                start: span,
                end: 0..0,
            });
        },
        MarkerKind::End(name) => close_range(&name, span, marks, open_ranges)?,
    }
    Ok(())
}

fn close_range(
    name: &str,
    end: Range<usize>,
    marks: &mut [Option<MarkLocation>],
    open_ranges: &mut HashMap<String, usize>,
) -> Result<()> {
    let order = open_ranges.remove(name).ok_or_else(|| {
        Error::InvalidFormat(format!("reference-mark-end has no open range '{name}'"))
    })?;
    match marks.get_mut(order).and_then(Option::as_mut) {
        Some(MarkLocation::Range { end: target, .. }) => *target = end,
        _ => return invalid("reference-mark range start is incomplete"),
    }
    Ok(())
}

fn capacity(count: usize) -> Result<()> {
    if count >= MAX_MARKS {
        return invalid("too many mutable reference marks");
    }
    Ok(())
}

fn insert_at(xml: &str, offset: usize, fragment: &str) -> Result<String> {
    apply_edits(xml, vec![(offset..offset, fragment.to_string())])
}

fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid empty paragraph span".to_string()))?;
    let trimmed = source
        .strip_suffix("/>")
        .ok_or_else(|| Error::InvalidFormat("empty paragraph does not end with />".to_string()))?;
    apply_edits(
        xml,
        vec![(start..end, format!("{trimmed}>{content}</{qname}>"))],
    )
}

fn apply_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String> {
    edits.sort_by_key(|edit| std::cmp::Reverse(edit.0.start));
    let mut output = xml.to_string();
    let mut previous_start = xml.len();
    for (span, replacement) in edits {
        if span.start > span.end || span.end > previous_start || span.end > output.len() {
            return invalid("overlapping or invalid reference-mark mutation spans");
        }
        output.replace_range(span.clone(), &replacement);
        previous_start = span.start;
    }
    Ok(output)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn wrapped(body: &str) -> String {
        format!(
            r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:text="{TEXT}">{body}</o:text>"#
        )
    }

    #[test]
    fn canonical_point_and_range_fragments_round_trip() {
        let point = ReferenceMark::point("point<&\"");
        assert_eq!(
            point.to_xml_fragments().unwrap(),
            ReferenceMarkFragments::Point(
                r#"<text:reference-mark text:name="point&lt;&amp;&quot;"/>"#.to_string()
            )
        );
        let range = ReferenceMark::range("range");
        assert_eq!(
            range.to_xml_fragments().unwrap(),
            ReferenceMarkFragments::Range {
                start: r#"<text:reference-mark-start text:name="range"/>"#.to_string(),
                end: r#"<text:reference-mark-end text:name="range"/>"#.to_string(),
            }
        );
        let inserted =
            insert_reference_mark_xml(&wrapped("<t:p>payload</t:p>"), 0, &range).unwrap();
        let parsed = parse_reference_marks(&inserted).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!(parsed[0].name(), "range");
        assert_eq!(parsed[0].text(), "payload");
    }

    #[test]
    fn lossless_insert_replace_remove_preserves_unrelated_and_enclosed_xml() {
        let source = wrapped(
            r#"<t:p t:style-name="P">A&amp;<t:span t:style-name="S">B</t:span><!--keep--></t:p><t:p/>"#,
        );
        let with_range =
            insert_reference_mark_xml(&source, 0, &ReferenceMark::range("r1")).unwrap();
        assert!(with_range.contains(r#"<text:reference-mark-start text:name="r1"/>A&amp;<t:span t:style-name="S">B</t:span><!--keep--><text:reference-mark-end text:name="r1"/>"#));
        let replaced =
            replace_reference_mark_xml(&with_range, 0, &ReferenceMark::point("p2")).unwrap();
        assert!(replaced.contains(r#"<text:reference-mark text:name="p2"/>A&amp;<t:span t:style-name="S">B</t:span><!--keep-->"#));
        let removed = remove_reference_mark_xml(&replaced, 0).unwrap();
        assert_eq!(removed, source);
        let empty = insert_reference_mark_xml(&source, 1, &ReferenceMark::point("empty")).unwrap();
        assert!(empty.contains(r#"<t:p><text:reference-mark text:name="empty"/></t:p>"#));
    }

    #[test]
    fn hostile_namespaces_identity_content_and_resources_are_rejected() {
        for body in [
            r#"<t:p><t:reference-mark t:name="x" u:extra="1"/></t:p>"#,
            r#"<t:p><t:reference-mark u:name="x"/></t:p>"#,
            r#"<t:p><t:reference-mark t:name="x"/><t:reference-mark t:name="x"/></t:p>"#,
            r#"<t:p><t:reference-mark t:name="x"/><t:reference-mark-start t:name="x"/><t:reference-mark-end t:name="x"/></t:p>"#,
            r#"<t:p><t:reference-mark t:name="x">bad</t:reference-mark></t:p>"#,
            r#"<!DOCTYPE x><t:p/>"#,
        ] {
            let xml = format!(
                r#"<o:text xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="{TEXT}" xmlns:u="urn:hostile">{body}</o:text>"#
            );
            assert!(
                insert_reference_mark_xml(&xml, 0, &ReferenceMark::point("new")).is_err(),
                "accepted {body}"
            );
        }
        assert!(
            ReferenceMark::point("x".repeat(65_537))
                .to_xml_fragments()
                .is_err()
        );
        assert!(
            ReferenceMark::point("bad\0name")
                .to_xml_fragments()
                .is_err()
        );
    }

    #[test]
    fn producer_shaped_point_field_and_overlapping_ranges_round_trip() {
        // LibreOffice/Zotero emits long metadata-bearing names and adjacent range markers.
        let name = r#"ZOTERO_ITEM CSL_CITATION {&quot;citationID&quot;:&quot;abc&quot;} RNDxyz"#;
        let xml = wrapped(&format!(
            r#"<t:p><t:reference-mark-start t:name="{name}"/>(<t:span t:style-name="T1">Author</t:span>, 2026)<t:reference-mark-start t:name="second"/> tail<t:reference-mark-end t:name="{name}"/><t:reference-mark-end t:name="second"/></t:p><t:p><t:reference-mark t:name="anchor"/><t:reference-ref t:reference-format="page" t:ref-name="anchor">1</t:reference-ref></t:p>"#
        ));
        let marks = parse_reference_marks(&xml).unwrap();
        assert_eq!(marks.len(), 3);
        assert_eq!(marks[0].text(), "(Author, 2026) tail");
        assert_eq!(marks[1].text(), " tail");
        let replaced =
            replace_reference_mark_xml(&xml, 2, &ReferenceMark::point("odfpy-anchor")).unwrap();
        assert!(replaced.contains(
            r#"<t:reference-ref t:reference-format="page" t:ref-name="anchor">1</t:reference-ref>"#
        ));
        assert_eq!(parse_reference_marks(&replaced).unwrap().len(), 3);
    }
}
