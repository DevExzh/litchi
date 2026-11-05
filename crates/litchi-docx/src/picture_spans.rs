//! The byte spans of a paragraph's inline pictures that a picture edit
//! rewrites (record 0335).
//!
//! [`picture_spans`] finds the same pictures, in the same order, as
//! [`crate::image::parse_inline_images`] (the order of
//! [`crate::Paragraph::images`]): a `wp:inline` or `wp:anchor` of a
//! `w:drawing` whose `pic:blipFill` names an image by `r:embed`. Each span
//! records the drawing, the run it is the only content of, and the start
//! tags an edit of its extent, name, alternative text or image rewrites.

use std::ops::Range;

use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::Reader;
use quick_xml::events::Event;
use smallvec::SmallVec;

use crate::error::{Error, Result};

/// One picture of a paragraph, as byte ranges of the paragraph's XML.
#[derive(Debug, Clone, Default)]
pub(crate) struct PictureSpan {
    /// The `w:drawing` element.
    pub(crate) drawing: Range<usize>,
    /// The enclosing `w:r`, when the drawing is its only content besides
    /// its properties.
    pub(crate) sole_run: Option<Range<usize>>,
    /// The last `wp:extent` start tag, the one the reader takes.
    pub(crate) extent: Option<Range<usize>>,
    /// The first `a:ext` of a `pic:spPr/a:xfrm`: the shape's own extent.
    pub(crate) shape_extent: Option<Range<usize>>,
    /// The last `wp:docPr` start tag.
    pub(crate) properties: Option<Range<usize>>,
    /// The last `a:blip` start tag of the picture's `pic:blipFill`.
    pub(crate) blip: Option<Range<usize>>,
}

/// Every inline or floating picture of a paragraph's XML, in the order the
/// reader reports them.
pub(crate) fn picture_spans(xml: &[u8]) -> Result<SmallVec<[PictureSpan; 4]>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut found: SmallVec<[PictureSpan; 4]> = SmallVec::new();
    // Open elements: (local name, start, non-property child count).
    let mut stack: Vec<(Vec<u8>, usize, usize)> = Vec::new();
    // The reader's state machine (`parse_inline_images`), kept verbatim so
    // the ordinals agree.
    let mut in_drawing = false;
    let mut in_image = false;
    let mut in_blip_fill = false;
    let mut current = PictureSpan::default();
    let mut embedded = false;
    // Pictures whose drawing has not closed yet, and pictures waiting for
    // their run to close: indexes into `found`, with the run start.
    let mut awaiting_drawing: Vec<usize> = Vec::new();
    let mut awaiting_run: Vec<(usize, usize)> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph position overflows".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("paragraph XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph position overflows".into()))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element.clone(), false),
            Event::Empty(element) => (element.clone(), true),
            Event::End(element) => {
                let local = element.local_name().as_ref().to_vec();
                match local.as_slice() {
                    b"drawing" => in_drawing = false,
                    b"inline" | b"anchor" if in_image => {
                        in_image = false;
                        if embedded {
                            awaiting_drawing.push(found.len());
                            found.push(std::mem::take(&mut current));
                        }
                    },
                    b"blipFill" => in_blip_fill = false,
                    _ => {},
                }
                if let Some((open, opened, children)) = stack.pop() {
                    if open == b"drawing" {
                        let run = stack
                            .iter()
                            .rev()
                            .find(|(name, _, _)| name.as_slice() == b"r")
                            .map(|(_, run, _)| *run);
                        for index in awaiting_drawing.drain(..) {
                            found[index].drawing = opened..end;
                            if let Some(run) = run {
                                awaiting_run.push((index, run));
                            }
                        }
                    } else if open == b"r" {
                        awaiting_run.retain(|(index, run)| {
                            if *run != opened {
                                return true;
                            }
                            if children == 1 {
                                found[*index].sole_run = Some(opened..end);
                            }
                            false
                        });
                    }
                }
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        let local = element.local_name().as_ref().to_vec();
        if let Some((_, _, children)) = stack.last_mut()
            && local != b"rPr"
        {
            *children += 1;
        }
        match local.as_slice() {
            b"drawing" => in_drawing = true,
            b"inline" | b"anchor" if in_drawing => {
                in_image = true;
                current = PictureSpan::default();
                embedded = false;
            },
            b"extent" if in_image => current.extent = Some(start..end),
            b"docPr" if in_image => current.properties = Some(start..end),
            b"blipFill" if in_image => in_blip_fill = true,
            b"blip" if in_blip_fill => {
                for attribute in first_wins(&element).flatten() {
                    let key = attribute.key.as_ref();
                    if (key == b"r:embed" || key.ends_with(b":embed"))
                        && !attribute.value.is_empty()
                    {
                        embedded = true;
                    }
                }
                current.blip = Some(start..end);
            },
            b"ext" if in_image && current.shape_extent.is_none() => {
                let parents: SmallVec<[&[u8]; 2]> = stack
                    .iter()
                    .rev()
                    .take(2)
                    .map(|(name, _, _)| name.as_slice())
                    .collect();
                if parents.as_slice() == [b"xfrm".as_slice(), b"spPr".as_slice()] {
                    current.shape_extent = Some(start..end);
                }
            },
            _ => {},
        }
        if !empty {
            stack.push((local, start, 0));
        }
    }
    Ok(found)
}

/// The value of the start tag's attribute named exactly `key`.
fn find_attribute(tag: &str, key: &str) -> Option<(Range<usize>, char)> {
    let bytes = tag.as_bytes();
    let mut at = tag.find(|c: char| c.is_whitespace() || c == '/' || c == '>')?;
    loop {
        while at < bytes.len() && bytes[at].is_ascii_whitespace() {
            at += 1;
        }
        if at >= bytes.len() || matches!(bytes[at], b'/' | b'>') {
            return None;
        }
        let name_start = at;
        while at < bytes.len() && !matches!(bytes[at], b'=' | b' ' | b'\t' | b'\r' | b'\n') {
            at += 1;
        }
        let name = &tag[name_start..at];
        while at < bytes.len() && bytes[at] != b'=' {
            at += 1;
        }
        at += 1;
        while at < bytes.len() && bytes[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = char::from(*bytes.get(at)?);
        if quote != '"' && quote != '\'' {
            return None;
        }
        let value_start = at + 1;
        let value_end = value_start + tag[value_start..].find(quote)?;
        if name == key {
            return Some((value_start..value_end, quote));
        }
        at = value_end + 1;
    }
}

/// One SmartArt drawing of an element's XML (record 0335): the data-model
/// relationship its `dgm:relIds` names (`r:dm`), the `w:drawing` element,
/// and the run it is the only content of.
#[derive(Debug, Clone)]
pub(crate) struct DiagramSpan {
    pub(crate) data_relationship_id: String,
    pub(crate) drawing: Range<usize>,
    pub(crate) sole_run: Option<Range<usize>>,
}

/// Every SmartArt drawing of an element's XML, in document order.
pub(crate) fn diagram_spans(xml: &[u8]) -> Result<SmallVec<[DiagramSpan; 1]>> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut found: SmallVec<[DiagramSpan; 1]> = SmallVec::new();
    let mut stack: Vec<(Vec<u8>, usize, usize)> = Vec::new();
    // The data relationship of the innermost open drawing.
    let mut drawings: Vec<Option<String>> = Vec::new();
    let mut awaiting_run: Vec<(usize, usize)> = Vec::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("element position overflows".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::InvalidFormat(format!("document XML: {error}")))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("element position overflows".into()))?;
        let (element, empty) = match &event {
            Event::Start(element) => (element.clone(), false),
            Event::Empty(element) => (element.clone(), true),
            Event::End(_) => {
                if let Some((open, opened, children)) = stack.pop() {
                    if open == b"drawing" {
                        if let Some(Some(id)) = drawings.pop() {
                            let run = stack
                                .iter()
                                .rev()
                                .find(|(name, _, _)| name.as_slice() == b"r")
                                .map(|(_, run, _)| *run);
                            if let Some(run) = run {
                                awaiting_run.push((found.len(), run));
                            }
                            found.push(DiagramSpan {
                                data_relationship_id: id,
                                drawing: opened..end,
                                sole_run: None,
                            });
                        }
                    } else if open == b"r" {
                        awaiting_run.retain(|(index, run)| {
                            if *run != opened {
                                return true;
                            }
                            if children == 1 {
                                found[*index].sole_run = Some(opened..end);
                            }
                            false
                        });
                    }
                }
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        let local = element.local_name().as_ref().to_vec();
        if let Some((_, _, children)) = stack.last_mut()
            && local != b"rPr"
        {
            *children += 1;
        }
        if local == b"relIds"
            && let Some(slot) = drawings.last_mut()
        {
            *slot = first_wins(&element).flatten().find_map(|attribute| {
                (attribute.key.local_name().as_ref() == b"dm")
                    .then(|| String::from_utf8(attribute.value.into_owned()).ok())
                    .flatten()
            });
        }
        if !empty {
            if local == b"drawing" {
                drawings.push(None);
            }
            stack.push((local, start, 0));
        }
    }
    Ok(found)
}

/// The qualified name of the start tag's first attribute whose name is
/// `local` or ends with `:local`.
pub(crate) fn attribute_named(tag: &[u8], local: &str) -> Option<String> {
    let mut reader = Reader::from_reader(tag);
    let (Ok(Event::Start(element)) | Ok(Event::Empty(element))) = reader.read_event() else {
        return None;
    };
    let suffix = format!(":{local}");
    first_wins(&element).flatten().find_map(|attribute| {
        let key = std::str::from_utf8(attribute.key.as_ref()).ok()?;
        (key == local || key.ends_with(&suffix)).then(|| key.to_owned())
    })
}

/// The start tag (or empty element) `tag` with attribute `key` (its exact
/// qualified name) set to `value`, appended when absent; `None` for a tag
/// this cannot rewrite.
pub(crate) fn with_attribute(tag: &[u8], key: &str, value: &str) -> Option<Vec<u8>> {
    let text = std::str::from_utf8(tag).ok()?;
    if !text.starts_with('<') {
        return None;
    }
    let escaped = escape(value);
    if let Some((range, _)) = find_attribute(text, key) {
        let mut out = String::with_capacity(text.len() + escaped.len());
        out.push_str(&text[..range.start - 1]);
        out.push('"');
        out.push_str(&escaped);
        out.push('"');
        out.push_str(&text[range.end + 1..]);
        return Some(out.into_bytes());
    }
    let close = if text.ends_with("/>") {
        text.len() - 2
    } else {
        text.len().checked_sub(1)?
    };
    let head = text[..close].trim_end();
    Some(format!("{head} {key}=\"{escaped}\"{}", &text[close..]).into_bytes())
}

/// The start tag (or empty element) `tag` without attribute `key` (its
/// exact qualified name), unchanged when absent; `None` for a tag this
/// cannot rewrite (record 0345).
pub(crate) fn without_attribute(tag: &[u8], key: &str) -> Option<Vec<u8>> {
    let text = std::str::from_utf8(tag).ok()?;
    if !text.starts_with('<') {
        return None;
    }
    let Some((range, _)) = find_attribute(text, key) else {
        return Some(tag.to_vec());
    };
    let bytes = text.as_bytes();
    // Back from the opening quote over `=` and whitespace to the name.
    let mut name_end = range.start.checked_sub(1)?;
    while name_end > 0 && matches!(bytes[name_end - 1], b'=' | b' ' | b'\t' | b'\r' | b'\n') {
        name_end -= 1;
    }
    let mut start = name_end.checked_sub(key.len())?;
    if &text[start..name_end] != key {
        return None;
    }
    while start > 0 && bytes[start - 1].is_ascii_whitespace() {
        start -= 1;
    }
    let mut out = String::with_capacity(text.len());
    out.push_str(&text[..start]);
    out.push_str(&text[range.end + 1..]);
    Some(out.into_bytes())
}

fn escape(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\t' => escaped.push_str("&#9;"),
            '\n' => escaped.push_str("&#10;"),
            '\r' => escaped.push_str("&#13;"),
            other => escaped.push(other),
        }
    }
    escaped
}

#[cfg(test)]
mod tests {
    use super::*;

    const PARAGRAPH: &str = concat!(
        r#"<w:p xmlns:w="w" xmlns:wp="wp" xmlns:a="a" xmlns:pic="pic" xmlns:r="r">"#,
        r#"<w:r><w:t>x</w:t></w:r>"#,
        r#"<w:r><w:rPr/><w:drawing><wp:inline><wp:extent cx="100" cy="200"/><wp:docPr id="1" name="One" descr="d"/>"#,
        r#"<a:graphic><a:graphicData uri="pic"><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill>"#,
        r#"<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
        r#"<w:r><w:t>y</w:t><w:drawing><wp:anchor><wp:extent cx="5" cy="6"/><wp:docPr id="2" name="Two"/>"#,
        r#"<a:graphic><a:graphicData uri="pic"><pic:pic><pic:blipFill><a:blip r:embed="rId8"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><wp:docPr id="3" name="Chart"/><a:graphic><a:graphicData uri="chart"/></a:graphic></wp:inline></w:drawing></w:r>"#,
        "</w:p>"
    );

    #[test]
    fn spans_follow_the_reader() {
        let xml = PARAGRAPH.as_bytes();
        let spans = picture_spans(xml).unwrap();
        let images = crate::image::parse_inline_images(xml).unwrap();
        assert_eq!(spans.len(), images.len());
        assert_eq!(spans.len(), 2);
        let tag = |range: &Option<Range<usize>>| {
            std::str::from_utf8(&xml[range.clone().unwrap()]).unwrap()
        };
        assert_eq!(tag(&spans[0].extent), r#"<wp:extent cx="100" cy="200"/>"#);
        assert_eq!(tag(&spans[0].shape_extent), r#"<a:ext cx="100" cy="200"/>"#);
        assert_eq!(tag(&spans[0].blip), r#"<a:blip r:embed="rId7"/>"#);
        assert!(tag(&spans[0].properties).contains(r#"name="One""#));
        assert!(spans[0].sole_run.is_some());
        assert!(
            std::str::from_utf8(&xml[spans[0].drawing.clone()])
                .unwrap()
                .starts_with("<w:drawing>")
        );
        // The second shares its run with text, and has no shape extent.
        assert!(spans[1].sole_run.is_none());
        assert!(spans[1].shape_extent.is_none());
    }

    #[test]
    fn diagram_drawings_are_found_by_their_data_relationship() {
        let xml = concat!(
            r#"<w:body xmlns:w="w" xmlns:wp="wp" xmlns:a="a" xmlns:dgm="dgm" xmlns:r="r"><w:p>"#,
            r#"<w:r><w:drawing><wp:inline><a:graphic><a:graphicData><dgm:relIds r:dm="rId4" r:lo="rId5"/>"#,
            r#"</a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>x</w:t><w:drawing><wp:inline>"#,
            r#"<a:graphic><a:graphicData><dgm:relIds r:dm="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
            "</w:p></w:body>"
        )
        .as_bytes();
        let spans = diagram_spans(xml).unwrap();
        assert_eq!(spans.len(), 2);
        assert_eq!(spans[0].data_relationship_id, "rId4");
        assert!(spans[0].sole_run.is_some());
        assert_eq!(spans[1].data_relationship_id, "rId9");
        assert!(spans[1].sole_run.is_none());
    }

    #[test]
    fn attributes_are_set_and_found() {
        let tag = br#"<wp:docPr id="1" name='Old'/>"#;
        assert_eq!(
            with_attribute(tag, "name", "A & B").unwrap(),
            br#"<wp:docPr id="1" name="A &amp; B"/>"#
        );
        assert_eq!(
            with_attribute(tag, "descr", "alt").unwrap(),
            br#"<wp:docPr id="1" name='Old' descr="alt"/>"#
        );
        assert_eq!(
            attribute_named(br#"<a:blip xmlns:x="r" x:embed="rId2">"#, "embed").as_deref(),
            Some("x:embed")
        );
        assert_eq!(attribute_named(b"<a:blip/>", "embed"), None);
    }
}
