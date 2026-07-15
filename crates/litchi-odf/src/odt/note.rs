//! Semantic footnote and endnote parsing.

use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, decode_reference, is_bound,
    namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const MAX_DEPTH: usize = 4_096;
const MAX_NOTES: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NoteClass {
    Footnote,
    Endnote,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Note {
    class: NoteClass,
    id: Option<String>,
    citation: String,
    label: Option<String>,
    body: String,
}

impl Note {
    pub fn class(&self) -> NoteClass {
        self.class
    }
    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }
    pub fn citation(&self) -> &str {
        &self.citation
    }
    pub fn label(&self) -> Option<&str> {
        self.label.as_deref()
    }
    pub fn body(&self) -> &str {
        &self.body
    }
}

struct ActiveNote {
    note: Note,
    depth: usize,
    citation_depth: Option<usize>,
    citation_seen: bool,
    body_depth: Option<usize>,
    body_seen: bool,
    paragraph_depth: Option<usize>,
    seen_paragraph: bool,
    skip_body_depth: Option<usize>,
    order: usize,
}

pub(crate) fn parse_notes(xml: &str) -> Result<Vec<Note>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active = Vec::<ActiveNote>::new();
    let mut notes = Vec::new();
    let mut next_order = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| Error::InvalidFormat(format!("invalid note XML: {e}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_depth(document_depth)?;
                for note in &mut active {
                    note.depth = checked_depth(note.depth)?;
                }
                if active
                    .last()
                    .is_some_and(|note| note.citation_depth.is_some())
                {
                    return Err(Error::InvalidFormat(
                        "text:note-citation may contain only text".to_string(),
                    ));
                }
                if text_element && element.local_name().as_ref() == b"note" {
                    if next_order >= MAX_NOTES {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_NOTES} notes"
                        )));
                    }
                    let class = match namespaced_attribute(
                        &reader,
                        element,
                        TEXT_NAMESPACE,
                        b"note-class",
                        "note",
                    )?
                    .as_deref()
                    {
                        Some("footnote") => NoteClass::Footnote,
                        Some("endnote") => NoteClass::Endnote,
                        Some(value) => {
                            return Err(Error::InvalidFormat(format!(
                                "unsupported text:note-class '{value}'"
                            )));
                        },
                        None => {
                            return Err(Error::InvalidFormat(
                                "text:note requires text:note-class".to_string(),
                            ));
                        },
                    };
                    let id = namespaced_attribute(&reader, element, TEXT_NAMESPACE, b"id", "note")?;
                    active.push(ActiveNote {
                        note: Note {
                            class,
                            id,
                            citation: String::new(),
                            label: None,
                            body: String::new(),
                        },
                        depth: 1,
                        citation_depth: None,
                        citation_seen: false,
                        body_depth: None,
                        body_seen: false,
                        paragraph_depth: None,
                        seen_paragraph: false,
                        skip_body_depth: None,
                        order: next_order,
                    });
                    next_order += 1;
                } else if !active.is_empty() {
                    let last = active.len() - 1;
                    if text_element && element.local_name().as_ref() == b"note-citation" {
                        if active[last].depth != 2
                            || active[last].citation_seen
                            || active[last].body_seen
                        {
                            return Err(Error::InvalidFormat(
                                "invalid text:note-citation placement".to_string(),
                            ));
                        }
                        active[last].citation_seen = true;
                        active[last].citation_depth = Some(active[last].depth);
                        active[last].note.label = namespaced_attribute(
                            &reader,
                            element,
                            TEXT_NAMESPACE,
                            b"label",
                            "note citation",
                        )?;
                    } else if text_element && element.local_name().as_ref() == b"note-body" {
                        if active[last].depth != 2
                            || !active[last].citation_seen
                            || active[last].body_seen
                        {
                            return Err(Error::InvalidFormat(
                                "invalid text:note-body placement".to_string(),
                            ));
                        }
                        for note in &mut active[..last] {
                            if note.body_depth.is_some() {
                                note.skip_body_depth = Some(note.depth);
                            }
                        }
                        active[last].body_seen = true;
                        active[last].body_depth = Some(active[last].depth);
                    } else {
                        for note in &mut active {
                            if note.body_depth.is_some()
                                && note.skip_body_depth.is_none()
                                && note.paragraph_depth.is_none()
                                && text_element
                                && matches!(element.local_name().as_ref(), b"p" | b"h")
                            {
                                if note.seen_paragraph {
                                    append_checked(&mut note.note.body, "\n")?;
                                }
                                note.seen_paragraph = true;
                                note.paragraph_depth = Some(note.depth);
                            }
                            if text_element
                                && note.paragraph_depth.is_some()
                                && note.skip_body_depth.is_none()
                            {
                                append_text_control(&reader, element, &mut note.note.body)?;
                            }
                        }
                    }
                }
            },
            Event::Empty(ref element) if !active.is_empty() => {
                let last = active.len() - 1;
                if text_element && element.local_name().as_ref() == b"note" {
                    return Err(Error::InvalidFormat(
                        "text:note requires citation and body".to_string(),
                    ));
                } else if text_element && element.local_name().as_ref() == b"note-citation" {
                    if active[last].depth != 1
                        || active[last].citation_seen
                        || active[last].body_seen
                    {
                        return Err(Error::InvalidFormat(
                            "invalid text:note-citation placement".to_string(),
                        ));
                    }
                    active[last].citation_seen = true;
                    active[last].note.label = namespaced_attribute(
                        &reader,
                        element,
                        TEXT_NAMESPACE,
                        b"label",
                        "note citation",
                    )?;
                } else if text_element && element.local_name().as_ref() == b"note-body" {
                    if active[last].depth != 1
                        || !active[last].citation_seen
                        || active[last].body_seen
                    {
                        return Err(Error::InvalidFormat(
                            "invalid text:note-body placement".to_string(),
                        ));
                    }
                    active[last].body_seen = true;
                } else {
                    for note in &mut active {
                        if note.body_depth.is_some()
                            && note.skip_body_depth.is_none()
                            && text_element
                            && matches!(element.local_name().as_ref(), b"p" | b"h")
                        {
                            if note.seen_paragraph {
                                append_checked(&mut note.note.body, "\n")?;
                            }
                            note.seen_paragraph = true;
                        } else if text_element
                            && note.paragraph_depth.is_some()
                            && note.skip_body_depth.is_none()
                        {
                            append_text_control(&reader, element, &mut note.note.body)?;
                        }
                    }
                }
            },
            Event::Empty(ref element)
                if text_element && element.local_name().as_ref() == b"note" =>
            {
                return Err(Error::InvalidFormat(
                    "text:note requires citation and body".to_string(),
                ));
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| Error::InvalidFormat(format!("invalid note text: {e}")))?;
                append_note_text(&mut active, &value)?;
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|e| Error::InvalidFormat(format!("invalid note CDATA: {e}")))?;
                append_note_text(&mut active, &value)?;
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                append_note_text(&mut active, &decode_reference(reference, "note")?)?
            },
            Event::End(_) => {
                document_depth = document_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("note XML stack underflow".to_string()))?;
                for note in &mut active {
                    if note.citation_depth == Some(note.depth) {
                        note.citation_depth = None;
                    }
                    if note.paragraph_depth == Some(note.depth) {
                        note.paragraph_depth = None;
                    }
                    if note.body_depth == Some(note.depth) {
                        note.body_depth = None;
                    }
                    if note.skip_body_depth == Some(note.depth) {
                        note.skip_body_depth = None;
                    }
                    note.depth = note
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::InvalidFormat("note stack underflow".to_string()))?;
                }
                if active.last().is_some_and(|note| note.depth == 0) {
                    let finished = active.pop().expect("checked note");
                    if !finished.citation_seen || !finished.body_seen {
                        return Err(Error::InvalidFormat(
                            "text:note requires citation and body".to_string(),
                        ));
                    }
                    notes.push((finished.order, finished.note));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete note XML structure".to_string(),
        ));
    }
    notes.sort_by_key(|(order, _)| *order);
    Ok(notes.into_iter().map(|(_, note)| note).collect())
}

fn append_note_text(active: &mut [ActiveNote], value: &str) -> Result<()> {
    for note in active {
        if note.citation_depth.is_some() {
            append_checked(&mut note.note.citation, value)?;
        } else if note.paragraph_depth.is_some() && note.skip_body_depth.is_none() {
            append_checked(&mut note.note.body, value)?;
        }
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("note depth overflow".to_string()))?;
    if depth > MAX_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "note nesting exceeds {MAX_DEPTH} levels"
        )));
    }
    Ok(depth)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_nested_footnotes_and_endnotes_with_exact_content() {
        let xml = format!(
            r#"<x:text xmlns:x="{TEXT}"><x:p>Visible<x:note x:note-class="footnote" x:id="n&amp;1"><x:note-citation x:label="*">1&amp;</x:note-citation><x:note-body><x:p>First<x:s x:c="2"/>X</x:p><x:list><x:list-item><x:p>Outer <x:note x:note-class="endnote" x:id="inner"><x:note-citation>i</x:note-citation><x:note-body><x:p>Inner &amp; <![CDATA[body]]></x:p></x:note-body></x:note> after</x:p></x:list-item></x:list><x:p>Last</x:p></x:note-body></x:note></x:p></x:text>"#
        );
        let notes = parse_notes(&xml).unwrap();
        assert_eq!(notes.len(), 2);
        assert_eq!(notes[0].class(), NoteClass::Footnote);
        assert_eq!(notes[0].id(), Some("n&1"));
        assert_eq!(notes[0].citation(), "1&");
        assert_eq!(notes[0].label(), Some("*"));
        assert_eq!(notes[0].body(), "First  X\nOuter i after\nLast");
        assert_eq!(notes[1].class(), NoteClass::Endnote);
        assert_eq!(notes[1].id(), Some("inner"));
        assert_eq!(notes[1].citation(), "i");
        assert_eq!(notes[1].label(), None);
        assert_eq!(notes[1].body(), "Inner & body");
    }

    #[test]
    fn notes_reject_invalid_structure_and_ambiguous_metadata() {
        let missing_class = format!(
            r#"<x:note xmlns:x="{TEXT}"><x:note-citation>1</x:note-citation><x:note-body/></x:note>"#
        );
        assert!(parse_notes(&missing_class).is_err());
        let invalid_class = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="side"><x:note-citation>1</x:note-citation><x:note-body/></x:note>"#
        );
        assert!(parse_notes(&invalid_class).is_err());
        let missing_body = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"><x:note-citation>1</x:note-citation></x:note>"#
        );
        assert!(parse_notes(&missing_body).is_err());
        let wrong_order = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"><x:note-body/><x:note-citation>1</x:note-citation></x:note>"#
        );
        assert!(parse_notes(&wrong_order).is_err());
        let duplicate = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"><x:note-citation>1</x:note-citation><x:note-citation>2</x:note-citation><x:note-body/></x:note>"#
        );
        assert!(parse_notes(&duplicate).is_err());
        let citation_child = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"><x:note-citation><x:span>1</x:span></x:note-citation><x:note-body/></x:note>"#
        );
        assert!(parse_notes(&citation_child).is_err());
        let aliases = format!(
            r#"<x:note xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:note-class="footnote" y:note-class="endnote"><x:note-citation>1</x:note-citation><x:note-body/></x:note>"#
        );
        assert!(parse_notes(&aliases).is_err());
        let empty = format!(r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"/>"#);
        assert!(parse_notes(&empty).is_err());
        assert!(parse_notes("<x:note>").is_err());
    }

    #[test]
    fn notes_enforce_nesting_bound() {
        let mut xml = format!(
            r#"<x:note xmlns:x="{TEXT}" x:note-class="footnote"><x:note-citation>1</x:note-citation><x:note-body><x:p>"#
        );
        for _ in 0..MAX_DEPTH {
            xml.push_str("<x:span>");
        }
        for _ in 0..MAX_DEPTH {
            xml.push_str("</x:span>");
        }
        xml.push_str("</x:p></x:note-body></x:note>");
        assert!(parse_notes(&xml).is_err());
    }
}
