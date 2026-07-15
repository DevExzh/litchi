//! Semantic parsing of OpenDocument ruby annotations.

use crate::elements::xml::{
    TEXT_NAMESPACE, append_checked, append_text_control, decode_reference, is_bound,
    namespaced_attribute,
};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

const MAX_DEPTH: usize = 4_096;
const MAX_RUBIES: usize = 1_000_000;

/// A ruby base/pronunciation pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Ruby {
    style_name: Option<String>,
    base: String,
    text: String,
    text_style_name: Option<String>,
}

impl Ruby {
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    pub fn base(&self) -> &str {
        &self.base
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }
}

struct ActiveRuby {
    ruby: Ruby,
    depth: usize,
    base_depth: Option<usize>,
    base_seen: bool,
    text_depth: Option<usize>,
    text_seen: bool,
    skip_text_depth: Option<usize>,
    order: usize,
}

pub(crate) fn parse_rubies(xml: &str) -> Result<Vec<Ruby>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut document_depth = 0usize;
    let mut active = Vec::<ActiveRuby>::new();
    let mut rubies = Vec::new();
    let mut next_order = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ruby XML: {error}")))?;
        let text_element = is_bound(&namespace, TEXT_NAMESPACE);
        match event {
            Event::Start(ref element) => {
                document_depth = checked_depth(document_depth)?;
                for ruby in &mut active {
                    ruby.depth = checked_depth(ruby.depth)?;
                }
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) {
                    return Err(Error::InvalidFormat(
                        "text:ruby-text may contain only text".to_string(),
                    ));
                }
                if text_element && element.local_name().as_ref() == b"ruby" {
                    if next_order >= MAX_RUBIES {
                        return Err(Error::InvalidFormat(format!(
                            "document exceeds {MAX_RUBIES} ruby annotations"
                        )));
                    }
                    active.push(ActiveRuby {
                        ruby: Ruby {
                            style_name: namespaced_attribute(
                                &reader,
                                element,
                                TEXT_NAMESPACE,
                                b"style-name",
                                "ruby",
                            )?,
                            base: String::new(),
                            text: String::new(),
                            text_style_name: None,
                        },
                        depth: 1,
                        base_depth: None,
                        base_seen: false,
                        text_depth: None,
                        text_seen: false,
                        skip_text_depth: None,
                        order: next_order,
                    });
                    next_order += 1;
                } else if !active.is_empty() {
                    let last = active.len() - 1;
                    if text_element && element.local_name().as_ref() == b"ruby-base" {
                        if active[last].depth != 2
                            || active[last].base_seen
                            || active[last].text_seen
                        {
                            return Err(Error::InvalidFormat(
                                "invalid text:ruby-base placement".to_string(),
                            ));
                        }
                        active[last].base_seen = true;
                        active[last].base_depth = Some(active[last].depth);
                    } else if text_element && element.local_name().as_ref() == b"ruby-text" {
                        if active[last].depth != 2
                            || !active[last].base_seen
                            || active[last].text_seen
                        {
                            return Err(Error::InvalidFormat(
                                "invalid text:ruby-text placement".to_string(),
                            ));
                        }
                        for ruby in &mut active[..last] {
                            if ruby.base_depth.is_some() {
                                ruby.skip_text_depth = Some(ruby.depth);
                            }
                        }
                        active[last].text_seen = true;
                        active[last].text_depth = Some(active[last].depth);
                        active[last].ruby.text_style_name = namespaced_attribute(
                            &reader,
                            element,
                            TEXT_NAMESPACE,
                            b"style-name",
                            "ruby text",
                        )?;
                    } else if text_element {
                        for ruby in &mut active {
                            if ruby.base_depth.is_some() && ruby.skip_text_depth.is_none() {
                                append_text_control(&reader, element, &mut ruby.ruby.base)?;
                            }
                        }
                    }
                }
            },
            Event::Empty(ref element) if !active.is_empty() => {
                let last = active.len() - 1;
                if text_element && element.local_name().as_ref() == b"ruby" {
                    return Err(Error::InvalidFormat(
                        "text:ruby requires base and text".to_string(),
                    ));
                } else if text_element && element.local_name().as_ref() == b"ruby-base" {
                    if active[last].depth != 1 || active[last].base_seen || active[last].text_seen {
                        return Err(Error::InvalidFormat(
                            "invalid text:ruby-base placement".to_string(),
                        ));
                    }
                    active[last].base_seen = true;
                } else if text_element && element.local_name().as_ref() == b"ruby-text" {
                    if active[last].depth != 1 || !active[last].base_seen || active[last].text_seen
                    {
                        return Err(Error::InvalidFormat(
                            "invalid text:ruby-text placement".to_string(),
                        ));
                    }
                    active[last].text_seen = true;
                    active[last].ruby.text_style_name = namespaced_attribute(
                        &reader,
                        element,
                        TEXT_NAMESPACE,
                        b"style-name",
                        "ruby text",
                    )?;
                } else if text_element {
                    for ruby in &mut active {
                        if ruby.base_depth.is_some() && ruby.skip_text_depth.is_none() {
                            append_text_control(&reader, element, &mut ruby.ruby.base)?;
                        }
                    }
                }
            },
            Event::Empty(ref element)
                if text_element && element.local_name().as_ref() == b"ruby" =>
            {
                return Err(Error::InvalidFormat(
                    "text:ruby requires base and text".to_string(),
                ));
            },
            Event::Text(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::InvalidFormat(format!("invalid ruby text: {error}")))?;
                append_ruby_text(&mut active, &value)?;
            },
            Event::CData(ref value) if !active.is_empty() => {
                let value = value
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid ruby CDATA: {error}"))
                    })?;
                append_ruby_text(&mut active, &value)?;
            },
            Event::GeneralRef(ref reference) if !active.is_empty() => {
                append_ruby_text(&mut active, &decode_reference(reference, "ruby")?)?;
            },
            Event::End(_) => {
                document_depth = document_depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ruby XML stack underflow".to_string()))?;
                for ruby in &mut active {
                    if ruby.base_depth == Some(ruby.depth) {
                        ruby.base_depth = None;
                    }
                    if ruby.text_depth == Some(ruby.depth) {
                        ruby.text_depth = None;
                    }
                    if ruby.skip_text_depth == Some(ruby.depth) {
                        ruby.skip_text_depth = None;
                    }
                    ruby.depth = ruby.depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("ruby element stack underflow".to_string())
                    })?;
                }
                if active.last().is_some_and(|ruby| ruby.depth == 0) {
                    let finished = active.pop().expect("checked ruby");
                    if !finished.base_seen || !finished.text_seen {
                        return Err(Error::InvalidFormat(
                            "text:ruby requires base and text".to_string(),
                        ));
                    }
                    rubies.push((finished.order, finished.ruby));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if document_depth != 0 || !active.is_empty() {
        return Err(Error::InvalidFormat(
            "incomplete ruby XML structure".to_string(),
        ));
    }
    rubies.sort_by_key(|(order, _)| *order);
    Ok(rubies.into_iter().map(|(_, ruby)| ruby).collect())
}

fn append_ruby_text(active: &mut [ActiveRuby], value: &str) -> Result<()> {
    for ruby in active {
        if ruby.text_depth.is_some() {
            append_checked(&mut ruby.ruby.text, value)?;
        } else if ruby.base_depth.is_some() && ruby.skip_text_depth.is_none() {
            append_checked(&mut ruby.ruby.base, value)?;
        }
    }
    Ok(())
}

fn checked_depth(depth: usize) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("ruby nesting depth overflow".to_string()))?;
    if depth > MAX_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "ruby nesting exceeds {MAX_DEPTH} levels"
        )));
    }
    Ok(depth)
}

#[cfg(test)]
mod tests {
    use super::*;

    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    #[test]
    fn parses_nested_ruby_pairs_with_styles_and_exact_text() {
        let xml = format!(
            r#"<x:p xmlns:x="{TEXT}"><x:ruby x:style-name="Outer"><x:ruby-base>A&amp;<x:s x:c="2"/><x:ruby><x:ruby-base>B</x:ruby-base><x:ruby-text x:style-name="InnerText">b</x:ruby-text></x:ruby>C</x:ruby-base><x:ruby-text x:style-name="Pronunciation"><![CDATA[abc]]></x:ruby-text></x:ruby></x:p>"#
        );
        let rubies = parse_rubies(&xml).unwrap();
        assert_eq!(rubies.len(), 2);
        assert_eq!(rubies[0].style_name(), Some("Outer"));
        assert_eq!(rubies[0].base(), "A&  BC");
        assert_eq!(rubies[0].text(), "abc");
        assert_eq!(rubies[0].text_style_name(), Some("Pronunciation"));
        assert_eq!(rubies[1].style_name(), None);
        assert_eq!(rubies[1].base(), "B");
        assert_eq!(rubies[1].text(), "b");
        assert_eq!(rubies[1].text_style_name(), Some("InnerText"));
    }

    #[test]
    fn rubies_reject_invalid_structure_and_ambiguous_attributes() {
        let missing = format!(r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base></x:ruby>"#);
        assert!(parse_rubies(&missing).is_err());
        let wrong_order = format!(
            r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-text>a</x:ruby-text><x:ruby-base>A</x:ruby-base></x:ruby>"#
        );
        assert!(parse_rubies(&wrong_order).is_err());
        let duplicate = format!(
            r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base><x:ruby-base>B</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>"#
        );
        assert!(parse_rubies(&duplicate).is_err());
        let text_child = format!(
            r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>A</x:ruby-base><x:ruby-text><x:span>a</x:span></x:ruby-text></x:ruby>"#
        );
        assert!(parse_rubies(&text_child).is_err());
        let aliases = format!(
            r#"<x:ruby xmlns:x="{TEXT}" xmlns:y="{TEXT}" x:style-name="A" y:style-name="B"><x:ruby-base>A</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>"#
        );
        assert!(parse_rubies(&aliases).is_err());
        let empty = format!(r#"<x:ruby xmlns:x="{TEXT}"/>"#);
        assert!(parse_rubies(&empty).is_err());
        assert!(parse_rubies("<x:ruby>").is_err());
    }

    #[test]
    fn rubies_enforce_nesting_bound() {
        let mut xml = format!(r#"<x:ruby xmlns:x="{TEXT}"><x:ruby-base>"#);
        for _ in 0..MAX_DEPTH {
            xml.push_str("<x:span>");
        }
        for _ in 0..MAX_DEPTH {
            xml.push_str("</x:span>");
        }
        xml.push_str("</x:ruby-base><x:ruby-text>a</x:ruby-text></x:ruby>");
        assert!(parse_rubies(&xml).is_err());
    }
}
