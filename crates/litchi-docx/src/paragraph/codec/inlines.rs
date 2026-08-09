//! Lossless direct paragraph-child traversal.

use std::sync::Arc;

use litchi_core::XmlSlice;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};

use super::super::model::{Inline, OpaqueInline, Paragraph, Run};
use super::xml::is_fragment_word_name;

const MAX_INLINE_DEPTH: usize = 128;
const MAX_INLINE_NODES: usize = 1_000_000;

impl Paragraph {
    /// Return every direct paragraph child in source order.
    ///
    /// Paragraph properties are metadata and are not included. Direct
    /// `WordprocessingML` runs become [`Inline::Run`]; every other child remains
    /// available as [`Inline::Unknown`] with exact inert XML bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, unterminated, or resource-exhausting
    /// paragraph XML.
    pub fn inlines(&self) -> Result<Vec<Inline>> {
        let xml = self.xml_bytes();
        let (source, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut paragraph_depth = None;
        let mut capture = None::<(bool, bool, usize, usize)>;
        let mut inlines = Vec::new();
        let mut saw_paragraph = false;
        let mut depth = 0usize;
        let mut nodes = 0usize;

        loop {
            let event_start =
                usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                    Error::InvalidFormat("Word inline offset does not fit usize".into())
                })?;
            let raw_event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(raw_event);
            let event_end =
                usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                    Error::InvalidFormat("Word inline offset does not fit usize".into())
                })?;

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word inline element counter overflow".into())
                })?;
                if nodes > MAX_INLINE_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word paragraph exceeds {MAX_INLINE_NODES} elements"
                    )));
                }
            }

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word inline nesting is too deep".into())
                    })?;
                    if depth > MAX_INLINE_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "Word inline nesting exceeds the {MAX_INLINE_DEPTH} depth limit"
                        )));
                    }
                    if paragraph_depth.is_none() && element.local_name().as_ref() == b"p" {
                        if saw_paragraph {
                            return Err(Error::InvalidFormat(
                                "paragraph fragment contains multiple roots".into(),
                            ));
                        }
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        if is_fragment_word_name(&namespace, element.name(), b"p", &fragment_prefix)
                        {
                            paragraph_depth = Some(depth);
                            saw_paragraph = true;
                        }
                    } else if paragraph_depth.is_some_and(|root| depth == root + 1)
                        && !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"pPr",
                            &fragment_prefix,
                        )
                    {
                        let is_run = is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"r",
                            &fragment_prefix,
                        );
                        let is_hyperlink = is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"hyperlink",
                            &fragment_prefix,
                        );
                        capture = Some((is_run, is_hyperlink, event_start, depth));
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word inline nesting is too deep".into())
                    })?;
                    if child_depth == 1 && element.local_name().as_ref() == b"p" {
                        if saw_paragraph {
                            return Err(Error::InvalidFormat(
                                "paragraph fragment contains multiple roots".into(),
                            ));
                        }
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        if is_fragment_word_name(&namespace, element.name(), b"p", &fragment_prefix)
                        {
                            saw_paragraph = true;
                        }
                    } else if paragraph_depth.is_some_and(|root| child_depth == root + 1)
                        && !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"pPr",
                            &fragment_prefix,
                        )
                    {
                        let is_run = is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"r",
                            &fragment_prefix,
                        );
                        let is_hyperlink = is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"hyperlink",
                            &fragment_prefix,
                        );
                        push_inline(
                            &mut inlines,
                            &source,
                            base_offset,
                            is_run,
                            is_hyperlink,
                            event_start,
                            event_end,
                        )?;
                    }
                },
                Event::End(element) => {
                    if let Some((is_run, is_hyperlink, start, capture_depth)) = capture
                        && depth == capture_depth
                    {
                        push_inline(
                            &mut inlines,
                            &source,
                            base_offset,
                            is_run,
                            is_hyperlink,
                            start,
                            event_end,
                        )?;
                        capture = None;
                    }
                    if paragraph_depth == Some(depth)
                        && is_fragment_word_name(&namespace, element.name(), b"p", &fragment_prefix)
                    {
                        paragraph_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word inline nesting".into())
                    })?;
                },
                Event::Eof if depth != 0 || capture.is_some() => {
                    return Err(Error::InvalidFormat(
                        "unterminated Word paragraph XML".into(),
                    ));
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
        }
        if !saw_paragraph {
            return Err(Error::InvalidFormat(
                "paragraph XML has no WordprocessingML root".into(),
            ));
        }
        Ok(inlines)
    }
}

fn push_inline(
    inlines: &mut Vec<Inline>,
    source: &Arc<Vec<u8>>,
    base_offset: u32,
    is_run: bool,
    is_hyperlink: bool,
    start: usize,
    end: usize,
) -> Result<()> {
    let relative_start = u32::try_from(start).map_err(|_conversion_error| {
        Error::InvalidFormat("Word inline offset exceeds u32".into())
    })?;
    let length = u32::try_from(
        end.checked_sub(start)
            .ok_or_else(|| Error::InvalidFormat("invalid Word inline range".into()))?,
    )
    .map_err(|_conversion_error| Error::InvalidFormat("Word inline length exceeds u32".into()))?;
    let absolute_start = base_offset
        .checked_add(relative_start)
        .ok_or_else(|| Error::InvalidFormat("Word inline absolute offset exceeds u32".into()))?;
    inlines.push(if is_run {
        Inline::Run(Box::new(Run::from_slice(XmlSlice::new(
            Arc::clone(source),
            absolute_start,
            length,
        ))))
    } else {
        Inline::Unknown(Box::new(OpaqueInline::from_arc_range(
            Arc::clone(source),
            absolute_start,
            length,
            is_hyperlink,
        )))
    });
    Ok(())
}
