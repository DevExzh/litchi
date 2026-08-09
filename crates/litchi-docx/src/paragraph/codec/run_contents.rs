//! Ordered, lossless direct run-child traversal.

use std::sync::Arc;

use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::error::{Error, Result};
use crate::image::parse_inline_images;

use super::super::model::{OpaqueRunContent, Run, RunContent};
use super::text::extract_word_text;
use super::xml::{is_fragment_word_name, word_attribute_value};

const MAX_RUN_CONTENT_DEPTH: usize = 128;
const MAX_RUN_CONTENT_NODES: usize = 1_000_000;

#[derive(Clone, Copy)]
enum Kind {
    Text,
    Tab,
    Break,
    CarriageReturn,
    NoBreakHyphen,
    SoftHyphen,
    Drawing,
    FootnoteReference,
    EndnoteReference,
    FootnoteMark,
    EndnoteMark,
    Unknown,
}

impl Run {
    /// Return every direct run child in source order.
    ///
    /// Run properties are metadata and are not included. Common textual
    /// controls, drawings, and note references have typed values; every other
    /// child remains available as exact inert XML through
    /// [`RunContent::Unknown`].
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, unterminated, or resource-exhausting
    /// run XML, or for invalid typed break/note/image content.
    pub fn contents(&self) -> Result<Vec<RunContent>> {
        let xml = self.xml_bytes();
        let (source, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut run_depth = None;
        let mut capture = None::<(Kind, usize, usize)>;
        let mut contents = Vec::new();
        let mut saw_run = false;
        let mut depth = 0usize;
        let mut nodes = 0usize;

        loop {
            let event_start =
                usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                    Error::InvalidFormat("Word run offset does not fit usize".into())
                })?;
            let raw_event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(raw_event);
            let event_end =
                usize::try_from(reader.buffer_position()).map_err(|_conversion_error| {
                    Error::InvalidFormat("Word run offset does not fit usize".into())
                })?;

            if matches!(event, Event::Start(_) | Event::Empty(_)) {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word run element counter overflow".into())
                })?;
                if nodes > MAX_RUN_CONTENT_NODES {
                    return Err(Error::InvalidFormat(format!(
                        "Word run exceeds {MAX_RUN_CONTENT_NODES} elements"
                    )));
                }
            }

            match event {
                Event::Start(element) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word run nesting is too deep".into())
                    })?;
                    if depth > MAX_RUN_CONTENT_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "Word run nesting exceeds the {MAX_RUN_CONTENT_DEPTH} depth limit"
                        )));
                    }
                    if run_depth.is_none() && element.local_name().as_ref() == b"r" {
                        if saw_run {
                            return Err(Error::InvalidFormat(
                                "run fragment contains multiple roots".into(),
                            ));
                        }
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        if is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix)
                        {
                            run_depth = Some(depth);
                            saw_run = true;
                        }
                    } else if run_depth.is_some_and(|root| depth == root + 1)
                        && !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"rPr",
                            &fragment_prefix,
                        )
                    {
                        capture = Some((
                            classify(&namespace, element.name(), &fragment_prefix),
                            event_start,
                            depth,
                        ));
                    }
                },
                Event::Empty(element) => {
                    let child_depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word run nesting is too deep".into())
                    })?;
                    if child_depth == 1 && element.local_name().as_ref() == b"r" {
                        if saw_run {
                            return Err(Error::InvalidFormat(
                                "run fragment contains multiple roots".into(),
                            ));
                        }
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        if is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix)
                        {
                            saw_run = true;
                        }
                    } else if run_depth.is_some_and(|root| child_depth == root + 1)
                        && !is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"rPr",
                            &fragment_prefix,
                        )
                    {
                        push_content(
                            &mut contents,
                            &source,
                            base_offset,
                            classify(&namespace, element.name(), &fragment_prefix),
                            event_start,
                            event_end,
                            &fragment_prefix,
                        )?;
                    }
                },
                Event::End(element) => {
                    if let Some((kind, start, capture_depth)) = capture
                        && depth == capture_depth
                    {
                        push_content(
                            &mut contents,
                            &source,
                            base_offset,
                            kind,
                            start,
                            event_end,
                            &fragment_prefix,
                        )?;
                        capture = None;
                    }
                    if run_depth == Some(depth)
                        && is_fragment_word_name(&namespace, element.name(), b"r", &fragment_prefix)
                    {
                        run_depth = None;
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::InvalidFormat("invalid Word run nesting".into()))?;
                },
                Event::Eof if depth != 0 || capture.is_some() => {
                    return Err(Error::InvalidFormat("unterminated Word run XML".into()));
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
        if !saw_run {
            return Err(Error::InvalidFormat(
                "run XML has no WordprocessingML root".into(),
            ));
        }
        Ok(contents)
    }
}

#[allow(
    clippy::option_option,
    clippy::ref_option,
    reason = "three states distinguish unobserved, unprefixed, and prefixed XML fragments"
)]
fn classify(
    namespace: &quick_xml::name::ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Kind {
    for (local, kind) in [
        (b"t".as_slice(), Kind::Text),
        (b"tab".as_slice(), Kind::Tab),
        (b"br".as_slice(), Kind::Break),
        (b"cr".as_slice(), Kind::CarriageReturn),
        (b"noBreakHyphen".as_slice(), Kind::NoBreakHyphen),
        (b"softHyphen".as_slice(), Kind::SoftHyphen),
        (b"drawing".as_slice(), Kind::Drawing),
        (b"footnoteReference".as_slice(), Kind::FootnoteReference),
        (b"endnoteReference".as_slice(), Kind::EndnoteReference),
        (b"footnoteRef".as_slice(), Kind::FootnoteMark),
        (b"endnoteRef".as_slice(), Kind::EndnoteMark),
    ] {
        if is_fragment_word_name(namespace, name, local, fragment_prefix) {
            return kind;
        }
    }
    Kind::Unknown
}

#[allow(
    clippy::option_option,
    clippy::ref_option,
    reason = "three states distinguish unobserved, unprefixed, and prefixed XML fragments"
)]
fn push_content(
    contents: &mut Vec<RunContent>,
    source: &Arc<Vec<u8>>,
    base_offset: u32,
    kind: Kind,
    start: usize,
    end: usize,
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> Result<()> {
    let relative_start = u32::try_from(start).map_err(|_conversion_error| {
        Error::InvalidFormat("Word run-child offset exceeds u32".into())
    })?;
    let length = u32::try_from(
        end.checked_sub(start)
            .ok_or_else(|| Error::InvalidFormat("invalid Word run-child range".into()))?,
    )
    .map_err(|_conversion_error| {
        Error::InvalidFormat("Word run-child length exceeds u32".into())
    })?;
    let absolute_start = base_offset
        .checked_add(relative_start)
        .ok_or_else(|| Error::InvalidFormat("Word run-child absolute offset exceeds u32".into()))?;
    let absolute_end = absolute_start
        .checked_add(length)
        .ok_or_else(|| Error::InvalidFormat("Word run-child range exceeds u32".into()))?;
    let raw = source
        .get(absolute_start as usize..absolute_end as usize)
        .ok_or_else(|| Error::InvalidFormat("Word run-child range is outside source".into()))?;

    match kind {
        Kind::Text => contents.push(RunContent::Text(extract_word_text(raw)?)),
        Kind::Tab => contents.push(RunContent::Tab),
        Kind::Break => {
            let breaks = Run::new(raw.to_vec()).breaks()?;
            let run_break = breaks
                .first()
                .copied()
                .ok_or_else(|| Error::InvalidFormat("Word break fragment has no break".into()))?;
            contents.push(RunContent::Break(run_break));
        },
        Kind::CarriageReturn => contents.push(RunContent::CarriageReturn),
        Kind::NoBreakHyphen => contents.push(RunContent::NoBreakHyphen),
        Kind::SoftHyphen => contents.push(RunContent::SoftHyphen),
        Kind::Drawing => {
            let images = parse_inline_images(raw)?;
            if images.is_empty() {
                push_unknown(contents, source, absolute_start, length);
            } else {
                contents.extend(
                    images
                        .into_iter()
                        .map(|image| RunContent::Image(Box::new(image))),
                );
            }
        },
        Kind::FootnoteReference => contents.push(RunContent::FootnoteReference(note_id(
            raw,
            fragment_prefix,
        )?)),
        Kind::EndnoteReference => {
            contents.push(RunContent::EndnoteReference(note_id(raw, fragment_prefix)?));
        },
        Kind::FootnoteMark => contents.push(RunContent::FootnoteMark),
        Kind::EndnoteMark => contents.push(RunContent::EndnoteMark),
        Kind::Unknown => push_unknown(contents, source, absolute_start, length),
    }
    Ok(())
}

#[allow(
    clippy::option_option,
    clippy::ref_option,
    reason = "three states distinguish unobserved, unprefixed, and prefixed XML fragments"
)]
fn note_id(raw: &[u8], fragment_prefix: &Option<Option<Vec<u8>>>) -> Result<u32> {
    let mut reader = NsReader::from_reader(raw);
    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let value = word_attribute_value(
                    &element,
                    b"id",
                    reader.decoder(),
                    &resolver,
                    fragment_prefix,
                )?
                .ok_or_else(|| {
                    Error::InvalidFormat("Word note reference is missing w:id".into())
                })?;
                return value.parse::<u32>().map_err(|_parse_error| {
                    Error::InvalidFormat(format!(
                        "Word note reference id '{value}' is not a non-negative u32"
                    ))
                });
            },
            Event::Eof => {
                return Err(Error::InvalidFormat(
                    "Word note reference has no element".into(),
                ));
            },
            Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
}

fn push_unknown(contents: &mut Vec<RunContent>, source: &Arc<Vec<u8>>, start: u32, length: u32) {
    contents.push(RunContent::Unknown(Box::new(
        OpaqueRunContent::from_arc_range(Arc::clone(source), start, length),
    )));
}
