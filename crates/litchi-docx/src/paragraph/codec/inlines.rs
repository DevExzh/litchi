//! Lossless direct paragraph-child traversal.

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::first_wins;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::bookmark::Bookmark;
use crate::error::{Error, Result};

use super::super::InlineKind;
use super::super::model::{Inline, OpaqueInline, Paragraph, Run, XmlRef};
use super::xml::is_fragment_word_name;

const MAX_INLINE_DEPTH: usize = 128;
const MAX_INLINE_NODES: usize = 1_000_000;

/// How a direct paragraph child is projected.
#[derive(Clone, Copy)]
enum Child {
    Run,
    BookmarkStart,
    BookmarkEnd,
    Opaque(InlineKind),
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
) -> Child {
    if is_fragment_word_name(namespace, name, b"r", fragment_prefix) {
        return Child::Run;
    }
    match InlineKind::classify(namespace, name, fragment_prefix) {
        InlineKind::BookmarkStart => Child::BookmarkStart,
        InlineKind::BookmarkEnd => Child::BookmarkEnd,
        kind => Child::Opaque(kind),
    }
}

impl Paragraph {
    /// Return every direct paragraph child in source order.
    ///
    /// Paragraph properties are metadata and are not included. Direct
    /// `WordprocessingML` runs become [`Inline::Run`] and bookmark boundaries
    /// become [`Inline::BookmarkStart`] and [`Inline::BookmarkEnd`]; every
    /// other child remains available as [`Inline::Unknown`] with exact inert
    /// XML bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed, unterminated, or resource-exhausting
    /// paragraph XML.
    pub fn inlines(&self) -> Result<Vec<Inline>> {
        let xml = self.xml_bytes();
        let source = self.xml_data.xml_ref()?;
        let _parser_admission = source.parser_admission()?;
        let mut reader = NsReader::from_reader(xml);
        for (prefix, namespace) in self.xml_data.inherited_namespaces()?.iter() {
            let prefix = prefix.as_deref().map_or(
                quick_xml::name::PrefixDeclaration::Default,
                quick_xml::name::PrefixDeclaration::Named,
            );
            reader
                .resolver_mut()
                .add(prefix, quick_xml::name::Namespace(namespace))
                .map_err(|_| Error::Xml("invalid inherited inline namespace".into()))?;
        }
        let origin = ReaderOrigin::of(xml);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut paragraph_depth = None;
        let mut capture = None::<(Child, usize, usize)>;
        let mut inlines = Vec::new();
        let mut namespaces = crate::namespace::NamespaceCapture::default();
        let mut saw_paragraph = false;
        let mut depth = 0usize;
        let mut nodes = 0usize;

        loop {
            let event_start = origin.offset(reader.buffer_position()).ok_or_else(|| {
                Error::InvalidFormat("Word inline offset does not fit usize".into())
            })?;
            let raw_event = reader
                .read_event()
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned();
            let resolver = reader.resolver();
            let (namespace, event) = resolver.resolve_event(raw_event);
            let event_end = origin.offset(reader.buffer_position()).ok_or_else(|| {
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
                        capture = Some((
                            classify(&namespace, element.name(), &fragment_prefix),
                            event_start,
                            depth,
                        ));
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
                        push_inline(
                            &mut inlines,
                            &source,
                            classify(&namespace, element.name(), &fragment_prefix),
                            event_start,
                            event_end,
                        )?;
                        capture_run_namespaces(&mut inlines, &mut namespaces, reader.resolver())?;
                    }
                },
                Event::End(element) => {
                    if let Some((kind, start, capture_depth)) = capture
                        && depth == capture_depth
                    {
                        push_inline(&mut inlines, &source, kind, start, event_end)?;
                        capture_run_namespaces(&mut inlines, &mut namespaces, reader.resolver())?;
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
    source: &XmlRef,
    kind: Child,
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
    let range = source.subrange(relative_start, length).ok_or_else(|| {
        Error::InvalidFormat("Word inline absolute offset exceeds source range".into())
    })?;
    match kind {
        Child::Run => inlines.push(Inline::Run(Box::new(Run::from_xml_ref(range)))),
        Child::BookmarkStart => match parse_bookmark(range.bytes()) {
            Some(bookmark) => inlines.push(Inline::BookmarkStart(bookmark)),
            None => push_unknown(inlines, range, InlineKind::BookmarkStart),
        },
        Child::BookmarkEnd => match parse_bookmark(range.bytes()) {
            Some(bookmark) => inlines.push(Inline::BookmarkEnd(bookmark.id())),
            None => push_unknown(inlines, range, InlineKind::BookmarkEnd),
        },
        Child::Opaque(kind) => push_unknown(inlines, range, kind),
    }
    Ok(())
}

fn push_unknown(inlines: &mut Vec<Inline>, source: XmlRef, kind: InlineKind) {
    inlines.push(Inline::Unknown(Box::new(OpaqueInline::from_xml_ref(
        source, kind,
    ))));
}

/// Read the `w:id` and `w:name` of one bookmark boundary fragment.
///
/// Returns `None` when the boundary has no non-negative identifier, leaving
/// the element to be retained as inert XML.
fn parse_bookmark(raw: &[u8]) -> Option<Bookmark> {
    let mut reader = NsReader::from_reader(raw);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => {
                let mut id = None;
                let mut name = String::new();
                for attr in first_wins(&e).flatten() {
                    match attr.key.local_name().as_ref() {
                        b"id" => {
                            id = atoi_simd::parse::<u32, false, false>(&attr.value).ok();
                        },
                        b"name" => {
                            name = String::from_utf8_lossy(&attr.value).into_owned();
                        },
                        _ => {},
                    }
                }
                return id.map(|id| Bookmark::new(id, name));
            },
            Ok(Event::Eof) => return None,
            Err(_) => return None,
            _ => {},
        }
    }
}

fn capture_run_namespaces(
    inlines: &mut [Inline],
    capture: &mut crate::namespace::NamespaceCapture,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    if let Some(Inline::Run(run)) = inlines.last_mut() {
        run.namespaces = Some(capture.capture(resolver)?);
    }
    Ok(())
}
