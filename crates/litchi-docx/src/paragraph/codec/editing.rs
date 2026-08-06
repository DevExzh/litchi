//! Lossless-ish typed editing of direct paragraph spacing.

use crate::error::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::fmt::Write as _;

use super::super::model::{Paragraph, ParagraphSpacing, XmlData};
use super::xml::{element_prefix, is_fragment_word_name};

#[derive(Clone, Copy)]
struct ByteRange {
    start: usize,
    end: usize,
}

#[derive(Default)]
struct Layout {
    root_prefix: Vec<u8>,
    root_start_end: Option<usize>,
    root_empty: Option<ByteRange>,
    ppr_close_start: Option<usize>,
    ppr_empty: Option<ByteRange>,
    spacing: Option<ByteRange>,
}

impl Paragraph {
    /// Replace the direct spacing properties on this paragraph.
    ///
    /// Some writes a typed CT_Spacing element under pPr and None removes the
    /// direct element. Existing paragraph content and unrelated properties remain
    /// byte-for-byte unchanged. Editing a shared paragraph materializes only
    /// this paragraph; the zero-copy runs path remains unchanged until an edit
    /// is requested.
    pub fn set_spacing(&mut self, spacing: Option<ParagraphSpacing>) -> Result<&mut Self> {
        let original = self.xml_bytes();
        let rewritten = rewrite_spacing(original, spacing)?;
        if rewritten.as_slice() != original {
            self.xml_data = XmlData::Owned(rewritten.into_boxed_slice());
        }
        Ok(self)
    }
}

fn rewrite_spacing(xml_bytes: &[u8], spacing: Option<ParagraphSpacing>) -> Result<Vec<u8>> {
    let layout = locate_layout(xml_bytes)?;
    match (layout.spacing, spacing) {
        (Some(range), Some(spacing)) => {
            let replacement = render_spacing(&layout.root_prefix, spacing)?;
            Ok(splice(xml_bytes, range, replacement.as_bytes()))
        },
        (Some(range), None) => Ok(splice(xml_bytes, range, &[])),
        (None, None) => Ok(xml_bytes.to_vec()),
        (None, Some(spacing)) => {
            let spacing = render_spacing(&layout.root_prefix, spacing)?;
            if let Some(close_start) = layout.ppr_close_start {
                return Ok(insert_at(xml_bytes, close_start, spacing.as_bytes()));
            }
            let spacing_bytes = spacing.as_bytes();
            let ppr_name = qualified_name(&layout.root_prefix, "pPr")?;
            if let Some(range) = layout.ppr_empty {
                let ppr_end = format!("</{ppr_name}>");
                return expand_empty(xml_bytes, range, spacing_bytes, ppr_end.as_bytes());
            }
            let ppr = format!("<{ppr_name}>{}</{ppr_name}>", spacing);
            if let Some(root_start_end) = layout.root_start_end {
                return Ok(insert_at(xml_bytes, root_start_end, ppr.as_bytes()));
            }
            let root_range = layout.root_empty.ok_or_else(|| {
                Error::InvalidFormat("paragraph XML has no editable root".to_owned())
            })?;
            let root_name = qualified_name(&layout.root_prefix, "p")?;
            let root_end = format!("</{root_name}>");
            expand_empty(xml_bytes, root_range, ppr.as_bytes(), root_end.as_bytes())
        },
    }
}

fn locate_layout(xml_bytes: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut layout = Layout::default();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut ppr_depth = None;
    let mut spacing_depth = None;
    let mut spacing_start = None;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph XML offset does not fit usize".into()))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| Error::InvalidFormat("paragraph XML offset does not fit usize".into()))?;

        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".into())
                })?;
                if depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        event_end,
                        None,
                    )?;
                } else {
                    let is_word = is_fragment_word_name(
                        &namespace,
                        element.name(),
                        element.local_name().as_ref(),
                        &fragment_prefix,
                    );
                    if depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                        if ppr_depth.is_some() || layout.ppr_empty.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate pPr".to_owned(),
                            ));
                        }
                        ppr_depth = Some(depth);
                    } else if depth == 3
                        && ppr_depth == Some(2)
                        && is_word
                        && element.local_name().as_ref() == b"spacing"
                    {
                        if layout.spacing.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate spacing".to_owned(),
                            ));
                        }
                        spacing_depth = Some(depth);
                        spacing_start = Some(event_start);
                    }
                }
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("paragraph XML nesting is too deep".into())
                })?;
                if child_depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        event_end,
                        Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        }),
                    )?;
                } else {
                    let is_word = is_fragment_word_name(
                        &namespace,
                        element.name(),
                        element.local_name().as_ref(),
                        &fragment_prefix,
                    );
                    if child_depth == 2 && is_word && element.local_name().as_ref() == b"pPr" {
                        if ppr_depth.is_some() || layout.ppr_empty.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate pPr".to_owned(),
                            ));
                        }
                        layout.ppr_empty = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                    } else if child_depth == 3
                        && ppr_depth == Some(2)
                        && is_word
                        && element.local_name().as_ref() == b"spacing"
                    {
                        if layout.spacing.is_some() {
                            return Err(Error::InvalidFormat(
                                "paragraph has duplicate spacing".to_owned(),
                            ));
                        }
                        layout.spacing = Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        });
                    }
                }
            },
            Event::End(_) => {
                if spacing_depth == Some(depth) {
                    let start = spacing_start.take().ok_or_else(|| {
                        Error::InvalidFormat("paragraph spacing has no start".to_owned())
                    })?;
                    layout.spacing = Some(ByteRange {
                        start,
                        end: event_end,
                    });
                    spacing_depth = None;
                }
                if ppr_depth == Some(depth) {
                    ppr_depth = None;
                    layout.ppr_close_start = Some(event_start);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid paragraph XML nesting".to_owned())
                })?;
            },
            Event::Eof if depth != 0 => {
                return Err(Error::InvalidFormat(
                    "unterminated paragraph XML".to_owned(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !saw_root {
        return Err(Error::InvalidFormat("paragraph XML has no root".to_owned()));
    }
    if spacing_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated paragraph spacing".to_owned(),
        ));
    }
    Ok(layout)
}

fn validate_root(
    layout: &mut Layout,
    fragment_prefix: &mut Option<Option<Vec<u8>>>,
    saw_root: &mut bool,
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    event_end: usize,
    empty: Option<ByteRange>,
) -> Result<()> {
    if !matches!(namespace, ResolveResult::Bound(_)) {
        *fragment_prefix = Some(
            element
                .name()
                .prefix()
                .map(|prefix| prefix.into_inner().to_vec()),
        );
    }
    if *saw_root
        || element.local_name().as_ref() != b"p"
        || !is_fragment_word_name(namespace, element.name(), b"p", fragment_prefix)
    {
        return Err(Error::InvalidFormat(
            "paragraph XML has an invalid root".to_owned(),
        ));
    }
    *saw_root = true;
    layout.root_prefix = element_prefix(element);
    if matches!(namespace, ResolveResult::Bound(_)) {
        layout.root_start_end = Some(event_end);
    } else {
        if empty.is_none() {
            layout.root_start_end = Some(event_end);
        }
    }
    layout.root_empty = empty;
    Ok(())
}

fn qualified_name(prefix: &[u8], local: &str) -> Result<String> {
    let prefix = std::str::from_utf8(prefix)
        .map_err(|error| Error::InvalidFormat(format!("invalid paragraph prefix: {error}")))?;
    if prefix.is_empty() {
        Ok(local.to_owned())
    } else {
        Ok(format!("{prefix}:{local}"))
    }
}

fn render_spacing(prefix: &[u8], spacing: ParagraphSpacing) -> Result<String> {
    let qname = qualified_name(prefix, "spacing")?;
    let attr = |name: &str| qualified_name(prefix, name);
    let mut xml = String::new();
    write!(xml, "<{qname}").map_err(|error| Error::Xml(error.to_string()))?;

    if let Some(value) = spacing.before {
        write!(xml, " {}=\"{value}\"", attr("before")?)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.before_lines {
        write!(xml, " {}=\"{value}\"", attr("beforeLines")?)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.before_auto_spacing {
        write!(xml, " {}=\"{}\"", attr("beforeAutospacing")?, on_off(value))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.after {
        write!(xml, " {}=\"{value}\"", attr("after")?)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.after_lines {
        write!(xml, " {}=\"{value}\"", attr("afterLines")?)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.after_auto_spacing {
        write!(xml, " {}=\"{}\"", attr("afterAutospacing")?, on_off(value))
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.line {
        write!(xml, " {}=\"{value}\"", attr("line")?)
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    if let Some(value) = spacing.line_rule {
        write!(xml, " {}=\"{}\"", attr("lineRule")?, value.as_str())
            .map_err(|error| Error::Xml(error.to_string()))?;
    }
    xml.push_str("/>");
    Ok(xml)
}

fn on_off(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

fn splice(source: &[u8], range: ByteRange, replacement: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(
        source
            .len()
            .saturating_sub(range.end.saturating_sub(range.start))
            .saturating_add(replacement.len()),
    );
    output.extend_from_slice(&source[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&source[range.end..]);
    output
}

fn insert_at(source: &[u8], offset: usize, insertion: &[u8]) -> Vec<u8> {
    splice(
        source,
        ByteRange {
            start: offset,
            end: offset,
        },
        insertion,
    )
}

fn expand_empty(source: &[u8], range: ByteRange, inner: &[u8], closing: &[u8]) -> Result<Vec<u8>> {
    let raw = &source[range.start..range.end];
    let close = raw
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| Error::InvalidFormat("empty paragraph element has no close".to_owned()))?;
    let slash = raw[..close]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .filter(|&index| raw[index] == b'/')
        .ok_or_else(|| {
            Error::InvalidFormat("paragraph empty element is missing '/>'".to_owned())
        })?;

    let mut replacement = Vec::with_capacity(raw.len() + inner.len() + closing.len());
    replacement.extend_from_slice(&raw[..slash]);
    replacement.push(b'>');
    replacement.extend_from_slice(inner);
    replacement.extend_from_slice(closing);
    Ok(splice(source, range, &replacement))
}
