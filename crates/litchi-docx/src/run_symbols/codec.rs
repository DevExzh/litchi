#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Namespace-aware, source-preserving `symEx` XML codec.

use crate::error::{Error, Result};
use crate::paragraph::is_fragment_word_name;
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as FmtWrite;

use super::model::{Symbol, Symbols};
use super::validation::{
    MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, MC_NAMESPACE, SYMEX_NAMESPACE, parse_char,
    validate_symbol, validate_symbols,
};

const SYMEX_PREFIX: &str = "w15sym";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ByteRange {
    start: usize,
    end: usize,
}

#[derive(Debug, Default)]
struct Layout {
    root_empty: Option<ByteRange>,
    root_close_start: Option<usize>,
    root_name: Vec<u8>,
    symbols: Vec<ByteRange>,
}

/// Read all direct `symEx` elements from one complete `w:r` fragment.
pub(crate) fn read(xml: &[u8]) -> Result<Symbols> {
    let (_, symbols) = locate(xml)?;
    Ok(symbols)
}

/// Rewrite only the direct `symEx` seams in one complete `w:r` fragment.
///
/// Existing run children, including foreign and unsupported XML, are copied
/// directly from the source.  New symbols are inserted at the end of the run
/// content, immediately before the root close.
pub(crate) fn rewrite(xml: &[u8], next: &Symbols) -> Result<Vec<u8>> {
    validate_symbols(next)?;
    let (layout, current) = locate(xml)?;
    if current == *next {
        return Ok(xml.to_vec());
    }

    let mut output = xml.to_vec();
    if next.len() > layout.symbols.len() {
        let mut insertion = String::new();
        for symbol in next.iter().skip(layout.symbols.len()) {
            write_symbol(symbol, &mut insertion)?;
        }
        if let Some(offset) = layout.root_close_start {
            output = insert_at(&output, offset, insertion.as_bytes());
        } else if let Some(range) = layout.root_empty {
            output = expand_empty(&output, range, insertion.as_bytes(), &layout.root_name)?;
        } else {
            return Err(Error::InvalidFormat(
                "Word run XML has no insertion point for symEx".into(),
            ));
        }
    }

    for index in (0..layout.symbols.len()).rev() {
        let replacement = if let Some(symbol) = next.get(index) {
            let mut rendered = String::new();
            write_symbol(symbol, &mut rendered)?;
            rendered.into_bytes()
        } else {
            Vec::new()
        };
        output = splice(&output, layout.symbols[index], &replacement);
    }
    Ok(output)
}

/// Append one canonical `symEx` element to generated run XML.
pub(crate) fn write_symbol(value: &Symbol, output: &mut String) -> Result<()> {
    validate_symbol(value)?;
    write!(
        output,
        r#"<{SYMEX_PREFIX}:symEx xmlns:{SYMEX_PREFIX}="{}" xmlns:mc="{}" mc:Ignorable="{SYMEX_PREFIX}""#,
        String::from_utf8_lossy(SYMEX_NAMESPACE),
        MC_NAMESPACE,
    )?;
    if let Some(font) = value.font_value() {
        write!(output, r#" {SYMEX_PREFIX}:font="{}""#, escape_xml(font))?;
    }
    if let Some(character) = value.character_value() {
        write!(output, r#" {SYMEX_PREFIX}:char="{character:08X}""#)?;
    }
    output.push_str("/>");
    Ok(())
}

fn locate(xml: &[u8]) -> Result<(Layout, Symbols)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "Word run XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;

    let mut layout = Layout::default();
    let mut symbols = Symbols::new();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;
    let mut symbol_depth = None;
    let mut symbol_start = None;

    loop {
        let event_start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = offset(&reader)?;
        let resolver = reader.resolver().clone();
        let decoder = reader.decoder();
        let (namespace, event) = resolver.resolve_event(event);

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Word run XML element counter overflows usize".into())
            })?;
            if nodes > MAX_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "Word run XML exceeds {MAX_XML_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word run XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        None,
                    )?;
                } else if symbol_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "Word symEx cannot contain child elements".into(),
                    ));
                } else if child_depth == 2
                    && is_symex_namespace(&namespace)
                    && element.local_name().as_ref() == b"symEx"
                {
                    symbols.push(parse_symbol(&element, decoder, &resolver)?)?;
                    symbol_depth = Some(child_depth);
                    symbol_start = Some(event_start);
                }
                depth = child_depth;
            },
            Event::Empty(element) => {
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("Word run XML nesting is too deep".into())
                })?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "Word run XML nesting exceeds {MAX_XML_DEPTH}"
                    )));
                }
                if child_depth == 1 {
                    validate_root(
                        &mut layout,
                        &mut fragment_prefix,
                        &mut saw_root,
                        &namespace,
                        &element,
                        Some(ByteRange {
                            start: event_start,
                            end: event_end,
                        }),
                    )?;
                    root_closed = true;
                } else if symbol_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "Word symEx cannot contain child elements".into(),
                    ));
                } else if child_depth == 2
                    && is_symex_namespace(&namespace)
                    && element.local_name().as_ref() == b"symEx"
                {
                    layout.symbols.push(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                    symbols.push(parse_symbol(&element, decoder, &resolver)?)?;
                }
            },
            Event::End(_) => {
                if symbol_depth == Some(depth) {
                    let start = symbol_start.take().ok_or_else(|| {
                        Error::InvalidFormat("Word symEx has no start offset".into())
                    })?;
                    layout.symbols.push(ByteRange {
                        start,
                        end: event_end,
                    });
                    symbol_depth = None;
                }
                if depth == 1 {
                    root_closed = true;
                    layout.root_close_start = Some(event_start);
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("invalid Word run XML nesting".into()))?;
            },
            Event::Text(text) if symbol_depth.is_some() => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word symEx cannot contain character data".into(),
                    ));
                }
            },
            Event::CData(text) if symbol_depth.is_some() => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word symEx cannot contain character data".into(),
                    ));
                }
            },
            Event::Text(text) if depth == 0 => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word run XML has text outside its root".into(),
                    ));
                }
            },
            Event::CData(text) if depth == 0 => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word run XML has character data outside its root".into(),
                    ));
                }
            },
            Event::Eof if depth != 0 || symbol_depth.is_some() || !root_closed => {
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

    if !saw_root {
        return Err(Error::InvalidFormat("Word run XML has no root".into()));
    }
    Ok((layout, symbols))
}

fn validate_root(
    layout: &mut Layout,
    fragment_prefix: &mut Option<Option<Vec<u8>>>,
    saw_root: &mut bool,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    empty: Option<ByteRange>,
) -> Result<()> {
    if *saw_root {
        return Err(Error::InvalidFormat(
            "Word run XML has multiple roots".into(),
        ));
    }
    if !matches!(namespace, ResolveResult::Bound(_)) {
        *fragment_prefix = Some(
            element
                .name()
                .prefix()
                .map(|prefix| prefix.into_inner().to_vec()),
        );
    }
    if element.local_name().as_ref() != b"r"
        || !is_fragment_word_name(namespace, element.name(), b"r", fragment_prefix)
    {
        return Err(Error::InvalidFormat(
            "Word run XML has an invalid root".into(),
        ));
    }
    *saw_root = true;
    layout.root_name = element.name().as_ref().to_vec();
    if let Some(empty) = empty {
        layout.root_empty = Some(empty);
    }
    Ok(())
}

fn parse_symbol(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Symbol> {
    let mut font = None;
    let mut character = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let local = attribute.key.local_name();
        if local.as_ref() != b"font" && local.as_ref() != b"char" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_symex_attribute(&namespace, element) {
            return Err(Error::InvalidFormat(format!(
                "Word symEx attribute '{}' is not in the symex namespace",
                String::from_utf8_lossy(local.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        match local.as_ref() {
            b"font" => {
                if font.replace(value).is_some() {
                    return Err(Error::InvalidFormat(
                        "Word symEx has duplicate font attributes".into(),
                    ));
                }
            },
            b"char" => {
                if character.is_some() {
                    return Err(Error::InvalidFormat(
                        "Word symEx has duplicate char attributes".into(),
                    ));
                }
                character = Some(parse_char(&value)?);
            },
            _ => unreachable!(),
        }
    }
    Symbol::from_parts(font, character)
}

fn is_symex_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == SYMEX_NAMESPACE
    )
}

fn is_symex_attribute(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_symex_namespace(namespace)
        || matches!(namespace, ResolveResult::Unbound)
        || matches!(
            namespace,
            ResolveResult::Unknown(prefix)
                if element
                    .name()
                    .prefix()
                    .is_some_and(|element_prefix| element_prefix.as_ref() == prefix.as_slice())
        )
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position()).map_err(|_source_error| {
        Error::InvalidFormat("Word run XML offset does not fit usize".into())
    })
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

fn expand_empty(
    source: &[u8],
    range: ByteRange,
    inner: &[u8],
    root_name: &[u8],
) -> Result<Vec<u8>> {
    let raw = &source[range.start..range.end];
    let close = raw
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| Error::InvalidFormat("empty Word run has no close".into()))?;
    let slash = raw[..close]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .filter(|&index| raw[index] == b'/')
        .ok_or_else(|| Error::InvalidFormat("empty Word run is missing '/>'".into()))?;

    let mut replacement = Vec::with_capacity(raw.len() + inner.len() + root_name.len() + 3);
    replacement.extend_from_slice(&raw[..slash]);
    replacement.push(b'>');
    replacement.extend_from_slice(inner);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(root_name);
    replacement.push(b'>');
    Ok(splice(source, range, &replacement))
}
