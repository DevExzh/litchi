#![expect(
    clippy::expect_used,
    reason = "the invariant is established immediately before extraction"
)]
#![expect(
    clippy::option_option,
    reason = "nested options distinguish omitted, present-empty, and present-valued XML"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
#![expect(
    clippy::unnecessary_wraps,
    reason = "the Result signature preserves a uniform fallible codec API"
)]
//! Bounded, namespace-aware, loss-preserving `sectPr` codec for
//! `w12:footnoteColumns`.

use crate::error::{Error, Result};
use crate::namespace::is_wordprocessing_namespace;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration, QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write;

use super::model::Layout;
use super::validation::{
    self, has_ignorable_prefix, is_inherited_markup_compatibility, is_word_2012_element,
    is_word_value_attribute,
};

const ROOT: &[u8] = b"sectPr";
const EXTENSION: &[u8] = b"footnoteColumns";
const VALUE: &[u8] = b"val";

/// Namespace and markup-compatibility state inherited by a detached section
/// fragment. The source bytes stay untouched; this context is only used to
/// resolve names that were declared on an enclosing document element.
#[derive(Debug, Clone, Default, Eq, PartialEq)]
pub(crate) struct Context {
    bindings: Vec<Binding>,
    ignorable: Option<String>,
}

#[derive(Debug, Clone, Eq, PartialEq)]
struct Binding {
    prefix: Option<Vec<u8>>,
    namespace: Vec<u8>,
}

impl Context {
    /// Capture the currently effective bindings without retaining the reader.
    pub(crate) fn from_resolver(
        resolver: &NamespaceResolver,
        ignorable: Option<String>,
    ) -> Result<Self> {
        let mut bindings = Vec::new();
        let mut bytes = ignorable.as_ref().map_or(0, String::len);
        for (prefix, namespace) in resolver.bindings() {
            if bindings.len() >= validation::MAX_CONTEXT_BINDINGS {
                return Err(Error::InvalidFormat(format!(
                    "section namespace context exceeds {} bindings",
                    validation::MAX_CONTEXT_BINDINGS
                )));
            }
            let prefix = match prefix {
                PrefixDeclaration::Default => None,
                PrefixDeclaration::Named(prefix) => Some(prefix.to_vec()),
            };
            bytes = bytes
                .checked_add(prefix.as_ref().map_or(0, Vec::len))
                .and_then(|bytes| bytes.checked_add(namespace.as_ref().len()))
                .ok_or_else(|| {
                    Error::InvalidFormat("section namespace context size overflow".into())
                })?;
            if bytes > validation::MAX_CONTEXT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "section namespace context exceeds {} bytes",
                    validation::MAX_CONTEXT_BYTES
                )));
            }
            bindings.push(Binding {
                prefix,
                namespace: namespace.as_ref().to_vec(),
            });
        }
        Ok(Self {
            bindings,
            ignorable,
        })
    }

    fn install(&self, reader: &mut NsReader<&[u8]>) -> Result<()> {
        for binding in &self.bindings {
            let prefix = match binding.prefix.as_deref() {
                None => PrefixDeclaration::Default,
                Some(prefix) => PrefixDeclaration::Named(prefix),
            };
            reader
                .resolver_mut()
                .add(prefix, Namespace(&binding.namespace))
                .map_err(|error| Error::InvalidFormat(error.to_string()))?;
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct Range {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
pub(crate) struct Parsed {
    pub(crate) value: Option<Layout>,
    root: Root,
    extension: Option<Range>,
    value_range: Option<Range>,
    sect_pr_change_start: Option<usize>,
    unsafe_insertion_child: bool,
}

#[derive(Debug, Clone)]
struct Root {
    start: usize,
    open_end: usize,
    end: usize,
    close_start: Option<usize>,
    name: Vec<u8>,
    open: Vec<u8>,
    word_prefix: Vec<u8>,
    word_namespace: Vec<u8>,
    resolver: NamespaceResolver,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Vec<u8>,
    value: Vec<u8>,
    value_start: usize,
    value_end: usize,
}

/// Parse a section fragment using namespace state captured from its owning
/// document part. The supplied context is never serialized into the source
/// snapshot.
pub(crate) fn read_with_context(xml: &[u8], context: &Context) -> Result<Parsed> {
    if xml.len() > validation::MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "section XML exceeds {} bytes",
            validation::MAX_XML_BYTES
        )));
    }

    let mut reader = NsReader::from_reader(xml);
    context.install(&mut reader)?;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut root = None;
    let mut extension = None;
    let mut extension_prefix = None;
    let mut value_prefix = None;
    let mut value_range = None;
    let mut extension_depth = None;
    let mut extension_start = None;
    let mut root_ignorable = None;
    let mut sect_pr_change_start = None;
    let mut unsafe_insertion_child = false;

    loop {
        let event_start = offset(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = offset(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("section XML element counter overflow".into())
            })?;
            if nodes > validation::MAX_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "section XML exceeds {} elements",
                    validation::MAX_XML_NODES
                )));
            }
        }

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !is_root(&namespace, &element, &mut fragment_prefix)
                    {
                        return Err(Error::InvalidFormat(
                            "section XML has an invalid sectPr root".into(),
                        ));
                    }
                    root_seen = true;
                    root_ignorable = root_ignorable_value(&element, &resolver, decoder)?
                        .or_else(|| context.ignorable.clone());
                    root = Some(Root {
                        start: event_start,
                        open_end: event_end,
                        end: 0,
                        close_start: None,
                        name: element.name().as_ref().to_vec(),
                        open: xml[event_start..event_end].to_vec(),
                        word_prefix: element_prefix(&element).unwrap_or_else(|| b"w".to_vec()),
                        word_namespace: namespace_uri(&namespace).unwrap_or_else(|| {
                            crate::namespace::WORDPROCESSINGML_NAMESPACE.to_vec()
                        }),
                        resolver: resolver.clone(),
                    });
                    depth = 1;
                } else {
                    if extension_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "footnoteColumns cannot contain child elements".into(),
                        ));
                    }
                    let is_extension = depth == 1
                        && is_extension_element(&namespace, &element)
                        && element.local_name().as_ref() == EXTENSION;
                    if depth == 1 {
                        observe_direct_word_child(
                            &namespace,
                            &element,
                            fragment_prefix.as_ref(),
                            event_start,
                            &mut sect_pr_change_start,
                            &mut unsafe_insertion_child,
                        );
                    }
                    if is_extension {
                        if extension.is_some() || extension_depth.is_some() {
                            return Err(Error::InvalidFormat(
                                "section has duplicate footnoteColumns elements".into(),
                            ));
                        }
                        let (layout, attribute_prefix) =
                            parse_extension(&element, &resolver, decoder, None)?;
                        let prefix = element_prefix(&element);
                        let value_span =
                            find_value_range(xml, event_start, event_end, &attribute_prefix)?;
                        require_ignorable(
                            root_ignorable.as_deref(),
                            prefix.as_deref().unwrap_or_default(),
                            matches!(namespace, ResolveResult::Unknown(_)),
                        )?;
                        extension_depth = Some(2);
                        extension_start = Some(event_start);
                        extension_prefix = prefix;
                        value_prefix = Some(attribute_prefix);
                        value_range = Some(value_span);
                        // The value is assigned when the matching end event
                        // closes the extension, so a malformed nested event
                        // cannot publish a partial state.
                        let _ = layout;
                    }
                    depth = depth.checked_add(1).ok_or_else(too_deep)?;
                    if depth > validation::MAX_XML_DEPTH {
                        return Err(too_deep());
                    }
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !is_root(&namespace, &element, &mut fragment_prefix)
                    {
                        return Err(Error::InvalidFormat(
                            "section XML has an invalid sectPr root".into(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                    root_ignorable = root_ignorable_value(&element, &resolver, decoder)?
                        .or_else(|| context.ignorable.clone());
                    root = Some(Root {
                        start: event_start,
                        open_end: event_end,
                        end: event_end,
                        close_start: None,
                        name: element.name().as_ref().to_vec(),
                        open: xml[event_start..event_end].to_vec(),
                        word_prefix: element_prefix(&element).unwrap_or_else(|| b"w".to_vec()),
                        word_namespace: namespace_uri(&namespace).unwrap_or_else(|| {
                            crate::namespace::WORDPROCESSINGML_NAMESPACE.to_vec()
                        }),
                        resolver: resolver.clone(),
                    });
                } else if extension_depth.is_some() {
                    return Err(Error::InvalidFormat(
                        "footnoteColumns cannot contain child elements".into(),
                    ));
                } else if depth == 1
                    && is_extension_element(&namespace, &element)
                    && element.local_name().as_ref() == EXTENSION
                {
                    if extension.is_some() {
                        return Err(Error::InvalidFormat(
                            "section has duplicate footnoteColumns elements".into(),
                        ));
                    }
                    let (layout, attribute_prefix) =
                        parse_extension(&element, &resolver, decoder, None)?;
                    let prefix = element_prefix(&element);
                    let value_span =
                        find_value_range(xml, event_start, event_end, &attribute_prefix)?;
                    require_ignorable(
                        root_ignorable.as_deref(),
                        prefix.as_deref().unwrap_or_default(),
                        matches!(namespace, ResolveResult::Unknown(_)),
                    )?;
                    extension = Some(Range {
                        start: event_start,
                        end: event_end,
                    });
                    extension_prefix = prefix;
                    value_prefix = Some(attribute_prefix);
                    value_range = Some(value_span);
                    let _ = layout;
                } else if depth == 1 {
                    observe_direct_word_child(
                        &namespace,
                        &element,
                        fragment_prefix.as_ref(),
                        event_start,
                        &mut sect_pr_change_start,
                        &mut unsafe_insertion_child,
                    );
                }
            },
            Event::Text(text)
                if extension_depth.is_some()
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::InvalidFormat(
                    "footnoteColumns cannot contain text".into(),
                ));
            },
            Event::CData(text)
                if extension_depth.is_some()
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(Error::InvalidFormat(
                    "footnoteColumns cannot contain text".into(),
                ));
            },
            Event::End(element) => {
                if extension_depth == Some(depth) {
                    if element.local_name().as_ref() != EXTENSION {
                        return Err(Error::InvalidFormat(
                            "footnoteColumns has mismatched XML nesting".into(),
                        ));
                    }
                    let start = extension_start.take().ok_or_else(|| {
                        Error::InvalidFormat("footnoteColumns range is missing".into())
                    })?;
                    extension = Some(Range {
                        start,
                        end: event_end,
                    });
                    extension_depth = None;
                }
                if depth == 1 {
                    if !root_seen || root_closed {
                        return Err(Error::InvalidFormat(
                            "section XML has an invalid sectPr close".into(),
                        ));
                    }
                    root_closed = true;
                    let value = root.as_mut().ok_or_else(|| {
                        Error::InvalidFormat("section XML root state is missing".into())
                    })?;
                    value.end = event_end;
                    value.close_start = Some(event_start);
                }
                depth = depth.checked_sub(1).ok_or_else(too_deep)?;
            },
            Event::Text(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "section XML has text outside sectPr".into(),
                    ));
                }
            },
            Event::CData(text) if depth == 0 => {
                if !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "section XML has text outside sectPr".into(),
                    ));
                }
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

    if !root_seen || !root_closed || depth != 0 || extension_depth.is_some() {
        return Err(Error::InvalidFormat(
            "section XML has an incomplete sectPr".into(),
        ));
    }
    let root = root.ok_or_else(|| Error::InvalidFormat("section XML has no sectPr".into()))?;
    let value = extension
        .map(|range| {
            parse_extension_range(
                xml,
                range,
                extension_prefix.as_deref(),
                value_prefix.as_deref(),
            )
        })
        .transpose()?
        .flatten();
    Ok(Parsed {
        value,
        root,
        extension,
        value_range,
        sect_pr_change_start,
        unsafe_insertion_child,
    })
}

/// Rewrite only the extension seam using inherited package context.
pub(crate) fn rewrite_with_context(
    xml: &[u8],
    value: Option<Layout>,
    context: &Context,
) -> Result<Vec<u8>> {
    validation::validate_layout(value)?;
    let parsed = read_with_context(xml, context)?;
    let Some(range) = parsed.extension else {
        return match value {
            None => Ok(xml.to_vec()),
            Some(value) => insert_extension(xml, &parsed, value),
        };
    };
    match value {
        Some(value) => replace_value(xml, parsed.value_range, value),
        None => Ok(splice(xml, range, &[])),
    }
}

fn insert_extension(xml: &[u8], parsed: &Parsed, value: Layout) -> Result<Vec<u8>> {
    if parsed.unsafe_insertion_child {
        return Err(Error::InvalidFormat(
            "cannot prove a schema-safe footnoteColumns insertion point".into(),
        ));
    }
    let root = &parsed.root;
    let (element_prefix, open) = prepare_root(
        xml,
        &root.open,
        &root.word_prefix,
        &root.word_namespace,
        &root.resolver,
    )?;
    let rendered = render(&element_prefix, &root.word_prefix, value)?;
    let Some(close_start) = root.close_start else {
        let mut expanded = expand_root(&open, &root.name)?;
        expanded.extend_from_slice(&rendered);
        expanded.extend_from_slice(&close_name(&root.name));
        let mut output = Vec::with_capacity(xml.len() + rendered.len() + 32);
        output.extend_from_slice(&xml[..root.start]);
        output.extend_from_slice(&expanded);
        output.extend_from_slice(&xml[root.end..]);
        return Ok(output);
    };

    let insertion = parsed
        .sect_pr_change_start
        .unwrap_or_else(|| trailing_whitespace_start(xml, root.open_end, close_start));
    let mut output = Vec::with_capacity(xml.len() + rendered.len() + 32);
    output.extend_from_slice(&xml[..root.start]);
    output.extend_from_slice(&open);
    output.extend_from_slice(&xml[root.open_end..insertion]);
    output.extend_from_slice(&rendered);
    output.extend_from_slice(&xml[insertion..]);
    Ok(output)
}

fn replace_value(xml: &[u8], range: Option<Range>, value: Layout) -> Result<Vec<u8>> {
    let range = range.ok_or_else(|| {
        Error::InvalidFormat("footnoteColumns val source range is missing".into())
    })?;
    let replacement = value.columns().to_string();
    Ok(splice(xml, range, replacement.as_bytes()))
}

fn parse_extension(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
    inherited_value_prefix: Option<&[u8]>,
) -> Result<(Layout, Vec<u8>)> {
    let mut value = None;
    let mut prefix = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != VALUE {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        let attribute_prefix = attribute
            .key
            .prefix()
            .map(|prefix| prefix.into_inner().to_vec());
        let inherited = matches!(
            (&namespace, inherited_value_prefix),
            (ResolveResult::Unknown(value), Some(prefix)) if value.as_slice() == prefix
        );
        if !is_word_value_attribute(&namespace, attribute_prefix.as_deref()) && !inherited {
            return Err(Error::InvalidFormat(
                "footnoteColumns val is not in the WordprocessingML namespace".into(),
            ));
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "footnoteColumns has duplicate val attributes".into(),
            ));
        }
        prefix = attribute
            .key
            .prefix()
            .map(|prefix| prefix.into_inner().to_vec());
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    let value = value
        .ok_or_else(|| Error::InvalidFormat("footnoteColumns is missing required val".into()))?;
    let prefix = prefix.or_else(|| element_prefix(element));
    let prefix = prefix.unwrap_or_else(|| b"w12".to_vec());
    Ok((validation::parse_columns(&value)?, prefix))
}

fn find_value_range(
    xml: &[u8],
    event_start: usize,
    event_end: usize,
    prefix: &[u8],
) -> Result<Range> {
    let tag = xml.get(event_start..event_end).ok_or_else(|| {
        Error::InvalidFormat("footnoteColumns value range is outside the section".into())
    })?;
    let attributes = scan_attributes(tag)?;
    let mut expected = Vec::with_capacity(prefix.len() + 1 + VALUE.len());
    expected.extend_from_slice(prefix);
    expected.push(b':');
    expected.extend_from_slice(VALUE);
    let attribute = attributes
        .iter()
        .find(|attribute| attribute.name == expected)
        .ok_or_else(|| {
            Error::InvalidFormat("footnoteColumns val source range is missing".into())
        })?;
    let start = event_start
        .checked_add(attribute.value_start)
        .ok_or_else(|| Error::InvalidFormat("footnoteColumns value offset overflow".into()))?;
    let end = event_start
        .checked_add(attribute.value_end)
        .ok_or_else(|| Error::InvalidFormat("footnoteColumns value offset overflow".into()))?;
    Ok(Range { start, end })
}

fn parse_extension_range(
    xml: &[u8],
    range: Range,
    element_prefix: Option<&[u8]>,
    value_prefix: Option<&[u8]>,
) -> Result<Option<Layout>> {
    let mut reader = NsReader::from_reader(&xml[range.start..range.end]);
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element)
                if (is_extension_element(&namespace, &element)
                    || matches!(
                        (&namespace, element_prefix),
                        (ResolveResult::Unknown(value), Some(prefix))
                            if value.as_slice() == prefix
                    ))
                    && element.local_name().as_ref() == EXTENSION =>
            {
                return Ok(Some(
                    parse_extension(&element, &resolver, decoder, value_prefix)?.0,
                ));
            },
            Event::Eof => break,
            Event::Start(_)
            | Event::End(_)
            | Event::Empty(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Err(Error::InvalidFormat(
        "footnoteColumns range has no extension element".into(),
    ))
}

fn root_ignorable_value(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"Ignorable" {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !is_inherited_markup_compatibility(&namespace) {
            continue;
        }
        if value.is_some() {
            return Err(Error::InvalidFormat(
                "sectPr has duplicate mc:Ignorable attributes".into(),
            ));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }
    Ok(value)
}

fn require_ignorable(value: Option<&str>, prefix: &[u8], inherited: bool) -> Result<()> {
    if prefix.is_empty() {
        return Err(Error::InvalidFormat(
            "footnoteColumns must use a prefixed Word 2012 element".into(),
        ));
    }
    if value.is_none_or(|value| !has_ignorable_prefix(value, prefix))
        && !(inherited && prefix == b"w12")
    {
        return Err(Error::InvalidFormat(
            "sectPr mc:Ignorable does not include the footnoteColumns namespace".into(),
        ));
    }
    Ok(())
}

fn prepare_root(
    source_xml: &[u8],
    root_open: &[u8],
    word_prefix: &[u8],
    word_namespace: &[u8],
    resolver: &NamespaceResolver,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let attributes = scan_attributes(root_open)?;
    let element_prefix = choose_prefix(
        &attributes,
        resolver,
        source_xml,
        validation::WORD_2012_NAMESPACE,
        b"w12",
    )?;
    let compatibility_prefix = choose_prefix(
        &attributes,
        resolver,
        source_xml,
        validation::MC_NAMESPACE,
        b"mc",
    )?;
    let mut edits = Vec::new();
    if !has_namespace(&attributes, resolver, word_prefix, word_namespace) {
        edits.push((
            close_insert(root_open),
            format!(" xmlns:{}=\"{}\"", text(word_prefix), text(word_namespace)).into_bytes(),
        ));
    }
    let has_word_decl = has_namespace(
        &attributes,
        resolver,
        &element_prefix,
        validation::WORD_2012_NAMESPACE,
    );
    if !has_word_decl {
        edits.push((
            close_insert(root_open),
            format!(
                " xmlns:{}=\"{}\"",
                text(&element_prefix),
                text(validation::WORD_2012_NAMESPACE)
            )
            .into_bytes(),
        ));
    }
    let has_mc_decl = has_namespace(
        &attributes,
        resolver,
        &compatibility_prefix,
        validation::MC_NAMESPACE,
    );
    if !has_mc_decl {
        edits.push((
            close_insert(root_open),
            format!(
                " xmlns:{}=\"{}\"",
                text(&compatibility_prefix),
                text(validation::MC_NAMESPACE)
            )
            .into_bytes(),
        ));
    }

    let ignorable = find_ignorable_attribute(&attributes, resolver)?;
    if let Some(ignorable) = ignorable {
        if !has_ignorable_prefix_bytes(&ignorable.value, &element_prefix) {
            edits.push((
                ignorable.value_end,
                format!(" {}", text(&element_prefix)).into_bytes(),
            ));
        }
    } else {
        edits.push((
            close_insert(root_open),
            format!(
                " {}:Ignorable=\"{}\"",
                text(&compatibility_prefix),
                text(&element_prefix)
            )
            .into_bytes(),
        ));
    }

    edits.sort_by_key(|(position, _)| *position);
    let mut output = root_open.to_vec();
    for (position, bytes) in edits.into_iter().rev() {
        output.splice(position..position, bytes);
    }
    Ok((element_prefix, output))
}

fn scan_attributes(tag: &[u8]) -> Result<Vec<Attribute>> {
    let mut index = 1usize;
    skip_name(tag, &mut index)?;
    let mut output = Vec::new();
    while index < tag.len() {
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= tag.len() || tag[index] == b'>' || tag[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < tag.len()
            && !tag[index].is_ascii_whitespace()
            && !matches!(tag[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        if name_start == index {
            return Err(Error::InvalidFormat("invalid sectPr root attribute".into()));
        }
        let name = tag[name_start..index].to_vec();
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        if tag.get(index) != Some(&b'=') {
            return Err(Error::InvalidFormat(
                "sectPr root attribute has no value".into(),
            ));
        }
        index += 1;
        while index < tag.len() && tag[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *tag
            .get(index)
            .ok_or_else(|| Error::InvalidFormat("sectPr root attribute is unterminated".into()))?;
        if !matches!(quote, b'"' | b'\'') {
            return Err(Error::InvalidFormat(
                "sectPr root attribute value is not quoted".into(),
            ));
        }
        index += 1;
        let value_start = index;
        while index < tag.len() && tag[index] != quote {
            index += 1;
        }
        let value_end = index;
        if index >= tag.len() {
            return Err(Error::InvalidFormat(
                "sectPr root attribute is unterminated".into(),
            ));
        }
        output.push(Attribute {
            name,
            value: tag[value_start..value_end].to_vec(),
            value_start,
            value_end,
        });
        index += 1;
    }
    Ok(output)
}

fn skip_name(bytes: &[u8], index: &mut usize) -> Result<()> {
    while *index < bytes.len()
        && !bytes[*index].is_ascii_whitespace()
        && !matches!(bytes[*index], b'>' | b'/')
    {
        *index += 1;
    }
    if *index == 1 {
        Err(Error::InvalidFormat(
            "sectPr root has no element name".into(),
        ))
    } else {
        Ok(())
    }
}

fn choose_prefix(
    attributes: &[Attribute],
    resolver: &NamespaceResolver,
    source_xml: &[u8],
    namespace: &[u8],
    preferred: &[u8],
) -> Result<Vec<u8>> {
    if let Some(prefix) = resolver.bindings().find_map(|(prefix, value)| {
        let PrefixDeclaration::Named(prefix) = prefix else {
            return None;
        };
        let prefix = prefix.to_vec();
        (value.as_ref() == namespace && has_namespace(attributes, resolver, &prefix, namespace))
            .then_some(prefix)
    }) {
        return Ok(prefix);
    }

    let mut candidate = preferred.to_vec();
    let mut suffix = 1u32;
    while prefix_is_in_use(attributes, resolver, source_xml, &candidate) {
        candidate = format!("{}_{}", text(preferred), suffix).into_bytes();
        suffix = suffix.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("namespace prefix candidate counter overflow".into())
        })?;
    }
    Ok(candidate)
}

fn has_namespace(
    attributes: &[Attribute],
    resolver: &NamespaceResolver,
    prefix: &[u8],
    namespace: &[u8],
) -> bool {
    if attributes.iter().any(|attribute| {
        attribute.name == [b"xmlns:".as_slice(), prefix].concat() && attribute.value == namespace
    }) {
        return true;
    }
    let qualified = [prefix, b":__litchi_namespace_probe"].concat();
    matches!(
        resolver.resolve_attribute(QName(&qualified)).0,
        ResolveResult::Bound(Namespace(value)) if value == namespace
    )
}

fn find_ignorable_attribute<'a>(
    attributes: &'a [Attribute],
    resolver: &NamespaceResolver,
) -> Result<Option<&'a Attribute>> {
    let mut found = None;
    for attribute in attributes {
        if !attribute.name.ends_with(b":Ignorable") {
            continue;
        }
        let is_compatibility = matches!(
            resolver.resolve_attribute(QName(&attribute.name)).0,
            ResolveResult::Bound(Namespace(value)) if value == validation::MC_NAMESPACE
        );
        if !is_compatibility {
            continue;
        }
        if found.is_some() {
            return Err(Error::InvalidFormat(
                "sectPr has duplicate mc:Ignorable attributes".into(),
            ));
        }
        found = Some(attribute);
    }
    Ok(found)
}

fn prefix_is_in_use(
    attributes: &[Attribute],
    resolver: &NamespaceResolver,
    source_xml: &[u8],
    prefix: &[u8],
) -> bool {
    if resolver
        .bindings()
        .any(|(binding, _)| matches!(binding, PrefixDeclaration::Named(value) if value == prefix))
    {
        return true;
    }
    attributes.iter().any(|attribute| {
        if attribute.name == [b"xmlns:".as_slice(), prefix].concat() {
            return true;
        }
        attribute
            .name
            .iter()
            .position(|byte| *byte == b':')
            .is_some_and(|index| &attribute.name[..index] == prefix)
    }) || qualified_prefix_occurs(source_xml, prefix)
}

fn qualified_prefix_occurs(xml: &[u8], prefix: &[u8]) -> bool {
    let Some(width) = prefix.len().checked_add(1) else {
        return true;
    };
    xml.windows(width)
        .any(|window| window[..prefix.len()] == *prefix && window[prefix.len()] == b':')
}

fn observe_direct_word_child(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    fragment_prefix: Option<&Option<Vec<u8>>>,
    event_start: usize,
    sect_pr_change_start: &mut Option<usize>,
    unsafe_insertion_child: &mut bool,
) {
    if !is_fragment_word_element(namespace, element, fragment_prefix) {
        return;
    }
    let local = element.local_name();
    if local.as_ref() == b"sectPrChange" {
        if sect_pr_change_start.replace(event_start).is_some() {
            *unsafe_insertion_child = true;
        }
    } else if sect_pr_change_start.is_some() || !is_pre_extension_section_child(local.as_ref()) {
        *unsafe_insertion_child = true;
    }
}

fn is_fragment_word_element(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    fragment_prefix: Option<&Option<Vec<u8>>>,
) -> bool {
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix.and_then(|value| value.as_deref()) == Some(prefix.as_slice())
                && element.name().prefix().is_some()
        },
        ResolveResult::Unbound => {
            fragment_prefix.is_some_and(Option::is_none) && element.name().prefix().is_none()
        },
        ResolveResult::Bound(_) => false,
    }
}

fn is_pre_extension_section_child(local: &[u8]) -> bool {
    matches!(
        local,
        b"headerReference"
            | b"footerReference"
            | b"footnotePr"
            | b"endnotePr"
            | b"type"
            | b"pgSz"
            | b"pgMar"
            | b"paperSrc"
            | b"pgBorders"
            | b"lnNumType"
            | b"pgNumType"
            | b"cols"
            | b"formProt"
            | b"vAlign"
            | b"noEndnote"
            | b"titlePg"
            | b"textDirection"
            | b"bidi"
            | b"rtlGutter"
            | b"docGrid"
            | b"printerSettings"
    )
}

fn has_ignorable_prefix_bytes(value: &[u8], prefix: &[u8]) -> bool {
    value
        .split(u8::is_ascii_whitespace)
        .any(|item| item == prefix)
}

fn render(element_prefix: &[u8], value_prefix: &[u8], value: Layout) -> Result<Vec<u8>> {
    let mut xml = String::new();
    write!(
        xml,
        "<{}:footnoteColumns {}:val=\"{}\"/>",
        text(element_prefix),
        text(value_prefix),
        value.columns()
    )
    .map_err(|error| Error::Xml(error.to_string()))?;
    Ok(xml.into_bytes())
}

fn is_root(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    fragment_prefix: &mut Option<Option<Vec<u8>>>,
) -> bool {
    if element.local_name().as_ref() != ROOT {
        return false;
    }
    if is_wordprocessing_namespace(namespace) {
        return true;
    }
    if fragment_prefix.is_none() {
        *fragment_prefix = Some(
            element
                .name()
                .prefix()
                .map(|prefix| prefix.into_inner().to_vec()),
        );
    }
    match namespace {
        ResolveResult::Unknown(prefix) => {
            fragment_prefix
                .as_ref()
                .and_then(|prefix| prefix.as_deref())
                == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
        ResolveResult::Bound(_) => false,
    }
}

fn element_prefix(element: &BytesStart<'_>) -> Option<Vec<u8>> {
    element
        .name()
        .prefix()
        .map(|prefix| prefix.into_inner().to_vec())
}

fn is_extension_element(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    let prefix = element_prefix(element);
    is_word_2012_element(namespace, prefix.as_deref())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_source_error| Error::InvalidFormat("section XML offset overflow".into()))
}

fn too_deep() -> Error {
    Error::InvalidFormat(format!(
        "section XML nesting exceeds {}",
        validation::MAX_XML_DEPTH
    ))
}

fn trailing_whitespace_start(xml: &[u8], start: usize, end: usize) -> usize {
    let mut index = end;
    while index > start && xml[index - 1].is_ascii_whitespace() {
        index -= 1;
    }
    index
}

fn close_insert(open: &[u8]) -> usize {
    open.iter()
        .rposition(|byte| *byte == b'>')
        .map_or(open.len(), |index| {
            if index > 0 && open[index - 1] == b'/' {
                index - 1
            } else {
                index
            }
        })
}

fn expand_root(open: &[u8], name: &[u8]) -> Result<Vec<u8>> {
    let insert = close_insert(open);
    let mut output = Vec::with_capacity(open.len() + name.len() + 3);
    output.extend_from_slice(&open[..insert]);
    output.extend_from_slice(b">");
    Ok(output)
}

fn close_name(name: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(name.len() + 3);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.extend_from_slice(b">");
    output
}

fn splice(source: &[u8], range: Range, replacement: &[u8]) -> Vec<u8> {
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

fn text(bytes: &[u8]) -> &str {
    std::str::from_utf8(bytes).expect("XML namespace and prefix names are UTF-8")
}
