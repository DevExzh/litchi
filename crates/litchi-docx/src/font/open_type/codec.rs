//! Bounded, namespace-aware, source-preserving codec for Word 2010 OpenType
//! run-property extensions.

use crate::error::{Error, Result};
use crate::paragraph::is_fragment_word_name;
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::fmt::Write as FmtWrite;

use super::model::{
    Ligatures, MC_NAMESPACE, NumForm, NumSpacing, OnOff, OpenType, StyleSet, StyleSetId,
    WORD_2010_NAMESPACE,
};
use super::validation::{
    MAX_STYLE_SETS, MAX_XML_BYTES, MAX_XML_DEPTH, MAX_XML_NODES, validate, validate_style_set,
};

const W14_PREFIX: &[u8] = b"w14";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ByteRange {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Ligatures,
    NumForm,
    NumSpacing,
    StylisticSets,
    CntxtAlts,
}

#[derive(Debug, Clone, Copy)]
struct KnownRange {
    range: ByteRange,
}

#[derive(Debug, Default)]
struct Layout {
    root_name: Vec<u8>,
    root_empty: Option<ByteRange>,
    root_close_start: Option<usize>,
    rpr_name: Vec<u8>,
    rpr_empty: Option<ByteRange>,
    rpr_close_start: Option<usize>,
    known: Vec<KnownRange>,
    extension_prefix: Option<Vec<u8>>,
}

/// Parse one complete `w:r` or `w:rPr` fragment.
pub(crate) fn parse(xml: &[u8]) -> Result<OpenType> {
    let (_, value) = locate(xml)?;
    Ok(value)
}

/// Rewrite only the modeled direct OpenType children of one `w:r`/`w:rPr`.
///
/// Every byte outside the modeled ranges is copied from the source.  In
/// particular, foreign children, unknown Word extensions, whitespace, and
/// lexical choices remain untouched.
pub(crate) fn rewrite(xml: &[u8], next: &OpenType) -> Result<Vec<u8>> {
    validate(next)?;
    let (layout, current) = locate(xml)?;
    if current == *next {
        return Ok(xml.to_vec());
    }

    let prefix = layout.extension_prefix.as_deref().unwrap_or(W14_PREFIX);
    let replacement = render(next, prefix)?;
    let insert_at = layout
        .known
        .iter()
        .map(|entry| entry.range.start)
        .min()
        .or(layout.rpr_close_start);

    if replacement.is_empty() && layout.known.is_empty() {
        return Ok(xml.to_vec());
    }

    if let Some(insert_at) = insert_at {
        return splice_known(xml, &layout.known, insert_at, &replacement);
    }

    let rpr_name = if layout.rpr_name.is_empty() {
        qualified_name(&layout.root_name, b"rPr")?
    } else {
        layout.rpr_name.clone()
    };
    if let Some(rpr_empty) = layout.rpr_empty {
        return expand_empty(xml, rpr_empty, &replacement, &rpr_name);
    }
    let wrapped = wrap_rpr(&rpr_name, &replacement);
    if let Some(root_close_start) = layout.root_close_start {
        return Ok(insert_at_bytes(xml, root_close_start, &wrapped));
    }
    if let Some(root_empty) = layout.root_empty {
        return expand_empty(xml, root_empty, &wrapped, &layout.root_name);
    }
    Err(Error::InvalidFormat(
        "Word run XML has no insertion point for OpenType properties".into(),
    ))
}

fn locate(xml: &[u8]) -> Result<(Layout, OpenType)> {
    if xml.len() > MAX_XML_BYTES {
        return Err(Error::InvalidFormat(format!(
            "Word run XML exceeds {MAX_XML_BYTES} bytes"
        )));
    }

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;

    let mut layout = Layout::default();
    let mut value = OpenType::default();
    let mut stack = Vec::<Frame>::new();
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut rpr_depth = None::<usize>;
    let mut depth = 0usize;
    let mut nodes = 0usize;

    loop {
        let event_start = offset(&reader)?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let event_end = offset(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        if depth == 0
            && fragment_prefix.is_none()
            && let Event::Start(element) | Event::Empty(element) = &event
        {
            fragment_prefix = Some(
                element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.into_inner().to_vec()),
            );
        }
        let is_word = match &event {
            Event::Start(element) | Event::Empty(element) => is_fragment_word_name(
                &namespace,
                element.name(),
                element.local_name().as_ref(),
                &fragment_prefix,
            ),
            _ => false,
        };
        let is_w14 = is_word_2010(&namespace);
        let event = event.into_owned();

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Word run XML element counter overflows usize".into())
            })?;
            if nodes > MAX_XML_NODES {
                return Err(Error::InvalidFormat(format!(
                    "Word OpenType XML exceeds {MAX_XML_NODES} elements"
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
                    if root_seen
                        || !is_word
                        || !matches!(element.local_name().as_ref(), b"r" | b"rPr")
                    {
                        return Err(Error::InvalidFormat(
                            "Word OpenType XML must have one w:r or w:rPr root".into(),
                        ));
                    }
                    root_seen = true;
                    layout.root_name = element.name().as_ref().to_vec();
                    if element.local_name().as_ref() == b"rPr" {
                        rpr_depth = Some(child_depth);
                        layout.rpr_name = layout.root_name.clone();
                    }
                } else if rpr_depth.is_none()
                    && child_depth == 2
                    && is_word
                    && element.local_name().as_ref() == b"rPr"
                {
                    rpr_depth = Some(child_depth);
                    layout.rpr_name = element.name().as_ref().to_vec();
                }

                let direct = rpr_depth == Some(depth);
                let kind = if direct {
                    direct_kind(&element, is_w14)
                } else {
                    None
                };
                if kind.is_some() {
                    if let Some(prefix) = element.name().prefix() {
                        layout.extension_prefix = Some(prefix.into_inner().to_vec());
                    }
                }
                if child_depth == 1 && is_empty_element_start(&element) {
                    layout.root_empty = Some(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                }
                if rpr_depth == Some(child_depth) && is_empty_element_start(&element) {
                    layout.rpr_empty = Some(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                }
                stack.push(Frame {
                    depth: child_depth,
                    start: event_start,
                    name: element.name().as_ref().to_vec(),
                    kind,
                });
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
                    if root_seen
                        || !is_word
                        || !matches!(element.local_name().as_ref(), b"r" | b"rPr")
                    {
                        return Err(Error::InvalidFormat(
                            "Word OpenType XML must have one w:r or w:rPr root".into(),
                        ));
                    }
                    root_seen = true;
                    root_closed = true;
                    layout.root_name = element.name().as_ref().to_vec();
                    layout.root_empty = Some(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                    if element.local_name().as_ref() == b"rPr" {
                        rpr_depth = Some(child_depth);
                        layout.rpr_name = layout.root_name.clone();
                        layout.rpr_empty = layout.root_empty;
                    }
                } else if rpr_depth.is_none()
                    && child_depth == 2
                    && is_word
                    && element.local_name().as_ref() == b"rPr"
                {
                    rpr_depth = Some(child_depth);
                    layout.rpr_name = element.name().as_ref().to_vec();
                    layout.rpr_empty = Some(ByteRange {
                        start: event_start,
                        end: event_end,
                    });
                }
                let direct = rpr_depth == Some(depth);
                if let Some(kind) = direct_kind(&element, is_w14).filter(|_| direct) {
                    if let Some(prefix) = element.name().prefix() {
                        layout.extension_prefix = Some(prefix.into_inner().to_vec());
                    }
                    layout.known.push(KnownRange {
                        range: ByteRange {
                            start: event_start,
                            end: event_end,
                        },
                    });
                    let element_prefix = element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec());
                    let parsed = parse_known(
                        &xml[event_start..event_end],
                        kind,
                        element_prefix.as_deref(),
                    )?;
                    assign(&mut value, kind, parsed)?;
                }
            },
            Event::End(element) => {
                let frame = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid Word OpenType XML nesting".into())
                })?;
                if frame.depth != depth || frame.name != element.name().as_ref() {
                    return Err(Error::InvalidFormat(
                        "mismatched Word OpenType XML end element".into(),
                    ));
                }
                if let Some(kind) = frame.kind {
                    if let Some(prefix) = name_prefix(&frame.name) {
                        layout.extension_prefix = Some(prefix.to_vec());
                    }
                    layout.known.push(KnownRange {
                        range: ByteRange {
                            start: frame.start,
                            end: event_end,
                        },
                    });
                    let parsed =
                        parse_known(&xml[frame.start..event_end], kind, name_prefix(&frame.name))?;
                    assign(&mut value, kind, parsed)?;
                }
                if rpr_depth == Some(frame.depth) && name_local(&frame.name) == b"rPr" {
                    layout.rpr_close_start = Some(event_start);
                    rpr_depth = None;
                }
                if depth == 1 {
                    root_closed = true;
                    layout.root_close_start = Some(event_start);
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word OpenType XML depth".into())
                })?;
            },
            Event::Text(text) if depth == 0 => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word OpenType XML has text outside its root".into(),
                    ));
                }
            },
            Event::CData(text) if depth == 0 => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(Error::InvalidFormat(
                        "Word OpenType XML has unexpected character data".into(),
                    ));
                }
            },
            Event::Decl(_) | Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "Word OpenType XML cannot contain declarations or processing instructions"
                        .into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if !root_seen || !root_closed || depth != 0 || !stack.is_empty() {
        return Err(Error::InvalidFormat(
            "Word OpenType XML is not one complete element".into(),
        ));
    }
    value.validate()?;
    Ok((layout, value))
}

#[derive(Debug)]
struct Frame {
    depth: usize,
    start: usize,
    name: Vec<u8>,
    kind: Option<Kind>,
}

fn direct_kind(element: &BytesStart<'_>, is_w14: bool) -> Option<Kind> {
    if !is_w14 {
        return None;
    }
    match element.local_name().as_ref() {
        b"ligatures" => Some(Kind::Ligatures),
        b"numForm" => Some(Kind::NumForm),
        b"numSpacing" => Some(Kind::NumSpacing),
        b"stylisticSets" => Some(Kind::StylisticSets),
        b"cntxtAlts" => Some(Kind::CntxtAlts),
        _ => None,
    }
}

fn assign(value: &mut OpenType, kind: Kind, parsed: Parsed) -> Result<()> {
    match (kind, parsed) {
        (Kind::Ligatures, Parsed::Ligatures(next)) => {
            if value.ligatures.replace(next).is_some() {
                return Err(super::model::invalid("duplicate w14:ligatures"));
            }
        },
        (Kind::NumForm, Parsed::NumForm(next)) => {
            if value.num_form.replace(next).is_some() {
                return Err(super::model::invalid("duplicate w14:numForm"));
            }
        },
        (Kind::NumSpacing, Parsed::NumSpacing(next)) => {
            if value.num_spacing.replace(next).is_some() {
                return Err(super::model::invalid("duplicate w14:numSpacing"));
            }
        },
        (Kind::StylisticSets, Parsed::StylisticSets(next)) => {
            if value.stylistic_sets_present() {
                return Err(super::model::invalid("duplicate w14:stylisticSets"));
            }
            value.set_stylistic_sets(Some(next))?;
        },
        (Kind::CntxtAlts, Parsed::CntxtAlts(next)) => {
            if value.cntxt_alts.replace(next).is_some() {
                return Err(super::model::invalid("duplicate w14:cntxtAlts"));
            }
        },
        _ => return Err(super::model::invalid("OpenType element/value mismatch")),
    }
    Ok(())
}

enum Parsed {
    Ligatures(Ligatures),
    NumForm(NumForm),
    NumSpacing(NumSpacing),
    StylisticSets(Vec<StyleSet>),
    CntxtAlts(OnOff),
}

fn parse_known(xml: &[u8], kind: Kind, prefix: Option<&[u8]>) -> Result<Parsed> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root = None::<Vec<u8>>;
    let mut result = None;
    let mut nodes = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        match event {
            Event::Start(element) => {
                depth += 1;
                nodes += 1;
                if depth > MAX_XML_DEPTH || nodes > 64 {
                    return Err(super::model::invalid("OpenType element exceeds XML bounds"));
                }
                if depth == 1 {
                    root = Some(element.name().as_ref().to_vec());
                    validate_element_name(&element, kind)?;
                    result = Some(parse_root_attributes(
                        &element,
                        kind,
                        prefix,
                        reader.decoder(),
                    )?);
                } else if kind == Kind::StylisticSets && depth == 2 {
                    if element.local_name().as_ref() != b"styleSet" {
                        return Err(super::model::invalid(
                            "w14:stylisticSets contains an unexpected child",
                        ));
                    }
                    let style = parse_style_set(&element, prefix, reader.decoder())?;
                    let parsed = result.as_mut().ok_or_else(|| {
                        super::model::invalid("stylisticSets attributes were not parsed")
                    })?;
                    if let Parsed::StylisticSets(values) = parsed {
                        if values.len() >= MAX_STYLE_SETS {
                            return Err(super::model::invalid("too many OpenType stylistic sets"));
                        }
                        values.push(style);
                    }
                } else {
                    return Err(super::model::invalid(
                        "OpenType element has unexpected children",
                    ));
                }
            },
            Event::Empty(element) => {
                depth += 1;
                nodes += 1;
                if depth > MAX_XML_DEPTH || nodes > 64 {
                    return Err(super::model::invalid("OpenType element exceeds XML bounds"));
                }
                if depth == 1 {
                    root = Some(element.name().as_ref().to_vec());
                    validate_element_name(&element, kind)?;
                    result = Some(parse_root_attributes(
                        &element,
                        kind,
                        prefix,
                        reader.decoder(),
                    )?);
                } else if kind == Kind::StylisticSets && depth == 2 {
                    if element.local_name().as_ref() != b"styleSet" {
                        return Err(super::model::invalid(
                            "w14:stylisticSets contains an unexpected child",
                        ));
                    }
                    let style = parse_style_set(&element, prefix, reader.decoder())?;
                    let parsed = result.as_mut().ok_or_else(|| {
                        super::model::invalid("stylisticSets attributes were not parsed")
                    })?;
                    if let Parsed::StylisticSets(values) = parsed {
                        if values.len() >= MAX_STYLE_SETS {
                            return Err(super::model::invalid("too many OpenType stylistic sets"));
                        }
                        values.push(style);
                    }
                } else {
                    return Err(super::model::invalid(
                        "OpenType element has unexpected children",
                    ));
                }
                depth -= 1;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(super::model::invalid("unexpected OpenType XML end"));
                }
                if depth == 1 && root.as_deref() != Some(element.name().as_ref()) {
                    return Err(super::model::invalid("mismatched OpenType XML end"));
                }
                depth -= 1;
            },
            Event::Text(text) => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(super::model::invalid(
                        "OpenType elements cannot contain text",
                    ));
                }
            },
            Event::CData(text) => {
                if !text.into_inner().iter().all(u8::is_ascii_whitespace) {
                    return Err(super::model::invalid(
                        "OpenType elements cannot contain text",
                    ));
                }
            },
            Event::Comment(_) | Event::Decl(_) => {},
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(super::model::invalid("invalid OpenType XML event"));
            },
            Event::Eof => break,
        }
    }
    if depth != 0 {
        return Err(super::model::invalid("unterminated OpenType element"));
    }
    result.ok_or_else(|| super::model::invalid("OpenType element is empty"))
}

fn validate_element_name(element: &BytesStart<'_>, kind: Kind) -> Result<()> {
    let expected = match kind {
        Kind::Ligatures => b"ligatures".as_slice(),
        Kind::NumForm => b"numForm".as_slice(),
        Kind::NumSpacing => b"numSpacing".as_slice(),
        Kind::StylisticSets => b"stylisticSets".as_slice(),
        Kind::CntxtAlts => b"cntxtAlts".as_slice(),
    };
    if element.local_name().as_ref() != expected {
        return Err(super::model::invalid("unexpected OpenType element name"));
    }
    Ok(())
}

fn parse_root_attributes(
    element: &BytesStart<'_>,
    kind: Kind,
    prefix: Option<&[u8]>,
    decoder: Decoder,
) -> Result<Parsed> {
    let mut value = None::<String>;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        if name == b"mc:Ignorable" {
            continue;
        }
        let local = attribute.key.local_name();
        if local.as_ref() != b"val" {
            return Err(super::model::invalid(format!(
                "unexpected OpenType attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        if !attribute_namespace_matches(name, prefix) {
            return Err(super::model::invalid(
                "OpenType val is not namespace-qualified",
            ));
        }
        if value.is_some() {
            return Err(super::model::invalid("duplicate OpenType val attribute"));
        }
        value = Some(
            attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| Error::Xml(error.to_string()))?
                .into_owned(),
        );
    }

    match kind {
        Kind::Ligatures => Ok(Parsed::Ligatures(Ligatures::parse(
            value
                .as_deref()
                .ok_or_else(|| super::model::invalid("w14:ligatures requires val"))?,
        )?)),
        Kind::NumForm => Ok(Parsed::NumForm(NumForm::parse(
            value
                .as_deref()
                .ok_or_else(|| super::model::invalid("w14:numForm requires val"))?,
        )?)),
        Kind::NumSpacing => Ok(Parsed::NumSpacing(NumSpacing::parse(
            value
                .as_deref()
                .ok_or_else(|| super::model::invalid("w14:numSpacing requires val"))?,
        )?)),
        Kind::StylisticSets => {
            if value.is_some() {
                return Err(super::model::invalid(
                    "w14:stylisticSets does not have a val attribute",
                ));
            }
            Ok(Parsed::StylisticSets(Vec::new()))
        },
        Kind::CntxtAlts => Ok(Parsed::CntxtAlts(OnOff::new(
            value.as_deref().map(parse_on_off).transpose()?,
        ))),
    }
}

fn parse_style_set(
    element: &BytesStart<'_>,
    prefix: Option<&[u8]>,
    decoder: Decoder,
) -> Result<StyleSet> {
    let mut id = None::<u8>;
    let mut enabled = None::<bool>;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if name == b"xmlns" || name.starts_with(b"xmlns:") {
            continue;
        }
        let local = attribute.key.local_name();
        if !attribute_namespace_matches(name, prefix) {
            return Err(super::model::invalid(
                "OpenType styleSet attribute is not qualified",
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        match local.as_ref() {
            b"id" => {
                if id.is_some() {
                    return Err(super::model::invalid("duplicate styleSet id attribute"));
                }
                id = Some(value.parse::<u8>().map_err(|_| {
                    super::model::invalid("styleSet id is not an unsigned decimal number")
                })?);
            },
            b"val" => {
                if enabled.is_some() {
                    return Err(super::model::invalid("duplicate styleSet val attribute"));
                }
                enabled = Some(parse_on_off(&value)?);
            },
            _ => return Err(super::model::invalid("unexpected styleSet attribute")),
        }
    }
    let id =
        StyleSetId::try_from(id.ok_or_else(|| super::model::invalid("styleSet requires id"))?)?;
    let value = StyleSet { id, enabled };
    validate_style_set(&value)?;
    Ok(value)
}

fn parse_on_off(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(super::model::invalid(format!(
            "invalid ST_OnOff value '{value}'"
        ))),
    }
}

fn attribute_namespace_matches(name: &[u8], prefix: Option<&[u8]>) -> bool {
    let Some(separator) = name.iter().position(|byte| *byte == b':') else {
        return false;
    };
    let attribute_prefix = &name[..separator];
    prefix.is_some_and(|expected| expected == attribute_prefix)
}

fn is_word_2010(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == WORD_2010_NAMESPACE)
        || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == W14_PREFIX)
}

fn name_prefix(name: &[u8]) -> Option<&[u8]> {
    name.iter()
        .position(|byte| *byte == b':')
        .and_then(|index| name.get(..index))
}

fn name_local(name: &[u8]) -> &[u8] {
    name_prefix(name)
        .and_then(|prefix| name.get(prefix.len() + 1..))
        .unwrap_or(name)
}

fn offset(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::InvalidFormat("Word OpenType XML offset overflow".into()))
}

fn is_empty_element_start(element: &BytesStart<'_>) -> bool {
    element.as_ref().ends_with(b"/>")
}

fn render(value: &OpenType, prefix: &[u8]) -> Result<Vec<u8>> {
    let prefix = std::str::from_utf8(prefix)
        .map_err(|_| Error::InvalidFormat("OpenType namespace prefix is not UTF-8".into()))?;
    let mut output = String::new();
    let mut first = true;
    if let Some(value) = value.ligatures {
        write_leaf(
            &mut output,
            prefix,
            "ligatures",
            Some(("val", value.as_str())),
            &mut first,
        )?;
    }
    if let Some(value) = value.num_form {
        write_leaf(
            &mut output,
            prefix,
            "numForm",
            Some(("val", value.as_str())),
            &mut first,
        )?;
    }
    if let Some(value) = value.num_spacing {
        write_leaf(
            &mut output,
            prefix,
            "numSpacing",
            Some(("val", value.as_str())),
            &mut first,
        )?;
    }
    if value.stylistic_sets_present() {
        write!(output, "<{prefix}:stylisticSets")?;
        if first {
            write_namespace_declarations(&mut output, prefix);
            first = false;
        }
        output.push('>');
        for style in &value.stylistic_sets {
            write!(
                output,
                "<{prefix}:styleSet {prefix}:id=\"{}\"",
                style.id.get()
            )?;
            if let Some(enabled) = style.enabled {
                write!(
                    output,
                    " {prefix}:val=\"{}\"",
                    if enabled { "1" } else { "0" }
                )?;
            }
            output.push_str("/>");
        }
        write!(output, "</{prefix}:stylisticSets>")?;
    }
    if let Some(value) = value.cntxt_alts {
        write_leaf(
            &mut output,
            prefix,
            "cntxtAlts",
            value
                .authored()
                .map(|value| ("val", if value { "1" } else { "0" })),
            &mut first,
        )?;
    }
    Ok(output.into_bytes())
}

fn write_leaf(
    output: &mut String,
    prefix: &str,
    local: &str,
    attribute: Option<(&str, &str)>,
    first: &mut bool,
) -> Result<()> {
    write!(output, "<{prefix}:{local}")?;
    if *first {
        write_namespace_declarations(output, prefix);
        *first = false;
    }
    if let Some((name, value)) = attribute {
        write!(output, " {prefix}:{name}=\"{value}\"")?;
    }
    output.push_str("/>");
    Ok(())
}

fn write_namespace_declarations(output: &mut String, prefix: &str) {
    output.push_str(" xmlns:");
    output.push_str(prefix);
    output.push_str("=\"");
    output.push_str(&String::from_utf8_lossy(WORD_2010_NAMESPACE));
    output.push_str("\" xmlns:mc=\"");
    output.push_str(&String::from_utf8_lossy(MC_NAMESPACE));
    output.push_str("\" mc:Ignorable=\"");
    output.push_str(prefix);
    output.push('"');
}

fn qualified_name(root: &[u8], local: &[u8]) -> Result<Vec<u8>> {
    let separator = root
        .iter()
        .position(|byte| *byte == b':')
        .ok_or_else(|| Error::InvalidFormat("Word run root has no usable prefix".into()))?;
    let prefix = root
        .get(..separator)
        .filter(|prefix| !prefix.is_empty())
        .ok_or_else(|| Error::InvalidFormat("Word run root has no usable prefix".into()))?;
    let mut name = prefix.to_vec();
    name.push(b':');
    name.extend_from_slice(local);
    Ok(name)
}

fn wrap_rpr(name: &[u8], body: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(name.len() * 2 + body.len() + 5);
    output.extend_from_slice(b"<");
    output.extend_from_slice(name);
    output.push(b'>');
    output.extend_from_slice(body);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    output
}

fn splice_known(
    source: &[u8],
    known: &[KnownRange],
    insertion: usize,
    replacement: &[u8],
) -> Result<Vec<u8>> {
    let mut ranges = known.to_vec();
    ranges.sort_by_key(|entry| entry.range.start);
    let mut output = Vec::with_capacity(source.len() + replacement.len());
    let mut cursor = 0usize;
    let mut inserted = false;
    for entry in ranges {
        if entry.range.start < cursor || entry.range.end > source.len() {
            return Err(Error::InvalidFormat(
                "overlapping OpenType XML ranges".into(),
            ));
        }
        if !inserted && insertion <= entry.range.start {
            output.extend_from_slice(&source[cursor..insertion]);
            output.extend_from_slice(replacement);
            inserted = true;
            cursor = insertion;
        }
        output.extend_from_slice(&source[cursor..entry.range.start]);
        cursor = entry.range.end;
    }
    if !inserted {
        if insertion < cursor || insertion > source.len() {
            return Err(Error::InvalidFormat(
                "invalid OpenType insertion offset".into(),
            ));
        }
        output.extend_from_slice(&source[cursor..insertion]);
        output.extend_from_slice(replacement);
        cursor = insertion;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

fn insert_at_bytes(source: &[u8], offset: usize, insertion: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(source.len() + insertion.len());
    output.extend_from_slice(&source[..offset]);
    output.extend_from_slice(insertion);
    output.extend_from_slice(&source[offset..]);
    output
}

fn expand_empty(source: &[u8], range: ByteRange, body: &[u8], name: &[u8]) -> Result<Vec<u8>> {
    let raw = source
        .get(range.start..range.end)
        .ok_or_else(|| Error::InvalidFormat("invalid empty Word root range".into()))?;
    let close = raw
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| Error::InvalidFormat("empty Word root has no close".into()))?;
    let slash = raw[..close]
        .iter()
        .rposition(|byte| !byte.is_ascii_whitespace())
        .filter(|index| raw[*index] == b'/')
        .ok_or_else(|| Error::InvalidFormat("empty Word root is missing '/>'".into()))?;
    let mut replacement = Vec::with_capacity(raw.len() + body.len() + name.len() + 3);
    replacement.extend_from_slice(&raw[..slash]);
    replacement.push(b'>');
    replacement.extend_from_slice(body);
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(name);
    replacement.push(b'>');
    Ok(splice_bytes(source, range, &replacement))
}

fn splice_bytes(source: &[u8], range: ByteRange, replacement: &[u8]) -> Vec<u8> {
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
