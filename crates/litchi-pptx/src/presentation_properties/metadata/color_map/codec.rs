//! `PresentationML` color-map parsing and resolution.

use std::ops::Range;

use super::super::{Limits, checked_output_len, process_ooxml};
use super::model::{Map, Override, Role, Slot, Value};
use crate::presentation_properties::metadata::is_presentationml_name;
use crate::{Error, Result};
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::{
    attributes::BytesStartExt as _, is_drawingml_name, unqualified_attribute_value,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::NsReader;

/// The XML owner context needed to resolve a typed color-map value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Source {
    Master,
    Override {
        root_name: Vec<u8>,
        root_label: String,
    },
}

/// A parsed color-map value together with source spans for bounded edits.
#[derive(Debug, Clone)]
pub(crate) struct Located {
    pub(crate) source: Source,
    pub(crate) value: Value,
    pub(crate) map_attributes: Option<[Range<usize>; 12]>,
}

pub(crate) fn locate_master(xml: &[u8], limits: Limits) -> Result<Located> {
    locate_source(xml, &Source::Master, limits)
}

pub(crate) fn locate_override(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
    limits: Limits,
) -> Result<Located> {
    locate_source(
        xml,
        &Source::Override {
            root_name: root_name.to_vec(),
            root_label: root_label.to_owned(),
        },
        limits,
    )
}

pub(crate) fn locate_source(xml: &[u8], source: &Source, limits: Limits) -> Result<Located> {
    check_input(xml, limits)?;
    let value = match source {
        Source::Master => Value::Master(parse_master(xml, limits)?),
        Source::Override {
            root_name,
            root_label,
        } => Value::Override(parse_override(xml, root_name, root_label, limits)?),
    };
    locate(xml, source.clone(), value, limits)
}

fn locate(xml: &[u8], source: Source, value: Value, limits: Limits) -> Result<Located> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut state = ScanState::default();
    let mut nodes = 0usize;

    loop {
        let start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                if state.depth >= limits.max_depth() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: state.depth + 1,
                        maximum: limits.max_depth(),
                    });
                }
                scan_start(&mut state, &namespace, &element, &source, start, limits)?;
            },
            Event::Empty(element) => {
                check_empty_depth(state.depth, limits)?;
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                scan_empty(&mut state, &namespace, &element, &source, start, limits)?;
            },
            Event::End(element) => scan_end(&mut state, &namespace, &element)?,
            Event::Eof => break,
            _ => {},
        }
    }

    state.finish(source, value)
}

#[derive(Default)]
struct ScanState {
    depth: usize,
    saw_root: bool,
    override_depth: Option<usize>,
    map_attributes: Option<[Range<usize>; 12]>,
}

impl ScanState {
    fn finish(self, source: Source, value: Value) -> Result<Located> {
        if self.depth != 0 || !self.saw_root || self.override_depth.is_some() {
            return Err(Error::Invalid("unterminated color-map XML".to_string()));
        }
        if matches!(&source, Source::Master) && self.map_attributes.is_none() {
            return Err(Error::Invalid(
                "slide master is missing its color map".to_string(),
            ));
        }
        Ok(Located {
            source,
            value,
            map_attributes: self.map_attributes,
        })
    }
}

fn scan_start(
    state: &mut ScanState,
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    source: &Source,
    start: usize,
    limits: Limits,
) -> Result<()> {
    state.depth = state.depth.checked_add(1).ok_or_else(|| Error::Limit {
        kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
        observed: usize::MAX,
        maximum: limits.max_depth(),
    })?;
    if state.depth == 1 {
        if state.saw_root {
            return Err(Error::Invalid(
                "color-map XML has multiple roots".to_string(),
            ));
        }
        require_source_root(namespace, element, source)?;
        state.saw_root = true;
    } else if is_master_map(source, state.depth, namespace, element) {
        store_map_attributes(
            &mut state.map_attributes,
            map_attribute_spans(element, start)?,
        )?;
    } else if is_override_container(source, state.depth, namespace, element) {
        if state.override_depth.is_some() {
            return Err(Error::Invalid(
                "color-map XML has multiple color-map overrides".to_string(),
            ));
        }
        state.override_depth = Some(state.depth);
    } else if state.override_depth == Some(state.depth - 1)
        && is_override_mapping(namespace, element)
    {
        store_map_attributes(
            &mut state.map_attributes,
            map_attribute_spans(element, start)?,
        )?;
    }
    Ok(())
}

fn scan_empty(
    state: &mut ScanState,
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    source: &Source,
    start: usize,
    limits: Limits,
) -> Result<()> {
    check_empty_depth(state.depth, limits)?;
    if state.depth == 0 {
        if state.saw_root {
            return Err(Error::Invalid(
                "color-map XML has multiple roots".to_string(),
            ));
        }
        require_source_root(namespace, element, source)?;
        state.saw_root = true;
    } else if is_master_map(
        source,
        state.depth.checked_add(1).ok_or_else(|| Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
            observed: usize::MAX,
            maximum: limits.max_depth(),
        })?,
        namespace,
        element,
    ) || (state.override_depth == Some(state.depth)
        && is_override_mapping(namespace, element))
    {
        store_map_attributes(
            &mut state.map_attributes,
            map_attribute_spans(element, start)?,
        )?;
    }
    Ok(())
}

fn scan_end(
    state: &mut ScanState,
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &quick_xml::events::BytesEnd<'_>,
) -> Result<()> {
    if state.override_depth == Some(state.depth)
        && is_presentationml_name(namespace, element.name(), b"clrMapOvr")
    {
        state.override_depth = None;
    }
    state.depth = state
        .depth
        .checked_sub(1)
        .ok_or_else(|| Error::Invalid("invalid color-map XML nesting".to_string()))?;
    Ok(())
}

pub(crate) fn rewrite(
    source: &[u8],
    located: &Located,
    desired: Value,
    limits: Limits,
) -> Result<Vec<u8>> {
    check_input(source, limits)?;
    if located.value == desired {
        if source.len() > limits.max_output_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
                observed: source.len(),
                maximum: limits.max_output_bytes(),
            });
        }
        return Ok(source.to_vec());
    }
    let before = mapped_value(&located.value).ok_or_else(|| {
        Error::Invalid("cannot edit a color-map without an explicit mapping".to_string())
    })?;
    let after = mapped_value(&desired).ok_or_else(|| {
        Error::Invalid("cannot create or remove a color-map in a bounded edit".to_string())
    })?;
    let attributes = located.map_attributes.as_ref().ok_or_else(|| {
        Error::Invalid("color-map source has no editable mapping attributes".to_string())
    })?;

    let mut removed = 0usize;
    let mut added = 0usize;
    for (index, slot) in Slot::ALL.into_iter().enumerate() {
        let old = before.color(slot);
        let new = after.color(slot);
        if old != new {
            let range = &attributes[index];
            removed =
                removed
                    .checked_add(range.end - range.start)
                    .ok_or(Error::CapacityOverflow {
                        resource: "serialized color-map XML",
                    })?;
            added = added
                .checked_add(new.as_str().len())
                .ok_or(Error::CapacityOverflow {
                    resource: "serialized color-map XML",
                })?;
        }
    }
    checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "serialized color-map XML",
    )?;

    let mut replacements = Vec::new();
    for (index, slot) in Slot::ALL.into_iter().enumerate() {
        let old = before.color(slot);
        let new = after.color(slot);
        if old != new {
            replacements.push(Replacement {
                range: attributes[index].clone(),
                value: new.as_str().as_bytes().to_vec(),
            });
        }
    }
    apply_replacements(source, replacements, limits)
}

fn check_input(xml: &[u8], limits: Limits) -> Result<()> {
    if xml.len() > limits.max_input_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
            observed: xml.len(),
            maximum: limits.max_input_bytes(),
        });
    }
    Ok(())
}

fn check_empty_depth(depth: usize, limits: Limits) -> Result<()> {
    let observed = depth.checked_add(1).unwrap_or(usize::MAX);
    if observed > limits.max_depth() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
            observed,
            maximum: limits.max_depth(),
        });
    }
    Ok(())
}

fn mapped_value(value: &Value) -> Option<Map> {
    match value {
        Value::Master(map) => Some(*map),
        Value::Override(Some(Override::Override(map))) => Some(*map),
        Value::Override(None | Some(Override::Master)) => None,
    }
}

fn require_source_root(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    source: &Source,
) -> Result<()> {
    let (expected_name, label) = match source {
        Source::Master => (b"sldMaster".as_slice(), "slide master"),
        Source::Override {
            root_name,
            root_label,
        } => (root_name.as_slice(), root_label.as_str()),
    };
    require_root(namespace, element, expected_name, label)
}

fn is_master_map(
    source: &Source,
    depth: usize,
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> bool {
    matches!(source, Source::Master)
        && depth == 2
        && is_presentationml_name(namespace, element.name(), b"clrMap")
}

fn is_override_container(
    source: &Source,
    depth: usize,
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> bool {
    matches!(source, Source::Override { .. })
        && depth == 2
        && is_presentationml_name(namespace, element.name(), b"clrMapOvr")
}

fn is_override_mapping(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> bool {
    is_drawingml_name(namespace, element.name(), b"overrideClrMapping")
}

fn store_map_attributes(
    target: &mut Option<[Range<usize>; 12]>,
    value: [Range<usize>; 12],
) -> Result<()> {
    if target.replace(value).is_some() {
        return Err(Error::Invalid(
            "color-map XML has multiple editable mappings".to_string(),
        ));
    }
    Ok(())
}

fn map_attribute_spans(element: &BytesStart<'_>, start: usize) -> Result<[Range<usize>; 12]> {
    let span = |slot: Slot| -> Result<Range<usize>> {
        let span = attribute_span(element.as_ref(), slot.as_str().as_bytes())?;
        let value_start = start
            .checked_add(1)
            .and_then(|value| value.checked_add(span.value_start))
            .ok_or_else(|| Error::Invalid("color-map attribute offset overflow".to_string()))?;
        let value_end = start
            .checked_add(1)
            .and_then(|value| value.checked_add(span.value_end))
            .ok_or_else(|| Error::Invalid("color-map attribute offset overflow".to_string()))?;
        Ok(value_start..value_end)
    };
    Ok([
        span(Slot::Background1)?,
        span(Slot::Text1)?,
        span(Slot::Background2)?,
        span(Slot::Text2)?,
        span(Slot::Accent1)?,
        span(Slot::Accent2)?,
        span(Slot::Accent3)?,
        span(Slot::Accent4)?,
        span(Slot::Accent5)?,
        span(Slot::Accent6)?,
        span(Slot::Hyperlink)?,
        span(Slot::FollowedHyperlink)?,
    ])
}

struct AttributeSpan {
    value_start: usize,
    value_end: usize,
}

fn attribute_span(raw: &[u8], key: &[u8]) -> Result<AttributeSpan> {
    let mut index = 0usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && !matches!(raw[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] == b'>' || raw[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'>' | b'/')
        {
            index += 1;
        }
        let name = &raw[name_start..index];
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len() || raw[index] != b'=' {
            return Err(Error::Invalid(
                "color-map attribute has no value".to_string(),
            ));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw.get(index).ok_or_else(|| {
            Error::Invalid("color-map attribute value is unterminated".to_string())
        })?;
        if quote != b'"' && quote != b'\'' {
            return Err(Error::Invalid(
                "color-map attribute value is not quoted".to_string(),
            ));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        if index >= raw.len() {
            return Err(Error::Invalid(
                "color-map attribute value is unterminated".to_string(),
            ));
        }
        if name == key {
            return Ok(AttributeSpan {
                value_start,
                value_end: index,
            });
        }
        index += 1;
    }
    Err(Error::Invalid(format!(
        "color-map attribute '{}' has no source span",
        String::from_utf8_lossy(key)
    )))
}

struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

fn apply_replacements(
    source: &[u8],
    mut replacements: Vec<Replacement>,
    limits: Limits,
) -> Result<Vec<u8>> {
    replacements.sort_by_key(|replacement| std::cmp::Reverse(replacement.range.start));
    let mut upper = source.len();
    let mut removed = 0usize;
    let mut added = 0usize;
    for replacement in &replacements {
        if replacement.range.start > replacement.range.end
            || replacement.range.end > source.len()
            || replacement.range.end > upper
        {
            return Err(Error::Invalid(
                "color-map patch ranges overlap or escape the source".to_string(),
            ));
        }
        removed = removed
            .checked_add(replacement.range.end - replacement.range.start)
            .ok_or(Error::CapacityOverflow {
                resource: "serialized color-map XML",
            })?;
        added = added
            .checked_add(replacement.value.len())
            .ok_or(Error::CapacityOverflow {
                resource: "serialized color-map XML",
            })?;
        upper = replacement.range.start;
    }
    let output_len = checked_output_len(
        source.len(),
        removed,
        added,
        limits,
        "serialized color-map XML",
    )?;
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "serialized color-map XML",
            source,
        })?;
    let mut cursor = 0usize;
    for replacement in replacements.into_iter().rev() {
        output.extend_from_slice(&source[cursor..replacement.range.start]);
        output.extend_from_slice(&replacement.value);
        cursor = replacement.range.end;
    }
    output.extend_from_slice(&source[cursor..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("color-map XML offset does not fit usize".to_string()))
}

impl Role {
    fn from_str(value: &str) -> Option<Self> {
        match value {
            "dk1" => Some(Self::Dark1),
            "lt1" => Some(Self::Light1),
            "dk2" => Some(Self::Dark2),
            "lt2" => Some(Self::Light2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hlink" => Some(Self::Hyperlink),
            "folHlink" => Some(Self::FollowedHyperlink),
            _ => None,
        }
    }
}

impl Map {
    fn from_element(
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        label: &str,
    ) -> Result<Self> {
        Ok(Self {
            background1: required_role(element, b"bg1", decoder, label)?,
            text1: required_role(element, b"tx1", decoder, label)?,
            background2: required_role(element, b"bg2", decoder, label)?,
            text2: required_role(element, b"tx2", decoder, label)?,
            accent1: required_role(element, b"accent1", decoder, label)?,
            accent2: required_role(element, b"accent2", decoder, label)?,
            accent3: required_role(element, b"accent3", decoder, label)?,
            accent4: required_role(element, b"accent4", decoder, label)?,
            accent5: required_role(element, b"accent5", decoder, label)?,
            accent6: required_role(element, b"accent6", decoder, label)?,
            hyperlink: required_role(element, b"hlink", decoder, label)?,
            followed_hyperlink: required_role(element, b"folHlink", decoder, label)?,
        })
    }
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_master(xml: &[u8], limits: Limits) -> Result<Map> {
    check_input(xml, limits)?;
    let xml = process_ooxml(xml, limits)?;
    if xml.len() > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: xml.len(),
            maximum: limits.max_output_bytes(),
        });
    }
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut color_map = None;
    let mut nodes = 0usize;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                depth = depth.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                    observed: usize::MAX,
                    maximum: limits.max_depth(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                if depth > limits.max_depth() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth,
                        maximum: limits.max_depth(),
                    });
                }
                if depth == 1 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, b"sldMaster", "slide master")?;
                    saw_root = true;
                } else if depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"clrMap")
                {
                    store_color_map(
                        &mut color_map,
                        Map::from_element(&element, decoder, "slide-master color map")?,
                    )?;
                }
            },
            Event::Empty(element) => {
                check_empty_depth(depth, limits)?;
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                if depth == 0 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, b"sldMaster", "slide master")?;
                    saw_root = true;
                } else if depth == 1
                    && is_presentationml_name(&namespace, element.name(), b"clrMap")
                {
                    store_color_map(
                        &mut color_map,
                        Map::from_element(&element, decoder, "slide-master color map")?,
                    )?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid color-map XML nesting".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !saw_root {
        return Err(Error::Invalid(
            "unterminated slide-master color-map XML".to_string(),
        ));
    }
    color_map.ok_or_else(|| Error::Invalid("slide master is missing its color map".to_string()))
}

/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn parse_override(
    xml: &[u8],
    root_name: &[u8],
    root_label: &str,
    limits: Limits,
) -> Result<Option<Override>> {
    check_input(xml, limits)?;
    let xml = process_ooxml(xml, limits)?;
    if xml.len() > limits.max_output_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::OutputBytes),
            observed: xml.len(),
            maximum: limits.max_output_bytes(),
        });
    }
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut saw_override = false;
    let mut override_depth = None;
    let mut mapping = None;
    let mut nodes = 0usize;

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                depth = depth.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                    observed: usize::MAX,
                    maximum: limits.max_depth(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                if depth > limits.max_depth() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlDepth,
                        ),
                        observed: depth,
                        maximum: limits.max_depth(),
                    });
                }
                if depth == 1 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 2
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    if saw_override {
                        return Err(Error::Invalid(format!(
                            "{root_label} has multiple color-map overrides"
                        )));
                    }
                    saw_override = true;
                    override_depth = Some(depth);
                } else if override_depth == Some(depth - 1) {
                    parse_override_mapping(&namespace, &element, decoder, &mut mapping)?;
                }
            },
            Event::Empty(element) => {
                check_empty_depth(depth, limits)?;
                check_attributes(&element, decoder, limits)?;
                nodes = nodes.checked_add(1).ok_or_else(|| Error::Limit {
                    kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                    observed: usize::MAX,
                    maximum: limits.max_nodes(),
                })?;
                if nodes > limits.max_nodes() {
                    return Err(Error::Limit {
                        kind: crate::opened::LimitKind::Codec(
                            crate::opened::CodecLimitKind::XmlNodes,
                        ),
                        observed: nodes,
                        maximum: limits.max_nodes(),
                    });
                }
                if depth == 0 {
                    if saw_root {
                        return Err(Error::Invalid(
                            "color-map XML has multiple roots".to_string(),
                        ));
                    }
                    require_root(&namespace, &element, root_name, root_label)?;
                    saw_root = true;
                } else if depth == 1
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    if saw_override {
                        return Err(Error::Invalid(format!(
                            "{root_label} has multiple color-map overrides"
                        )));
                    }
                    saw_override = true;
                } else if override_depth == Some(depth) {
                    parse_override_mapping(&namespace, &element, decoder, &mut mapping)?;
                }
            },
            Event::End(element) => {
                if override_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"clrMapOvr")
                {
                    override_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| Error::Invalid("invalid color-map XML nesting".to_string()))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || !saw_root {
        return Err(Error::Invalid(format!(
            "unterminated {root_label} color-map XML"
        )));
    }
    if saw_override && mapping.is_none() {
        return Err(Error::Invalid(format!(
            "{root_label} color-map override has no mapping"
        )));
    }
    Ok(mapping)
}

fn require_root(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_name: &[u8],
    label: &str,
) -> Result<()> {
    if is_presentationml_name(namespace, element.name(), expected_name) {
        Ok(())
    } else {
        Err(Error::Invalid(format!(
            "color-map XML must have a PresentationML {label} root"
        )))
    }
}

fn check_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    limits: Limits,
) -> Result<()> {
    let mut count = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count.checked_add(1).ok_or_else(|| Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
            observed: usize::MAX,
            maximum: limits.max_attributes(),
        })?;
        if count > limits.max_attributes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlAttributes),
                observed: count,
                maximum: limits.max_attributes(),
            });
        }
        if attribute.key.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.key.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        if attribute.value.as_ref().len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(
                    crate::opened::CodecLimitKind::XmlAttributeBytes,
                ),
                observed: attribute.value.as_ref().len(),
                maximum: limits.max_string_bytes(),
            });
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        if value.len() > limits.max_string_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::StringBytes),
                observed: value.len(),
                maximum: limits.max_string_bytes(),
            });
        }
    }
    Ok(())
}

fn parse_override_mapping(
    namespace: &quick_xml::name::ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    mapping: &mut Option<Override>,
) -> Result<()> {
    let value = if is_drawingml_name(namespace, element.name(), b"masterClrMapping") {
        Override::Master
    } else if is_drawingml_name(namespace, element.name(), b"overrideClrMapping") {
        Override::Override(Map::from_element(element, decoder, "color-map override")?)
    } else {
        return Ok(());
    };

    if mapping.replace(value).is_some() {
        return Err(Error::Invalid(
            "color-map override has multiple mappings".to_string(),
        ));
    }
    Ok(())
}

fn store_color_map(slot: &mut Option<Map>, value: Map) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::Invalid(
            "slide master has multiple color maps".to_string(),
        ));
    }
    Ok(())
}

fn required_role(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
    label: &str,
) -> Result<Role> {
    let value = unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        Error::Invalid(format!(
            "{label} is missing its {} attribute",
            String::from_utf8_lossy(name)
        ))
    })?;
    Role::from_str(&value).ok_or_else(|| {
        Error::Invalid(format!(
            "{label} has unsupported {} value '{value}'",
            String::from_utf8_lossy(name)
        ))
    })
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::shape::theme::{Color as ThemeColor, Palette, Slot as ThemeSlot};

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    const MAP: &str = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"
        accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4"
        accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;

    #[test]
    fn parses_master_map_and_layout_override_by_namespace() {
        let master = parse_master(
            format!(r#"<p:sldMaster xmlns:p="{P}"><p:clrMap {MAP}/></p:sldMaster>"#).as_bytes(),
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(master.color(Slot::Background1), Role::Light1);
        assert_eq!(master.color(Slot::Text1), Role::Dark1);

        let layout = parse_override(
            format!(
                r#"<q:sldLayout xmlns:q="{P}" xmlns:d="{A}"><q:clrMapOvr>
                <d:overrideClrMapping {MAP}/></q:clrMapOvr></q:sldLayout>"#
            )
            .as_bytes(),
            b"sldLayout",
            "slide layout",
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(layout, Some(Override::Override(master)));

        let palette =
            Palette::new("Office").with(ThemeSlot::Light1, ThemeColor::Rgb("FFFFFF".to_owned()));
        assert!(matches!(
            master.resolve(&palette, Slot::Background1),
            Some(ThemeColor::Rgb(value)) if value == "FFFFFF"
        ));
    }

    #[test]
    fn rejects_incomplete_or_duplicate_color_maps() {
        let incomplete =
            format!(r#"<p:sldMaster xmlns:p="{P}"><p:clrMap bg1="lt1"/></p:sldMaster>"#);
        assert!(parse_master(incomplete.as_bytes(), Limits::DEFAULT).is_err());

        let duplicate = format!(
            r#"<p:sld xmlns:p="{P}" xmlns:a="{A}"><p:clrMapOvr>
                <a:masterClrMapping/><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
        );
        assert!(parse_override(duplicate.as_bytes(), b"sld", "slide", Limits::DEFAULT).is_err());

        let multiple_roots = format!(
            r#"<p:sldMaster xmlns:p="{P}"><p:clrMap {MAP}/></p:sldMaster>
            <p:sldMaster><p:clrMap {MAP}/></p:sldMaster>"#
        );
        assert!(parse_master(multiple_roots.as_bytes(), Limits::DEFAULT).is_err());
    }

    #[test]
    fn empty_color_map_depth_is_checked_before_attributes() {
        let limits = Limits::DEFAULT.with_depth(2).expect("depth limit");
        let exact = format!(r#"<p:sldMaster xmlns:p="{P}"><p:clrMap {MAP}/></p:sldMaster>"#);
        parse_master(exact.as_bytes(), limits).expect("depth-two empty element");

        let one_over = format!(
            r#"<p:sldMaster xmlns:p="{P}"><p:holder><p:clrMap {MAP}/></p:holder></p:sldMaster>"#
        );
        assert_xml_depth_limit(
            parse_master(one_over.as_bytes(), limits).expect_err("depth-three empty element"),
            3,
            2,
        );
        assert!(matches!(
            parse_master(b"<p:sldMaster xmlns:p=\"bad\">", Limits::DEFAULT),
            Err(Error::Invalid(_))
        ));
    }

    fn assert_xml_depth_limit(error: Error, observed: usize, maximum: usize) {
        match error {
            Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlDepth),
                observed: actual,
                maximum: configured,
            } => {
                assert_eq!(actual, observed);
                assert_eq!(configured, maximum);
            },
            other => panic!("expected XML depth limit, got {other:?}"),
        }
    }

    #[test]
    fn supports_strict_color_map_namespaces() {
        const STRICT_P: &str = "http://purl.oclc.org/ooxml/presentationml/main";
        const STRICT_A: &str = "http://purl.oclc.org/ooxml/drawingml/main";

        let master = parse_master(
            format!(r#"<q:sldMaster xmlns:q="{STRICT_P}"><q:clrMap {MAP}/></q:sldMaster>"#)
                .as_bytes(),
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(
            master.color(Slot::FollowedHyperlink),
            Role::FollowedHyperlink
        );

        let slide = parse_override(
            format!(
                r#"<q:sld xmlns:q="{STRICT_P}" xmlns:d="{STRICT_A}"><q:clrMapOvr>
                <d:masterClrMapping/></q:clrMapOvr></q:sld>"#
            )
            .as_bytes(),
            b"sld",
            "slide",
            Limits::DEFAULT,
        )
        .unwrap();
        assert_eq!(slide, Some(Override::Master));
    }
}
