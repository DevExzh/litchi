//! Bounded PresentationML transition reader.

use std::sync::Arc;

use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

use super::model::{
    Axis, Corner, InOut, Kind, Ms, Origin, Preserved, Raw, Ripple, Shape, Side, Speed, Spokes,
    Transition,
};

const PRESENTATIONML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
pub(super) const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P14_BYTES: &[u8] = P14.as_bytes();

/// Resource limits for one transition read.
///
/// The ordinary [`read`](super::read) entry point uses [`Limits::DEFAULT`].
/// Advanced callers may lower or raise finite limits explicitly without
/// changing the semantic API.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    depth: usize,
    nodes: usize,
    retained_bytes: usize,
}

impl Limits {
    /// Conservative defaults for one slide, layout, or master XML part.
    pub const DEFAULT: Self = Self {
        input_bytes: 64 * 1024 * 1024,
        depth: 128,
        nodes: 1_000_000,
        retained_bytes: 8 * 1024 * 1024,
    };

    /// Creates a finite, nonzero limit set.
    ///
    /// Returns `None` when any limit is zero.
    pub const fn new(
        input_bytes: usize,
        depth: usize,
        nodes: usize,
        retained_bytes: usize,
    ) -> Option<Self> {
        if input_bytes == 0 || depth == 0 || nodes == 0 || retained_bytes == 0 {
            None
        } else {
            Some(Self {
                input_bytes,
                depth,
                nodes,
                retained_bytes,
            })
        }
    }

    /// Maximum input and post-MCE output size.
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    /// Maximum XML nesting depth.
    pub const fn depth(self) -> usize {
        self.depth
    }

    /// Maximum number of start or empty elements.
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    /// Maximum total bytes retained for opaque child subtrees.
    pub const fn retained_bytes(self) -> usize {
        self.retained_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Reads the first slide transition using bounded default resources.
pub fn read(xml: &[u8]) -> Result<Option<Transition>> {
    read_with(xml, Limits::DEFAULT)
}

/// Reads the first slide transition using explicit finite resource limits.
pub fn read_with(xml: &[u8], limits: Limits) -> Result<Option<Transition>> {
    if xml.len() > limits.input_bytes {
        return Err(Error::Limit {
            resource: "input bytes",
            limit: limits.input_bytes,
        });
    }

    let mut capabilities = Capabilities::ooxml_baseline();
    capabilities.understand_namespace(P14);
    let mce_limits = litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.input_bytes,
        max_output_bytes: limits.input_bytes,
        max_depth: limits.depth,
        ..litchi_ooxml_common::mce::Limits::default()
    };
    let processed = process_markup_compatibility(xml, &capabilities, &mce_limits)?.xml;
    if processed.len() > limits.input_bytes {
        return Err(Error::Limit {
            resource: "post-MCE bytes",
            limit: limits.input_bytes,
        });
    }

    let bytes = processed.as_ref();
    let mut reader = NsReader::from_reader(bytes);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut selected_depth = None;
    let mut draft = None;
    let mut capture = None;

    loop {
        let start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader.read_event()?.into_owned();
        let end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                nodes = count_node(nodes, limits.nodes)?;
                let event_depth = enter_depth(depth, limits.depth)?;

                if draft.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"transition")
                {
                    draft = Some(Draft::new(parse_attributes(&element, decoder, &resolver)?));
                    selected_depth = Some(event_depth);
                } else if selected_depth.and_then(|value| value.checked_add(1)) == Some(event_depth)
                {
                    let role = classify_child(
                        &namespace,
                        &element,
                        decoder,
                        draft.as_mut().ok_or_else(|| {
                            Error::Invalid("transition parser lost its selected value".into())
                        })?,
                    )?;
                    capture = Some(Capture {
                        start,
                        depth: event_depth,
                        role,
                    });
                }

                depth = event_depth;
            },
            Event::Empty(element) => {
                nodes = count_node(nodes, limits.nodes)?;
                let event_depth = enter_depth(depth, limits.depth)?;

                if draft.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"transition")
                {
                    draft = Some(Draft::new(parse_attributes(&element, decoder, &resolver)?));
                } else if selected_depth.and_then(|value| value.checked_add(1)) == Some(event_depth)
                {
                    let role = classify_child(
                        &namespace,
                        &element,
                        decoder,
                        draft.as_mut().ok_or_else(|| {
                            Error::Invalid("transition parser lost its selected value".into())
                        })?,
                    )?;
                    let raw = bytes.get(start..end).ok_or_else(|| {
                        Error::Invalid("transition child byte range is outside its XML part".into())
                    })?;
                    draft
                        .as_mut()
                        .ok_or_else(|| {
                            Error::Invalid("transition parser lost its selected value".into())
                        })?
                        .finish_raw(role, raw, limits.retained_bytes)?;
                }
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(Error::Invalid(
                        "transition XML contains an unmatched end element".into(),
                    ));
                }

                if capture.as_ref().is_some_and(|active| active.depth == depth) {
                    let active = capture.take().ok_or_else(|| {
                        Error::Invalid("transition child capture state is inconsistent".into())
                    })?;
                    let raw = bytes.get(active.start..end).ok_or_else(|| {
                        Error::Invalid("transition child byte range is outside its XML part".into())
                    })?;
                    draft
                        .as_mut()
                        .ok_or_else(|| {
                            Error::Invalid("transition parser lost its selected value".into())
                        })?
                        .finish_raw(active.role, raw, limits.retained_bytes)?;
                }

                if selected_depth == Some(depth) {
                    selected_depth = None;
                }
                depth -= 1;
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DOCTYPE is forbidden in transition XML".into(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }

    if depth != 0 || capture.is_some() {
        return Err(Error::Invalid(
            "transition XML ended with unclosed elements".into(),
        ));
    }

    Ok(draft.map(Draft::finish))
}

#[derive(Debug)]
struct Draft {
    value: Transition,
    has_effect: bool,
    effect: Option<Raw>,
    before: Vec<Raw>,
    after: Vec<Raw>,
    retained: usize,
}

impl Draft {
    fn new(value: Transition) -> Self {
        Self {
            value,
            has_effect: false,
            effect: None,
            before: Vec::new(),
            after: Vec::new(),
            retained: 0,
        }
    }

    fn finish_raw(&mut self, role: Role, bytes: &[u8], limit: usize) -> Result<()> {
        let retained = self.retained.checked_add(bytes.len()).ok_or(Error::Limit {
            resource: "retained transition bytes",
            limit,
        })?;
        if retained > limit {
            return Err(Error::Limit {
                resource: "retained transition bytes",
                limit,
            });
        }
        let xml = std::str::from_utf8(bytes).map_err(|error| {
            Error::Xml(format!(
                "opaque transition child is not UTF-8 at byte {}",
                error.valid_up_to()
            ))
        })?;
        let raw = Raw {
            xml: Arc::<str>::from(xml),
            portable: raw_is_portable(xml)?,
        };
        self.retained = retained;

        match role {
            Role::Known => self.effect = Some(raw),
            Role::RawEffect => self.value.kind = Kind::Raw(raw),
            Role::Before => self.before.push(raw),
            Role::After => self.after.push(raw),
        }
        Ok(())
    }

    fn finish(mut self) -> Transition {
        if self.effect.is_some() || !self.before.is_empty() || !self.after.is_empty() {
            self.value.preserved = Some(Arc::new(Preserved {
                effect: self.effect,
                before: self.before.into_boxed_slice(),
                after: self.after.into_boxed_slice(),
            }));
        }
        self.value
    }
}

#[derive(Debug, Clone, Copy)]
enum Role {
    Known,
    RawEffect,
    Before,
    After,
}

#[derive(Debug)]
struct Capture {
    start: usize,
    depth: usize,
    role: Role,
}

fn classify_child(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    draft: &mut Draft,
) -> Result<Role> {
    if let Some(kind) = parse_kind(namespace, element, decoder)? {
        if draft.has_effect {
            return Err(Error::Invalid(
                "a transition contains more than one visual effect".into(),
            ));
        }
        draft.has_effect = true;
        draft.value.kind = kind;
        return Ok(Role::Known);
    }

    if is_auxiliary(namespace, element.name()) {
        return Ok(if draft.has_effect {
            Role::After
        } else {
            Role::Before
        });
    }

    if draft.has_effect {
        Ok(Role::After)
    } else {
        draft.has_effect = true;
        Ok(Role::RawEffect)
    }
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Transition> {
    let mut speed = Speed::Medium;
    let mut legacy_duration = None;
    let mut extended_duration = None;
    let mut click = true;
    let mut after = None;
    let mut seen_speed = false;
    let mut seen_legacy_duration = false;
    let mut seen_extended_duration = false;
    let mut seen_click = false;
    let mut seen_after = false;

    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let value = value.as_ref();

        if key.prefix().is_none() {
            match key.local_name().as_ref() {
                b"spd" => {
                    reject_duplicate(&mut seen_speed, "spd")?;
                    speed = parse_speed(value)?;
                },
                b"dur" => {
                    reject_duplicate(&mut seen_legacy_duration, "dur")?;
                    legacy_duration = Some(parse_ms(value, "transition duration")?);
                },
                b"advClick" => {
                    reject_duplicate(&mut seen_click, "advClick")?;
                    click = parse_bool(value, "advClick")?;
                },
                b"advTm" => {
                    reject_duplicate(&mut seen_after, "advTm")?;
                    after = Some(parse_advance_ms(value)?);
                },
                _ => {},
            }
        } else {
            let (namespace, _) = resolver.resolve_attribute(key);
            if is_p14_namespace(&namespace) && key.local_name().as_ref() == b"dur" {
                reject_duplicate(&mut seen_extended_duration, "p14:dur")?;
                extended_duration = Some(parse_ms(value, "transition duration")?);
            }
        }
    }

    let mut value = Transition::new(Kind::None);
    value.speed = speed;
    value.duration = extended_duration.or(legacy_duration);
    value.click = click;
    value.after = after;
    Ok(value)
}

fn parse_kind(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
) -> Result<Option<Kind>> {
    if is_p14_name(namespace, element.name(), b"ripple") {
        let value = unqualified_attribute_value(element, b"dir", decoder)?
            .unwrap_or_else(|| "center".to_string());
        return Ok(Some(Kind::Ripple(parse_ripple(&value)?)));
    }

    let local = element.local_name();
    if !is_presentationml_name(namespace, element.name(), local.as_ref()) {
        return Ok(None);
    }

    let kind = match local.as_ref() {
        b"cut" => Kind::Cut {
            black: parse_optional_bool(element, b"thruBlk", decoder)?,
        },
        b"fade" => Kind::Fade {
            black: parse_optional_bool(element, b"thruBlk", decoder)?,
        },
        b"push" => Kind::Push(parse_side(element, decoder)?),
        b"wipe" => Kind::Wipe(parse_side(element, decoder)?),
        b"split" => Kind::Split {
            axis: parse_axis(element, b"orient", decoder)?,
            toward: parse_optional_in_out(element, b"dir", decoder)?,
        },
        b"pull" => Kind::Uncover(parse_origin(element, decoder)?),
        b"cover" => Kind::Cover(parse_origin(element, decoder)?),
        b"dissolve" => Kind::Dissolve,
        b"blinds" => Kind::Blinds(parse_axis(element, b"dir", decoder)?),
        b"checker" => Kind::Checker(parse_axis(element, b"dir", decoder)?),
        b"randomBar" => Kind::RandomBars(parse_axis(element, b"dir", decoder)?),
        b"circle" => Kind::Shape(Shape::Circle),
        b"diamond" => Kind::Shape(Shape::Diamond),
        b"plus" => Kind::Shape(Shape::Plus),
        b"wedge" => Kind::Wedge,
        b"zoom" => Kind::Zoom(parse_in_out(element, b"dir", "in", decoder)?),
        b"wheel" => Kind::Wheel(parse_spokes(element, decoder)?),
        b"random" => Kind::Random,
        b"newsflash" => Kind::Newsflash,
        b"strips" => Kind::Strips(parse_corner(element, decoder)?),
        b"comb" => Kind::Comb(parse_axis(element, b"dir", decoder)?),
        _ => return Ok(None),
    };
    Ok(Some(kind))
}

fn parse_speed(value: &str) -> Result<Speed> {
    match value {
        "slow" => Ok(Speed::Slow),
        "med" => Ok(Speed::Medium),
        "fast" => Ok(Speed::Fast),
        _ => Err(Error::Invalid(format!(
            "invalid transition speed '{value}'"
        ))),
    }
}

fn parse_ms(value: &str, field: &str) -> Result<Ms> {
    let digits = value.strip_suffix("ms").unwrap_or(value);
    parse_bounded_ms(digits, field)
}

fn parse_advance_ms(value: &str) -> Result<Ms> {
    parse_bounded_ms(value, "automatic-advance delay")
}

fn parse_bounded_ms(value: &str, field: &str) -> Result<Ms> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_| Error::Invalid(format!("invalid {field} '{value}'")))?;
    let parsed = u32::try_from(parsed)
        .map_err(|_| Error::Invalid(format!("{field} '{value}' is outside its domain")))?;
    Ms::new(parsed).map_err(|error| Error::Invalid(format!("invalid {field}: {error}")))
}

fn parse_bool(value: &str, attribute: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(Error::Invalid(format!(
            "invalid boolean value '{value}' for transition attribute '{attribute}'"
        ))),
    }
}

fn parse_optional_bool(
    element: &BytesStart<'_>,
    attribute: &[u8],
    decoder: Decoder,
) -> Result<Option<bool>> {
    unqualified_attribute_value(element, attribute, decoder)?
        .map(|value| parse_bool(&value, &String::from_utf8_lossy(attribute)))
        .transpose()
}

fn parse_side(element: &BytesStart<'_>, decoder: Decoder) -> Result<Side> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "l".to_string());
    match value.as_str() {
        "l" => Ok(Side::Left),
        "r" => Ok(Side::Right),
        "u" => Ok(Side::Up),
        "d" => Ok(Side::Down),
        _ => invalid_direction("side", &value),
    }
}

fn parse_axis(element: &BytesStart<'_>, attribute: &[u8], decoder: Decoder) -> Result<Axis> {
    let value = unqualified_attribute_value(element, attribute, decoder)?
        .unwrap_or_else(|| "horz".to_string());
    match value.as_str() {
        "horz" => Ok(Axis::Horizontal),
        "vert" => Ok(Axis::Vertical),
        _ => invalid_direction("axis", &value),
    }
}

fn parse_corner(element: &BytesStart<'_>, decoder: Decoder) -> Result<Corner> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "lu".to_string());
    match value.as_str() {
        "lu" => Ok(Corner::LeftUp),
        "ru" => Ok(Corner::RightUp),
        "ld" => Ok(Corner::LeftDown),
        "rd" => Ok(Corner::RightDown),
        _ => invalid_direction("corner", &value),
    }
}

fn parse_origin(element: &BytesStart<'_>, decoder: Decoder) -> Result<Origin> {
    let value =
        unqualified_attribute_value(element, b"dir", decoder)?.unwrap_or_else(|| "l".to_string());
    match value.as_str() {
        "l" => Ok(Origin::Left),
        "r" => Ok(Origin::Right),
        "u" => Ok(Origin::Up),
        "d" => Ok(Origin::Down),
        "lu" => Ok(Origin::LeftUp),
        "ru" => Ok(Origin::RightUp),
        "ld" => Ok(Origin::LeftDown),
        "rd" => Ok(Origin::RightDown),
        _ => invalid_direction("origin", &value),
    }
}

fn parse_in_out(
    element: &BytesStart<'_>,
    attribute: &[u8],
    default: &str,
    decoder: Decoder,
) -> Result<InOut> {
    let value = unqualified_attribute_value(element, attribute, decoder)?
        .unwrap_or_else(|| default.to_string());
    match value.as_str() {
        "in" => Ok(InOut::In),
        "out" => Ok(InOut::Out),
        _ => invalid_direction("in/out", &value),
    }
}

fn parse_optional_in_out(
    element: &BytesStart<'_>,
    attribute: &[u8],
    decoder: Decoder,
) -> Result<Option<InOut>> {
    unqualified_attribute_value(element, attribute, decoder)?
        .map(|value| match value.as_str() {
            "in" => Ok(InOut::In),
            "out" => Ok(InOut::Out),
            _ => invalid_direction("in/out", &value),
        })
        .transpose()
}

fn parse_ripple(value: &str) -> Result<Ripple> {
    match value {
        "center" => Ok(Ripple::Center),
        "lu" => Ok(Ripple::LeftUp),
        "ru" => Ok(Ripple::RightUp),
        "ld" => Ok(Ripple::LeftDown),
        "rd" => Ok(Ripple::RightDown),
        _ => invalid_direction("PowerPoint 2010 ripple", value),
    }
}

fn parse_spokes(element: &BytesStart<'_>, decoder: Decoder) -> Result<Spokes> {
    let value = unqualified_attribute_value(element, b"spokes", decoder)?
        .unwrap_or_else(|| "4".to_string());
    match value.as_str() {
        "1" => Ok(Spokes::One),
        "2" => Ok(Spokes::Two),
        "3" => Ok(Spokes::Three),
        "4" => Ok(Spokes::Four),
        "8" => Ok(Spokes::Eight),
        _ => Err(Error::Invalid(format!(
            "wheel transition spoke count '{value}' is not supported by PowerPoint"
        ))),
    }
}

fn invalid_direction<T>(kind: &str, value: &str) -> Result<T> {
    Err(Error::Invalid(format!(
        "invalid {kind} transition direction '{value}'"
    )))
}

fn reject_duplicate(seen: &mut bool, attribute: &str) -> Result<()> {
    if *seen {
        return Err(Error::Invalid(format!(
            "duplicate transition attribute '{attribute}'"
        )));
    }
    *seen = true;
    Ok(())
}

fn count_node(nodes: usize, limit: usize) -> Result<usize> {
    let nodes = nodes.checked_add(1).ok_or(Error::Limit {
        resource: "element count",
        limit,
    })?;
    if nodes > limit {
        Err(Error::Limit {
            resource: "element count",
            limit,
        })
    } else {
        Ok(nodes)
    }
}

fn enter_depth(depth: usize, limit: usize) -> Result<usize> {
    let depth = depth.checked_add(1).ok_or(Error::Limit {
        resource: "nesting depth",
        limit,
    })?;
    if depth > limit {
        Err(Error::Limit {
            resource: "nesting depth",
            limit,
        })
    } else {
        Ok(depth)
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::Invalid("transition XML position does not fit usize".into()))
}

fn is_auxiliary(namespace: &ResolveResult<'_>, name: QName<'_>) -> bool {
    is_presentationml_name(namespace, name, b"sndAc")
        || is_presentationml_name(namespace, name, b"extLst")
}

fn is_presentationml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    if name.local_name().as_ref() != local {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == PRESENTATIONML || *value == STRICT_PRESENTATIONML
        },
        // Standalone fragments commonly inherit the conventional prefix from
        // their slide root.
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}

fn is_p14_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P14_BYTES)
}

fn is_p14_name(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local && is_p14_namespace(namespace)
}

fn raw_is_portable(xml: &str) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        let event = reader.read_event()?.into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) | Event::Empty(element) => {
                if unknown_prefix_is_nonportable(&namespace) {
                    return Ok(false);
                }
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    let (namespace, _) = resolver.resolve_attribute(attribute.key);
                    if unknown_prefix_is_nonportable(&namespace) {
                        return Ok(false);
                    }
                }
            },
            Event::End(_) => {
                if unknown_prefix_is_nonportable(&namespace) {
                    return Ok(false);
                }
            },
            Event::DocType(_) => {
                return Err(Error::Invalid(
                    "DOCTYPE is forbidden in retained transition XML".into(),
                ));
            },
            Event::Eof => return Ok(true),
            _ => {},
        }
    }
}

fn unknown_prefix_is_nonportable(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Unknown(prefix)
            if !matches!(prefix.as_slice(), b"p" | b"a" | b"r")
    )
}
