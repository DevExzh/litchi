//! Bounded scalar and structural validation for `PresentationML` timing values.

use super::super::model::{
    Direction, Duration, Effect, GroupId, MotionFraction, ParagraphBuildType, Repeat, Sequence,
    SequenceContext,
};
use super::super::{invalid, limit};
use crate::animations::Limits;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::HashSet;

pub(crate) const MAX_TIMING_XML_BYTES: usize = 64 * 1024 * 1024;
pub(crate) const MAX_PRESERVED_TIMING_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_TIMING_DEPTH: usize = 128;
pub(crate) const MAX_TIMING_NODES: usize = 250_000;
pub(crate) const MAX_TIMING_TEXT_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_TIMING_ATTRIBUTES: usize = 64;
pub(crate) const MAX_ANIMATIONS: usize = 10_000;
pub(crate) const MAX_ANIMATION_BUILDS: usize = 10_000;
pub(crate) const MAX_PARAGRAPH_TEMPLATES: usize = 9;
pub(crate) const MAX_TEMPLATE_TIME_NODE_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_TIME_FILTER_BYTES: usize = 64 * 1024;
pub(crate) const MAX_TIME_FILTER_POINTS: usize = 4_096;
pub(crate) const MAX_NORMALIZED_TIME_DECIMALS: usize = 18;
pub const MAX_TIMING_MILLISECONDS: u32 = 2_147_483_625;
const DRAWINGML_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWINGML_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const CHART_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_STRICT_NS: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
pub(super) const DIAGRAM_NS: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const PRESENTATIONML_NS: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PRESENTATIONML_NS: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";

pub(super) fn is_presentationml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(value) => {
            value.as_ref() == PRESENTATIONML_NS || value.as_ref() == STRICT_PRESENTATIONML_NS
        },
        ResolveResult::Unknown(prefix) => prefix.as_slice() == b"p",
        ResolveResult::Unbound => false,
    }
}
#[derive(Clone, Copy)]
pub(super) enum TimingValue {
    Indefinite,
    Milliseconds(u32),
}

pub(super) fn parse_group_id(value: &str) -> Result<GroupId> {
    value
        .parse::<u32>()
        .map(GroupId::new)
        .map_err(|_err| invalid("invalid unsigned animation group ID"))
}

pub(super) fn parse_build_auto_advance(value: &str) -> Result<Duration> {
    if value == "indefinite" {
        return Ok(Duration::Indefinite);
    }
    value
        .parse::<u32>()
        .map(Duration::Finite)
        .map_err(|_err| invalid("invalid paragraph build auto-advance time"))
}

pub(super) fn validate_template_time_node(xml: &str, limits: Limits) -> Result<()> {
    if xml.len() > limits.max_template_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            xml.len(),
            limits.max_template_bytes,
        ));
    }
    let wrapped = format!(
        r#"<root xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{xml}</root>"#
    );
    let mut reader = NsReader::from_reader(wrapped.as_bytes());
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;
    let mut saw_par = false;
    let mut saw_ctn = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        limits.max_depth,
                    )
                })?;
                if depth > limits.max_depth {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        depth,
                        limits.max_depth,
                    ));
                }
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes,
                    )
                })?;
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                check_attribute_count(&element, limits)?;
                if depth == 2 {
                    if saw_par || !is_presentationml_name(&namespace, element.name(), b"par") {
                        return Err(invalid(
                            "paragraph template must contain exactly one par root",
                        ));
                    }
                    saw_par = true;
                } else if depth == 3 {
                    if saw_ctn || !is_presentationml_name(&namespace, element.name(), b"cTn") {
                        return Err(invalid(
                            "paragraph template par must contain exactly one cTn",
                        ));
                    }
                    saw_ctn = true;
                }
            },
            Event::Empty(element) => {
                nodes = nodes.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes,
                    )
                })?;
                if nodes > limits.max_nodes {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        nodes,
                        limits.max_nodes,
                    ));
                }
                check_attribute_count(&element, limits)?;
                let element_depth = depth.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        usize::MAX,
                        limits.max_depth,
                    )
                })?;
                if element_depth > limits.max_depth {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        element_depth,
                        limits.max_depth,
                    ));
                }
                if element_depth == 2 {
                    return Err(invalid("paragraph template par node cannot be empty"));
                }
                if element_depth == 3 {
                    if saw_ctn || !is_presentationml_name(&namespace, element.name(), b"cTn") {
                        return Err(invalid(
                            "paragraph template par must contain exactly one cTn",
                        ));
                    }
                    saw_ctn = true;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced paragraph template XML"))?;
            },
            Event::Text(text) => {
                text_bytes = text_bytes.checked_add(text.len()).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::TextBytes,
                        usize::MAX,
                        limits.max_text_bytes,
                    )
                })?;
                if text_bytes > limits.max_text_bytes {
                    return Err(limit(
                        crate::error::CodecLimitKind::TextBytes,
                        text_bytes,
                        limits.max_text_bytes,
                    ));
                }
                if depth <= 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("paragraph template has text outside its par root"));
                }
            },
            Event::CData(text) => {
                text_bytes = text_bytes.checked_add(text.len()).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::TextBytes,
                        usize::MAX,
                        limits.max_text_bytes,
                    )
                })?;
                if text_bytes > limits.max_text_bytes {
                    return Err(limit(
                        crate::error::CodecLimitKind::TextBytes,
                        text_bytes,
                        limits.max_text_bytes,
                    ));
                }
                if depth <= 1 && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(invalid("paragraph template has text outside its par root"));
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid(
                    "active XML constructs are not allowed in paragraph templates",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !saw_par || !saw_ctn {
        return Err(invalid("incomplete paragraph template time node"));
    }
    Ok(())
}

pub(super) fn direction_subtype(effect: &Effect, direction: &Direction) -> Option<u32> {
    match effect {
        Effect::FlyIn | Effect::Wipe => Some(match direction {
            Direction::Up => 1,
            Direction::Right => 2,
            Direction::UpRight => 3,
            Direction::Down => 4,
            Direction::DownRight => 6,
            Direction::Left => 8,
            Direction::UpLeft => 9,
            Direction::DownLeft => 12,
            _ => return None,
        }),
        Effect::Split => Some(match direction {
            Direction::VerticalIn => 21,
            Direction::HorizontalIn => 26,
            Direction::VerticalOut => 37,
            Direction::HorizontalOut => 42,
            _ => return None,
        }),
        Effect::Zoom => Some(match direction {
            Direction::In => 16,
            Direction::Out => 32,
            Direction::OutFromScreenCenter => 36,
            Direction::InSlightly => 272,
            Direction::OutSlightly => 288,
            Direction::InFromScreenCenter => 528,
            _ => return None,
        }),
        _ => None,
    }
}

pub(super) fn direction_from_subtype(effect: &Effect, subtype: u32) -> Option<Direction> {
    match effect {
        Effect::FlyIn | Effect::Wipe => match subtype {
            1 => Some(Direction::Up),
            2 => Some(Direction::Right),
            3 => Some(Direction::UpRight),
            4 => Some(Direction::Down),
            6 => Some(Direction::DownRight),
            8 => Some(Direction::Left),
            9 => Some(Direction::UpLeft),
            12 => Some(Direction::DownLeft),
            _ => None,
        },
        Effect::Split => match subtype {
            21 => Some(Direction::VerticalIn),
            26 => Some(Direction::HorizontalIn),
            37 => Some(Direction::VerticalOut),
            42 => Some(Direction::HorizontalOut),
            _ => None,
        },
        Effect::Zoom => match subtype {
            16 => Some(Direction::In),
            32 => Some(Direction::Out),
            36 => Some(Direction::OutFromScreenCenter),
            272 => Some(Direction::InSlightly),
            288 => Some(Direction::OutSlightly),
            528 => Some(Direction::InFromScreenCenter),
            _ => None,
        },
        _ => None,
    }
}

pub(super) fn parse_timing_value(value: &str) -> Result<TimingValue> {
    if value == "indefinite" {
        return Ok(TimingValue::Indefinite);
    }
    let value = value
        .parse::<u32>()
        .map_err(|_err| invalid("invalid animation timing value"))?;
    if value > MAX_TIMING_MILLISECONDS {
        return Err(invalid(
            "animation timing value exceeds the supported OOXML limit",
        ));
    }
    Ok(TimingValue::Milliseconds(value))
}

pub(super) fn parse_xml_bool(value: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid("invalid animation boolean value")),
    }
}

pub(super) fn parse_progress(value: &str, name: &str) -> Result<MotionFraction> {
    let value = value
        .parse::<u32>()
        .map_err(|_err| invalid(format!("invalid animation {name} percentage")))?;
    MotionFraction::new(value)
}

pub(super) fn parse_shape_id(value: &str) -> Result<u32> {
    let id = value
        .parse::<u32>()
        .map_err(|_err| invalid("invalid animation shape target ID"))?;
    if id == 0 {
        return Err(invalid("animation shape target ID must be nonzero"));
    }
    Ok(id)
}

pub(super) fn attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
) -> Result<Option<String>> {
    Ok(litchi_ooxml_common::xml::unqualified_attribute_value(
        element, name, decoder,
    )?)
}

pub(super) fn check_attribute_count(element: &BytesStart<'_>, limits: Limits) -> Result<()> {
    let mut count = 0usize;
    let mut bytes = 0usize;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count.checked_add(1).ok_or_else(|| {
            limit(
                crate::error::CodecLimitKind::XmlAttributes,
                usize::MAX,
                limits.max_attributes,
            )
        })?;
        if count > limits.max_attributes {
            return Err(limit(
                crate::error::CodecLimitKind::XmlAttributes,
                count,
                limits.max_attributes,
            ));
        }
        bytes = bytes
            .checked_add(attribute.key.as_ref().len())
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or_else(|| {
                limit(
                    crate::error::CodecLimitKind::XmlAttributeBytes,
                    usize::MAX,
                    limits.max_text_bytes,
                )
            })?;
        if bytes > limits.max_text_bytes {
            return Err(limit(
                crate::error::CodecLimitKind::XmlAttributeBytes,
                bytes,
                limits.max_text_bytes,
            ));
        }
    }
    Ok(())
}

pub(super) fn check_xml_size(size: usize, limits: Limits) -> Result<()> {
    if size > limits.max_input_bytes {
        Err(limit(
            crate::error::CodecLimitKind::XmlBytes,
            size,
            limits.max_input_bytes,
        ))
    } else {
        Ok(())
    }
}

pub(super) fn is_namespace_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    expected_namespace: &[u8],
    expected_local_name: &[u8],
) -> bool {
    matches!(namespace, ResolveResult::Bound(value) if value.as_ref() == expected_namespace)
        && name.local_name().as_ref() == expected_local_name
}

pub(super) fn is_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local: &[u8],
) -> bool {
    is_namespace_name(namespace, name, DRAWINGML_NS, local)
        || is_namespace_name(namespace, name, DRAWINGML_STRICT_NS, local)
}

pub(super) fn is_chartml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local: &[u8],
) -> bool {
    is_namespace_name(namespace, name, CHART_NS, local)
        || is_namespace_name(namespace, name, CHART_STRICT_NS, local)
}

pub(super) fn is_known_ole_chart_program_id(value: &str) -> bool {
    value == "Excel.Chart"
        || value.starts_with("Excel.Chart.")
        || value == "MSGraph.Chart"
        || value.starts_with("MSGraph.Chart.")
}

impl Sequence {
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml_for_slide(&self, valid_targets: &HashSet<u32>, limits: Limits) -> Result<String> {
        if self.len() > MAX_ANIMATIONS {
            return Err(limit(
                crate::error::CodecLimitKind::Animations,
                self.len(),
                MAX_ANIMATIONS,
            ));
        }
        let build_count = self
            .paragraph_builds
            .len()
            .checked_add(self.diagram_builds.len())
            .and_then(|count| count.checked_add(self.graphic_builds.len()))
            .and_then(|count| count.checked_add(self.ole_chart_builds.len()))
            .ok_or_else(|| {
                limit(
                    crate::error::CodecLimitKind::AnimationBuilds,
                    usize::MAX,
                    MAX_ANIMATION_BUILDS,
                )
            })?;
        if build_count > MAX_ANIMATION_BUILDS {
            return Err(limit(
                crate::error::CodecLimitKind::AnimationBuilds,
                build_count,
                MAX_ANIMATION_BUILDS,
            ));
        }
        let animation_groups: HashSet<_> = self
            .animations
            .iter()
            .filter_map(|animation| animation.group_id)
            .collect();
        let mut build_groups = HashSet::new();
        let mut build_pairs = HashSet::new();
        for build in &self.paragraph_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "paragraph build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !build_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate paragraph build shape/group pair"));
            }
            if build.build_type != ParagraphBuildType::Paragraph && build.build_level != 1 {
                return Err(invalid(
                    "non-default paragraph build level requires build type p",
                ));
            }
            if build.reverse && build.build_type != ParagraphBuildType::Paragraph {
                return Err(invalid("reverse paragraph order requires build type p"));
            }
            if build.templates.len() > MAX_PARAGRAPH_TEMPLATES {
                return Err(limit(
                    crate::error::CodecLimitKind::ParagraphTemplates,
                    build.templates.len(),
                    MAX_PARAGRAPH_TEMPLATES,
                ));
            }
            let mut levels = HashSet::new();
            for template in &build.templates {
                if template.level > 9 {
                    return Err(invalid("paragraph template level exceeds PowerPoint limit"));
                }
                if !levels.insert(template.level) {
                    return Err(invalid("duplicate paragraph template level"));
                }
            }
            if build.build_type == ParagraphBuildType::Whole && build.templates.len() > 1 {
                return Err(invalid(
                    "whole paragraph builds support exactly one template effect",
                ));
            }
            build_groups.insert(build.group_id);
        }
        let mut diagram_pairs = HashSet::new();
        for build in &self.diagram_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "diagram build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !diagram_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate diagram build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        let mut graphic_pairs = HashSet::new();
        for build in &self.graphic_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "graphical-object build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !graphic_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate graphical-object build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        let mut ole_chart_pairs = HashSet::new();
        for build in &self.ole_chart_builds {
            if build.shape_id == 0 || !valid_targets.contains(&build.shape_id) {
                return Err(invalid(format!(
                    "OLE chart build target {} is not a supported shape on the current slide",
                    build.shape_id
                )));
            }
            if !ole_chart_pairs.insert((build.shape_id, build.group_id)) {
                return Err(invalid("duplicate OLE chart build shape/group pair"));
            }
            build_groups.insert(build.group_id);
        }
        if animation_groups != build_groups {
            return Err(invalid(
                "animation cTn group IDs and paragraph build group IDs do not match",
            ));
        }
        for animation in &self.animations {
            if animation.shape_id == 0 || !valid_targets.contains(&animation.shape_id) {
                return Err(invalid(format!(
                    "animation target {} is not a supported shape on the current slide",
                    animation.shape_id
                )));
            }
            if let SequenceContext::Interactive {
                trigger_shape_id, ..
            } = &animation.sequence_context
                && (*trigger_shape_id == 0 || !valid_targets.contains(trigger_shape_id))
            {
                return Err(invalid(format!(
                    "interactive animation trigger {trigger_shape_id} is not a supported shape on the current slide"
                )));
            }
            if animation.delay > MAX_TIMING_MILLISECONDS {
                return Err(invalid("animation delay exceeds the supported OOXML limit"));
            }
            if let Duration::Finite(duration) = animation.duration
                && duration > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation duration exceeds the supported OOXML limit",
                ));
            }
            if let Some(direction) = &animation.direction
                && direction_subtype(&animation.effect, direction).is_none()
            {
                return Err(invalid(
                    "animation direction is not supported for this animation effect",
                ));
            }
            if let Some(Repeat::Finite(repeat)) = animation.repeat
                && repeat > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation repeat count exceeds the supported OOXML limit",
                ));
            }
            if let Some(Duration::Finite(repeat_duration)) = animation.repeat_duration
                && repeat_duration > MAX_TIMING_MILLISECONDS
            {
                return Err(invalid(
                    "animation repeat duration exceeds the supported OOXML limit",
                ));
            }
            if let Some(time_filter) = &animation.time_filter
                && time_filter.write_value().len() > MAX_TIME_FILTER_BYTES
            {
                return Err(limit(
                    crate::error::CodecLimitKind::XmlBytes,
                    time_filter.write_value().len(),
                    MAX_TIME_FILTER_BYTES,
                ));
            }
        }
        self.to_xml(limits)
    }
}
