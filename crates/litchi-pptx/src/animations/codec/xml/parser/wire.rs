use super::super::super::super::{invalid, model::*};
use super::semantic::TimingParser;
use super::validation::*;
use crate::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

struct RecursiveNodeFrame {
    depth: usize,
    sub_node: bool,
    node: TimingNode,
}

pub(super) fn parse_recursive_timing_tree(xml: &str) -> Result<TimingTree> {
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut count = 0usize;
    let mut timing_depth = None;
    let mut timing_start = None;
    let mut frames = Vec::<RecursiveNodeFrame>::new();
    let mut roots = Vec::new();
    let mut child_lists = Vec::<(usize, bool)>::new();
    let mut condition_lists = Vec::<(usize, bool)>::new();
    let mut condition: Option<(usize, bool, TimeCondition)> = None;
    let mut source_range = None;
    loop {
        let event_start = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                depth += 1;
                count += 1;
                if depth > MAX_TIMING_DEPTH || count > MAX_TIMING_NODES {
                    return Err(invalid("animation timing tree exceeds safety limit"));
                }
                check_attribute_count(element)?;
                let empty = matches!(event, Event::Empty(_));
                if timing_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"timing")
                {
                    timing_depth = Some(depth);
                    timing_start = Some(event_start);
                } else if timing_depth.is_some() {
                    let local = element.local_name();
                    if is_presentationml_name(&namespace, element.name(), b"par")
                        || is_presentationml_name(&namespace, element.name(), b"seq")
                        || is_presentationml_name(&namespace, element.name(), b"excl")
                    {
                        if empty {
                            return Err(invalid("animation time container cannot be empty"));
                        }
                        let kind = if local.as_ref() == b"seq" {
                            let concurrent = attribute(element, b"concurrent", reader.decoder())?
                                .map(|v| parse_xml_bool(&v))
                                .transpose()?
                                .unwrap_or(false);
                            let next_action = match attribute(element, b"nextAc", reader.decoder())?
                                .as_deref()
                                .unwrap_or("none")
                            {
                                "none" => NextAction::None,
                                "seek" => NextAction::Seek,
                                _ => return Err(invalid("invalid animation next action")),
                            };
                            let previous_action =
                                match attribute(element, b"prevAc", reader.decoder())?
                                    .as_deref()
                                    .unwrap_or("none")
                                {
                                    "none" => PreviousAction::None,
                                    "skipTimed" => PreviousAction::SkipTimed,
                                    _ => return Err(invalid("invalid animation previous action")),
                                };
                            TimingNodeKind::Sequence {
                                concurrent,
                                next_action,
                                previous_action,
                            }
                        } else if local.as_ref() == b"excl" {
                            TimingNodeKind::Exclusive
                        } else {
                            TimingNodeKind::Parallel
                        };
                        frames.push(RecursiveNodeFrame {
                            depth,
                            sub_node: child_lists.last().is_some_and(|(_, sub)| *sub),
                            node: TimingNode {
                                kind,
                                common: CommonTimeNode {
                                    id: None,
                                    duration: None,
                                    node_type: None,
                                    preset: None,
                                    start_conditions: Vec::new(),
                                    end_conditions: Vec::new(),
                                    children: Vec::new(),
                                    sub_nodes: Vec::new(),
                                    opaque_children: Vec::new(),
                                },
                                opaque_children: Vec::new(),
                            },
                        });
                    } else if is_presentationml_name(&namespace, element.name(), b"cTn")
                        && frames.last().is_some_and(|frame| depth == frame.depth + 1)
                    {
                        let frame = frames
                            .last_mut()
                            .ok_or_else(|| invalid("common time node has no container"))?;
                        frame.node.common.id = attribute(element, b"id", reader.decoder())?
                            .map(|value| {
                                value
                                    .parse::<u32>()
                                    .map_err(|_| invalid("invalid common time-node ID"))
                            })
                            .transpose()?;
                        frame.node.common.duration = attribute(element, b"dur", reader.decoder())?
                            .map(|v| parse_timing_value(&v))
                            .transpose()?
                            .map(|v| match v {
                                TimingValue::Indefinite => Duration::Indefinite,
                                TimingValue::Milliseconds(ms) => Duration::Finite(ms),
                            });
                        frame.node.common.node_type =
                            attribute(element, b"nodeType", reader.decoder())?
                                .map(|v| TimeNodeType::parse(&v))
                                .transpose()?;
                        if let Some(value) = attribute(element, b"presetID", reader.decoder())? {
                            let preset_id = value
                                .parse::<u32>()
                                .map_err(|_| invalid("invalid animation preset ID"))?;
                            let class = PresetClass::parse(
                                attribute(element, b"presetClass", reader.decoder())?
                                    .as_deref()
                                    .unwrap_or("entr"),
                            )?;
                            let subtype = attribute(element, b"presetSubtype", reader.decoder())?
                                .map(|v| {
                                    v.parse::<u32>()
                                        .map_err(|_| invalid("invalid animation preset subtype"))
                                })
                                .transpose()?;
                            frame.node.common.preset = Some(PresetTimeNode {
                                preset_id,
                                class,
                                subtype,
                            });
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"childTnLst")
                        || is_presentationml_name(&namespace, element.name(), b"subTnLst")
                    {
                        if !empty {
                            child_lists.push((depth, local.as_ref() == b"subTnLst"));
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"stCondLst")
                        || is_presentationml_name(&namespace, element.name(), b"endCondLst")
                    {
                        if !empty {
                            condition_lists.push((depth, local.as_ref() == b"stCondLst"));
                        }
                    } else if is_presentationml_name(&namespace, element.name(), b"cond")
                        && !condition_lists.is_empty()
                    {
                        let delay = attribute(element, b"delay", reader.decoder())?
                            .map(|v| parse_timing_value(&v))
                            .transpose()?
                            .unwrap_or(TimingValue::Milliseconds(0));
                        let delay = match delay {
                            TimingValue::Indefinite => Duration::Indefinite,
                            TimingValue::Milliseconds(ms) => Duration::Finite(ms),
                        };
                        let current = TimeCondition {
                            event: attribute(element, b"evt", reader.decoder())?
                                .map(|v| ConditionEvent::parse(&v))
                                .transpose()?,
                            delay,
                            target: None,
                        };
                        let start = condition_lists.last().expect("checked above").1;
                        if empty {
                            let common = &mut frames
                                .last_mut()
                                .ok_or_else(|| invalid("condition has no common time node"))?
                                .node
                                .common;
                            if start {
                                common.start_conditions.push(current)
                            } else {
                                common.end_conditions.push(current)
                            }
                        } else if condition.replace((depth, start, current)).is_some() {
                            return Err(invalid("nested animation conditions are invalid"));
                        }
                    } else if let Some((_, _, current)) = condition.as_mut() {
                        if is_presentationml_name(&namespace, element.name(), b"spTgt") {
                            current.target = Some(ConditionTarget::Shape(parse_shape_id(
                                &attribute(element, b"spid", reader.decoder())?.ok_or_else(
                                    || invalid("condition shape target is missing its ID"),
                                )?,
                            )?));
                        } else if is_presentationml_name(&namespace, element.name(), b"sldTgt") {
                            current.target = Some(ConditionTarget::Slide);
                        } else if is_presentationml_name(&namespace, element.name(), b"tn") {
                            let id = attribute(element, b"val", reader.decoder())?
                                .ok_or_else(|| {
                                    invalid("condition time-node target is missing its ID")
                                })?
                                .parse::<u32>()
                                .map_err(|_| invalid("invalid condition time-node ID"))?;
                            current.target = Some(ConditionTarget::TimeNode(id));
                        } else if is_presentationml_name(&namespace, element.name(), b"rtn") {
                            current.target = Some(ConditionTarget::Runtime(RuntimeTrigger::parse(
                                &attribute(element, b"val", reader.decoder())?.ok_or_else(
                                    || invalid("runtime condition target is missing its value"),
                                )?,
                            )?));
                        }
                    }
                }
                if empty {
                    depth -= 1;
                }
            },
            Event::End(name) => {
                if condition.as_ref().is_some_and(|(d, _, _)| *d == depth)
                    && is_presentationml_name(&namespace, name.name(), b"cond")
                {
                    let (_, start, value) = condition.take().expect("checked above");
                    let common = &mut frames
                        .last_mut()
                        .ok_or_else(|| invalid("condition has no common time node"))?
                        .node
                        .common;
                    if start {
                        common.start_conditions.push(value)
                    } else {
                        common.end_conditions.push(value)
                    }
                }
                if condition_lists.last().is_some_and(|(d, _)| *d == depth) {
                    condition_lists.pop();
                }
                if child_lists.last().is_some_and(|(d, _)| *d == depth) {
                    child_lists.pop();
                }
                if let Some(frame) = frames.pop_if(|frame| frame.depth == depth) {
                    let child = TimingChild::Node(frame.node);
                    if let Some(parent) = frames.last_mut() {
                        if frame.sub_node {
                            parent.node.common.sub_nodes.push(child)
                        } else {
                            parent.node.common.children.push(child)
                        }
                    } else {
                        roots.push(child);
                    }
                }
                if timing_depth == Some(depth)
                    && is_presentationml_name(&namespace, name.name(), b"timing")
                {
                    source_range =
                        Some(timing_start.expect("timing start set")..reader.buffer_position());
                    timing_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced animation timing XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if timing_depth.is_some() || !frames.is_empty() || condition.is_some() {
        return Err(invalid("incomplete recursive animation timing tree"));
    }
    let range = source_range.ok_or_else(|| invalid("animation timing subtree is missing"))?;
    let range = (range.start as usize)..(range.end as usize);
    let source = xml
        .get(range)
        .ok_or_else(|| invalid("animation timing range is invalid"))?
        .to_string()
        .into_boxed_str();
    let mut tree = TimingTree {
        roots,
        opaque_children: Vec::new(),
        source_xml: Some(source),
        source_roots: None,
        source_opaque_children: None,
    };
    tree.source_roots = Some(tree.roots.clone().into_boxed_slice());
    tree.source_opaque_children = Some(tree.opaque_children.clone().into_boxed_slice());
    Ok(tree)
}

pub(super) fn parse_processed_timing(xml: &[u8], require_valid_targets: bool) -> Result<Sequence> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut parser = TimingParser::new(require_valid_targets);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("animation XML offset does not fit usize"))?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("animation XML offset does not fit usize"))?;
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| invalid("animation XML node counter overflow"))?;
        if nodes > MAX_TIMING_NODES {
            return Err(invalid("animation XML node count exceeds safety limit"));
        }
        match event {
            Event::Start(element) => {
                parser.start(
                    &namespace,
                    &element,
                    decoder,
                    depth,
                    false,
                    event_start,
                    event_end,
                )?;
                depth += 1;
                if depth > MAX_TIMING_DEPTH {
                    return Err(invalid("animation XML depth exceeds safety limit"));
                }
            },
            Event::Empty(element) => parser.start(
                &namespace,
                &element,
                decoder,
                depth,
                true,
                event_start,
                event_end,
            )?,
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("animation XML has an unmatched end element"))?;
                parser.end(&namespace, element.name(), depth, event_end)?;
            },
            Event::Text(text) => {
                text_bytes = text_bytes
                    .checked_add(text.as_ref().len())
                    .ok_or_else(|| invalid("animation XML text counter overflow"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("animation XML text exceeds safety limit"));
                }
            },
            Event::CData(text) => {
                text_bytes = text_bytes
                    .checked_add(text.as_ref().len())
                    .ok_or_else(|| invalid("animation XML text counter overflow"))?;
                if text_bytes > MAX_TIMING_TEXT_BYTES {
                    return Err(invalid("animation XML text exceeds safety limit"));
                }
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in animation XML")),
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("incomplete animation XML"));
    }
    parser.finish(xml)
}
