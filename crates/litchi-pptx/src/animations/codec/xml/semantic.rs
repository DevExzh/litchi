use super::super::super::invalid;
use super::super::super::model::*;
use super::super::validation::TimingValue;
use crate::Result;
use std::ops::Range;

pub(super) struct TimeNodeFrame {
    pub(super) depth: usize,
    pub(super) start_delay: Option<TimingValue>,
    pub(super) start_on_click: bool,
    pub(super) start_target: Option<u32>,
    pub(super) interactive_event_filter: Option<Option<EventFilter>>,
}

pub(super) struct PendingAnimation {
    pub(super) depth: usize,
    pub(super) animation: EffectInstance,
    pub(super) target: Option<u32>,
    pub(super) target_element_depth: Option<usize>,
}

pub(super) struct PendingParagraphTemplate {
    pub(super) depth: usize,
    pub(super) build_index: usize,
    pub(super) level: u8,
    pub(super) time_list_depth: Option<usize>,
    pub(super) saw_time_list: bool,
    pub(super) root_depth: Option<usize>,
    pub(super) root_start: Option<usize>,
    pub(super) root_range: Option<Range<usize>>,
}

pub(super) struct PendingGraphicBuild {
    pub(super) depth: usize,
    pub(super) shape_id: u32,
    pub(super) group_id: GroupId,
    pub(super) ui_expand: bool,
    pub(super) sub_build_depth: Option<usize>,
    pub(super) mode: Option<GraphicBuildMode>,
}

pub(super) fn parse_sequence_context(time_nodes: &[TimeNodeFrame]) -> Result<SequenceContext> {
    let Some((interactive_index, event_filter)) = time_nodes
        .iter()
        .enumerate()
        .rev()
        .find_map(|(index, frame)| frame.interactive_event_filter.map(|filter| (index, filter)))
    else {
        return Ok(SequenceContext::Main);
    };
    let trigger_shape_id = time_nodes[interactive_index..]
        .iter()
        .find_map(|frame| frame.start_on_click.then_some(frame.start_target).flatten())
        .ok_or_else(|| {
            invalid("interactive animation sequence lacks a shape-targeted onClick condition")
        })?;
    Ok(SequenceContext::Interactive {
        trigger_shape_id,
        event_filter,
    })
}

pub(super) fn trigger(node_type: Option<&str>, ancestors: &[TimeNodeFrame]) -> Trigger {
    match node_type {
        Some("withEffect" | "withGroup") => Trigger::WithPrevious,
        Some("afterEffect" | "afterGroup") => Trigger::AfterPrevious,
        Some("clickEffect" | "clickPar") if ancestors.iter().any(|node| node.start_on_click) => {
            Trigger::OnClick
        },
        Some("clickEffect" | "clickPar") => {
            match ancestors.iter().find_map(|node| node.start_delay) {
                Some(TimingValue::Milliseconds(_)) => Trigger::WithPrevious,
                _ => Trigger::OnClick,
            }
        },
        _ => match ancestors.iter().find_map(|node| node.start_delay) {
            Some(TimingValue::Indefinite) => Trigger::OnClick,
            Some(TimingValue::Milliseconds(_)) => Trigger::WithPrevious,
            None => Trigger::OnClick,
        },
    }
}
