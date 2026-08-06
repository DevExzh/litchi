use super::super::{
    BuildListEntry, ExtendedTimeNode, SlideAnimationExtension, TimeModifier, TimeNodeBehavior,
    TimeSubEffectBehavior, TimeVisualElement, write_build_list, write_extended_time_node,
};
use super::semantic::EditorLimits;
use crate::package::{Error, Result};
use std::collections::{BTreeSet, HashSet};

pub(super) fn validate_limits(limits: EditorLimits) -> Result<()> {
    if limits.max_persist_records == 0
        || limits.max_record_bytes < 8
        || limits.max_timeline_nodes == 0
        || limits.max_timeline_depth == 0
        || limits.max_build_entries == 0
        || limits.max_shapes == 0
    {
        return invalid("all animation editor limits must be nonzero");
    }
    Ok(())
}

pub(super) fn validate_extension(
    extension: &SlideAnimationExtension,
    shapes: &BTreeSet<u32>,
    limits: EditorLimits,
) -> Result<()> {
    let mut count = 0usize;
    if let Some(root) = &extension.time_node {
        validate_node(root, 1, &mut count, shapes, limits)?;
        let _ = write_extended_time_node(root)?;
    }
    if let Some(builds) = &extension.build_list {
        if builds.builds.len() > limits.max_build_entries {
            return invalid("build list exceeds resource limit");
        }
        let mut ids = HashSet::new();
        for build in &builds.builds {
            let atom = match build {
                BuildListEntry::Paragraph(value) => &value.atom,
                BuildListEntry::Chart(value) => &value.atom,
                BuildListEntry::Diagram(value) => &value.atom,
            };
            if !shapes.contains(&atom.shape_id_ref) {
                return invalid("build atom references a missing shape");
            }
            if !ids.insert(atom.build_id) {
                return invalid("build list contains duplicate build IDs");
            }
        }
        let _ = write_build_list(builds)?;
    }
    Ok(())
}

pub(super) fn validate_node(
    node: &ExtendedTimeNode,
    depth: usize,
    count: &mut usize,
    shapes: &BTreeSet<u32>,
    limits: EditorLimits,
) -> Result<()> {
    *count = count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("timeline node count overflow".into()))?;
    if depth > limits.max_timeline_depth || *count > limits.max_timeline_nodes {
        return invalid("timeline nesting or node count exceeds limits");
    }
    if let Some(target) = &node.visual_target {
        validate_target(target, shapes)?;
    }
    if let Some(behavior) = &node.behavior {
        let target = match behavior {
            TimeNodeBehavior::Animate(v) => &v.behavior.target,
            TimeNodeBehavior::Color(v) => &v.behavior.target,
            TimeNodeBehavior::Effect(v) => &v.behavior.target,
            TimeNodeBehavior::Motion(v) => &v.behavior.target,
            TimeNodeBehavior::Rotation(v) => &v.behavior.target,
            TimeNodeBehavior::Scale(v) => &v.behavior.target,
            TimeNodeBehavior::Set(v) => &v.behavior.target,
            TimeNodeBehavior::Command(v) => &v.behavior.target,
        };
        validate_target(target, shapes)?;
    }
    for modifier in &node.modifiers {
        validate_modifier(modifier)?;
    }
    for effect in &node.sub_effects {
        if let Some(target) = &effect.visual_target {
            validate_target(target, shapes)?;
        }
        if let Some(behavior) = &effect.behavior {
            let target = match behavior {
                TimeSubEffectBehavior::Color(v) => &v.behavior.target,
                TimeSubEffectBehavior::Set(v) => &v.behavior.target,
                TimeSubEffectBehavior::Command(v) => &v.behavior.target,
            };
            validate_target(target, shapes)?;
        }
        for modifier in &effect.modifiers {
            validate_modifier(modifier)?;
        }
    }
    for child in &node.children {
        validate_node(child, depth + 1, count, shapes, limits)?;
    }
    Ok(())
}

pub(super) fn validate_target(target: &TimeVisualElement, shapes: &BTreeSet<u32>) -> Result<()> {
    match target {
        TimeVisualElement::Shape {
            shape_id_ref,
            data1,
            data2,
            ..
        } => {
            if !shapes.contains(shape_id_ref) {
                return invalid("behavior references a missing shape");
            }
            if *data1 < -1 || *data2 < -1 {
                return invalid("text-range target contains an invalid range");
            }
        },
        TimeVisualElement::Chart { shape_id_ref, .. } if !shapes.contains(shape_id_ref) => {
            return invalid("chart behavior references a missing shape");
        },
        _ => {},
    }
    Ok(())
}

pub(super) fn validate_modifier(_value: &TimeModifier) -> Result<()> {
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
