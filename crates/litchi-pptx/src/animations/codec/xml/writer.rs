use super::super::super::model::{
    ConditionTarget, EffectInstance, NextAction, PreviousAction, TimeCondition, TimingChild,
    TimingNodeKind, Trigger,
};
use super::super::validation::direction_subtype;

pub(super) fn write_animation_xml(
    xml: &mut String,
    anim: &EffectInstance,
    tn_id: &mut u32,
    interactive_trigger: Option<u32>,
) {
    xml.push_str(&format!(
        r#"<p:par><p:cTn id="{}" fill="hold"><p:stCondLst>"#,
        *tn_id
    ));
    *tn_id += 1;
    if anim.trigger == Trigger::OnClick {
        if let Some(trigger_shape_id) = interactive_trigger {
            xml.push_str(&format!(
                r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_shape_id}"/></p:tgtEl></p:cond>"#
            ));
        } else {
            xml.push_str(r#"<p:cond delay="indefinite"/>"#);
        }
    } else {
        xml.push_str(r#"<p:cond delay="0"/>"#);
    }
    xml.push_str("</p:stCondLst><p:childTnLst><p:par>");
    xml.push_str(&format!(
        r#"<p:cTn id="{}" fill="hold"><p:stCondLst><p:cond delay="{}"/></p:stCondLst>"#,
        *tn_id, anim.delay
    ));
    *tn_id += 1;

    xml.push_str("<p:childTnLst><p:par>");
    let node_type = match anim.trigger {
        Trigger::OnClick => "clickEffect",
        Trigger::WithPrevious => "withEffect",
        Trigger::AfterPrevious => "afterEffect",
    };
    let preset_subtype = anim
        .direction
        .as_ref()
        .and_then(|direction| direction_subtype(&anim.effect, direction))
        .unwrap_or(0);
    xml.push_str(&format!(
        r#"<p:cTn id="{}" presetID="{}" presetClass="{}" presetSubtype="{}""#,
        *tn_id,
        anim.effect.preset_id(),
        anim.effect.preset_class(),
        preset_subtype
    ));
    if let Some(fill) = anim.fill {
        xml.push_str(&format!(r#" fill="{}""#, fill.as_str()));
    }
    if let Some(restart) = anim.restart {
        xml.push_str(&format!(r#" restart="{}""#, restart.as_str()));
    }
    if anim.auto_reverse {
        xml.push_str(r#" autoRev="1""#);
    }
    if let Some(repeat) = anim.repeat {
        xml.push_str(&format!(r#" repeatCount="{}""#, repeat.write_value()));
    }
    if let Some(speed) = anim.speed {
        xml.push_str(&format!(r#" spd="{}""#, speed.thousandths_percent()));
    }
    if let Some(acceleration) = anim.acceleration {
        xml.push_str(&format!(
            r#" accel="{}""#,
            acceleration.thousandths_percent()
        ));
    }
    if let Some(deceleration) = anim.deceleration {
        xml.push_str(&format!(
            r#" decel="{}""#,
            deceleration.thousandths_percent()
        ));
    }
    if let Some(display) = anim.display {
        xml.push_str(if display {
            r#" display="1""#
        } else {
            r#" display="0""#
        });
    }
    if let Some(repeat_duration) = anim.repeat_duration {
        xml.push_str(&format!(
            r#" repeatDur="{}""#,
            repeat_duration.write_value()
        ));
    }
    if let Some(sync_behavior) = anim.sync_behavior {
        xml.push_str(&format!(r#" syncBehavior="{}""#, sync_behavior.as_str()));
    }
    if let Some(after_effect) = anim.after_effect {
        xml.push_str(if after_effect {
            r#" afterEffect="1""#
        } else {
            r#" afterEffect="0""#
        });
    }
    if let Some(time_filter) = &anim.time_filter {
        xml.push_str(&format!(r#" tmFilter="{}""#, time_filter.write_value()));
    }
    if let Some(group_id) = anim.group_id {
        xml.push_str(&format!(r#" grpId="{}""#, group_id.value()));
    }
    xml.push_str(&format!(
        r#" nodeType="{}" dur="{}">"#,
        node_type,
        anim.duration.write_value()
    ));
    *tn_id += 1;

    xml.push_str("<p:childTnLst>");
    xml.push_str(&format!(r#"<p:set><p:cBhvr><p:cTn id="{}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>"#, *tn_id));
    *tn_id += 1;
    xml.push_str(&format!(
        r#"<p:tgtEl><p:spTgt spid="{}"/></p:tgtEl>"#,
        anim.shape_id
    ));
    xml.push_str(r#"<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>"#);
    xml.push_str("</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>");
}

pub(super) fn write_timing_child(xml: &mut String, child: &TimingChild) {
    let TimingChild::Node(node) = child else {
        if let TimingChild::Opaque(raw) = child {
            xml.push_str(raw);
        }
        return;
    };
    match node.kind {
        TimingNodeKind::Parallel => xml.push_str("<p:par>"),
        TimingNodeKind::Exclusive => xml.push_str("<p:excl>"),
        TimingNodeKind::Sequence {
            concurrent,
            next_action,
            previous_action,
        } => {
            xml.push_str("<p:seq");
            if concurrent {
                xml.push_str(" concurrent=\"1\"");
            }
            if next_action == NextAction::Seek {
                xml.push_str(" nextAc=\"seek\"");
            }
            if previous_action == PreviousAction::SkipTimed {
                xml.push_str(" prevAc=\"skipTimed\"");
            }
            xml.push('>');
        },
    }
    let common = &node.common;
    xml.push_str("<p:cTn");
    if let Some(id) = common.id {
        xml.push_str(&format!(" id=\"{id}\""));
    }
    if let Some(duration) = common.duration {
        xml.push_str(&format!(" dur=\"{}\"", duration.write_value()));
    }
    if let Some(node_type) = common.node_type {
        xml.push_str(&format!(" nodeType=\"{}\"", node_type.as_str()));
    }
    if let Some(preset) = &common.preset {
        xml.push_str(&format!(
            " presetID=\"{}\" presetClass=\"{}\"",
            preset.preset_id,
            preset.class.as_str()
        ));
        if let Some(subtype) = preset.subtype {
            xml.push_str(&format!(" presetSubtype=\"{subtype}\""));
        }
    }
    xml.push('>');
    write_condition_list(xml, "stCondLst", &common.start_conditions);
    write_condition_list(xml, "endCondLst", &common.end_conditions);
    if !common.children.is_empty() {
        xml.push_str("<p:childTnLst>");
        for child in &common.children {
            write_timing_child(xml, child);
        }
        xml.push_str("</p:childTnLst>");
    }
    if !common.sub_nodes.is_empty() {
        xml.push_str("<p:subTnLst>");
        for child in &common.sub_nodes {
            write_timing_child(xml, child);
        }
        xml.push_str("</p:subTnLst>");
    }
    for raw in &common.opaque_children {
        xml.push_str(raw);
    }
    xml.push_str("</p:cTn>");
    for raw in &node.opaque_children {
        xml.push_str(raw);
    }
    match node.kind {
        TimingNodeKind::Parallel => xml.push_str("</p:par>"),
        TimingNodeKind::Exclusive => xml.push_str("</p:excl>"),
        TimingNodeKind::Sequence { .. } => xml.push_str("</p:seq>"),
    }
}

fn write_condition_list(xml: &mut String, name: &str, conditions: &[TimeCondition]) {
    if conditions.is_empty() {
        return;
    }
    xml.push_str(&format!("<p:{name}>"));
    for condition in conditions {
        xml.push_str("<p:cond");
        if let Some(event) = condition.event {
            xml.push_str(&format!(" evt=\"{}\"", event.as_str()));
        }
        xml.push_str(&format!(" delay=\"{}\"", condition.delay.write_value()));
        match condition.target {
            None => xml.push_str("/>"),
            Some(ConditionTarget::Shape(id)) => xml.push_str(&format!(
                "><p:tgtEl><p:spTgt spid=\"{id}\"/></p:tgtEl></p:cond>"
            )),
            Some(ConditionTarget::Slide) => {
                xml.push_str("><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>");
            },
            Some(ConditionTarget::TimeNode(id)) => {
                xml.push_str(&format!("><p:tn val=\"{id}\"/></p:cond>"));
            },
            Some(ConditionTarget::Runtime(value)) => {
                xml.push_str(&format!("><p:rtn val=\"{}\"/></p:cond>", value.as_str()));
            },
        }
    }
    xml.push_str(&format!("</p:{name}>"));
}
