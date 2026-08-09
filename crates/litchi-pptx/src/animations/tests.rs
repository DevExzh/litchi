#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::*;
use super::model::*;
use std::collections::HashSet;
#[cfg(test)]
mod recursive_timing_tests {
    use super::*;

    const NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    #[test]
    fn preserves_nested_presets_and_ordered_conditions() {
        let xml = format!(
            r#"<p:timing xmlns:p="{NS}"><p:tnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="1" nodeType="interactiveSeq" presetID="10" presetClass="entr"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cond><p:cond evt="onEnd" delay="25"><p:tn val="9"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="2" presetID="11" presetClass="emph"/></p:par></p:childTnLst></p:cTn><p:extLst><p:ext uri="opaque"><x:data xmlns:x="urn:test"/></p:ext></p:extLst></p:seq></p:tnLst></p:timing>"#
        );
        let tree = TimingTree::parse(&xml).expect("nested timing parses");
        assert_eq!(tree.to_xml(), xml);
        let TimingChild::Node(root) = &tree.roots[0] else {
            panic!("typed root")
        };
        assert!(matches!(
            root.kind,
            TimingNodeKind::Sequence {
                concurrent: true,
                next_action: NextAction::Seek,
                ..
            }
        ));
        assert_eq!(root.common.start_conditions.len(), 2);
        assert!(matches!(
            root.common.start_conditions[0].target,
            Some(ConditionTarget::Shape(7))
        ));
        assert!(matches!(
            root.common.start_conditions[1].target,
            Some(ConditionTarget::TimeNode(9))
        ));
        assert!(root.common.preset.is_some());
        let TimingChild::Node(child) = &root.common.children[0] else {
            panic!("typed child")
        };
        assert_eq!(
            child.common.preset.as_ref().map(|preset| preset.preset_id),
            Some(11)
        );
    }

    #[test]
    fn rejects_malformed_common_time_node_id() {
        let xml = format!(
            r#"<p:timing xmlns:p="{NS}"><p:tnLst><p:par><p:cTn id="not-a-number"/></p:par></p:tnLst></p:timing>"#
        );
        assert!(TimingTree::parse(&xml).is_err());
    }

    #[test]
    fn rejects_excessive_recursive_depth() {
        let mut xml = format!(r#"<p:timing xmlns:p="{NS}"><p:tnLst>"#);
        for id in 1..=MAX_TIMING_DEPTH + 1 {
            xml.push_str(&format!("<p:par><p:cTn id=\"{id}\"><p:childTnLst>"));
        }
        for _ in 1..=MAX_TIMING_DEPTH + 1 {
            xml.push_str("</p:childTnLst></p:cTn></p:par>");
        }
        xml.push_str("</p:tnLst></p:timing>");
        assert!(TimingTree::parse(&xml).is_err());
    }

    #[test]
    fn rejects_excessive_node_count() {
        let mut xml = format!(r#"<p:timing xmlns:p="{NS}"><p:tnLst>"#);
        for id in 1..=MAX_TIMING_NODES + 1 {
            xml.push_str(&format!("<p:par><p:cTn id=\"{id}\"/></p:par>"));
        }
        xml.push_str("</p:tnLst></p:timing>");
        assert!(TimingTree::parse(&xml).is_err());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

    fn slide(timing: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="{P}"><p:cSld><p:spTree>
                <p:sp><p:nvSpPr><p:cNvPr id="3" name="A"/></p:nvSpPr></p:sp>
                <p:pic><p:nvPicPr><p:cNvPr id="4" name="B"/></p:nvPicPr></p:pic>
                <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="C"/></p:nvGraphicFramePr></p:graphicFrame>
            </p:spTree></p:cSld>{timing}</p:sld>"#
        )
    }

    fn effect(
        shape: &str,
        preset: u32,
        class: &str,
        node_type: &str,
        trigger_delay: &str,
        delay: u32,
        duration: u32,
    ) -> String {
        format!(
            r#"<p:par><p:cTn><p:stCondLst><p:cond delay="{trigger_delay}"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
            <p:childTnLst><p:par><p:cTn presetID="{preset}" presetClass="{class}" presetSubtype="0" nodeType="{node_type}" dur="{duration}">
            <p:childTnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="{shape}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst>
            </p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"#
        )
    }

    fn interactive_effect(trigger_shape: &str, event_filter: &str) -> String {
        let triggered = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replacen(
                r#"<p:cond delay="indefinite"/>"#,
                &format!(
                    r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{trigger_shape}"/></p:tgtEl></p:cond>"#
                ),
                1,
            );
        format!(
            r#"<p:seq><p:cTn nodeType="interactiveSeq" evtFilter="{event_filter}"><p:childTnLst>{triggered}</p:childTnLst></p:cTn></p:seq>"#
        )
    }

    fn grouped_timing(shape_id: &str, group_id: &str) -> String {
        let grouped = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldP spid="{shape_id}" grpId="{group_id}"/></p:bldLst></p:timing>"#
        )
    }

    fn diagram_timing(shape_id: &str, group_id: &str, attributes: &str) -> String {
        let grouped = effect("5", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldDgm spid="{shape_id}" grpId="{group_id}"{attributes}/></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_ole(timing: &str) -> String {
        slide(timing).replace(
            r"</p:nvGraphicFramePr></p:graphicFrame>",
            r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj/></a:graphicData></a:graphic></p:graphicFrame>"#,
        )
    }

    fn ole_chart_timing(shape_id: &str, group_id: &str, attributes: &str) -> String {
        let grouped = effect(shape_id, 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldOleChart spid="{shape_id}" grpId="{group_id}"{attributes}/></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_ole_chart(timing: &str) -> String {
        slide_with_ole(timing).replace("<p:oleObj/>", r#"<p:oleObj progId="MSGraph.Chart.8"/>"#)
    }

    fn graphic_timing(shape_id: &str, group_id: &str, attributes: &str, content: &str) -> String {
        let grouped = effect(shape_id, 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            &format!(r#" grpId="{group_id}" nodeType="clickEffect""#),
        );
        format!(
            r#"<p:timing><p:tnLst>{grouped}</p:tnLst><p:bldLst><p:bldGraphic spid="{shape_id}" grpId="{group_id}"{attributes}>{content}</p:bldGraphic></p:bldLst></p:timing>"#
        )
    }

    fn slide_with_graphic_hosts(timing: &str) -> String {
        slide(timing)
            .replace(
                r#"<p:sld xmlns:p=""#,
                r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p=""#,
            )
            .replace(
                r"</p:nvGraphicFramePr></p:graphicFrame>",
                r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="SmartArt"/></p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:graphicData></a:graphic></p:graphicFrame>"#,
            )
    }

    #[test]
    fn test_animation_effect_preset() {
        assert_eq!(Effect::Fade.preset_class(), "entr");
        assert_eq!(Effect::Fade.preset_id(), 10);
        assert_eq!(Effect::from_preset("fade"), Effect::Fade);
    }

    #[test]
    fn test_animation_sequence() {
        let mut seq = Sequence::new();
        seq.add(EffectInstance::new(1, Effect::Fade).with_duration_ms(1000));
        seq.add(EffectInstance::new(2, Effect::FlyIn).with_trigger(Trigger::AfterPrevious));

        assert_eq!(seq.len(), 2);
        assert!(!seq.to_xml().is_empty());
    }

    #[test]
    fn parses_typed_timing_metadata_from_slide() {
        let timing = format!(
            "<p:timing><p:tnLst>{}{}{}</p:tnLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 125, 750),
            effect("4", 42, "entr", "withEffect", "0", 20, 600),
            effect("5", 8, "emph", "afterEffect", "0", 40, 900),
        );
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.len(), 3);
        assert_eq!(sequence.animations[0].shape_id, 3);
        assert_eq!(sequence.animations[0].effect, Effect::Fade);
        assert_eq!(sequence.animations[0].trigger, Trigger::OnClick);
        assert_eq!(sequence.animations[0].duration, Duration::Finite(750));
        assert_eq!(sequence.animations[0].delay, 125);
        assert_eq!(sequence.animations[1].effect, Effect::FloatIn);
        assert_eq!(sequence.animations[1].trigger, Trigger::WithPrevious);
        assert_eq!(sequence.animations[2].effect, Effect::Spin);
        assert_eq!(sequence.animations[2].trigger, Trigger::AfterPrevious);
        assert_eq!(sequence.animations[2].order, 3);
    }

    #[test]
    fn rejects_malformed_missing_duplicate_spoofed_and_off_slide_targets() {
        let cases = [
            effect("0", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("nope", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("99", 10, "entr", "clickEffect", "indefinite", 0, 500),
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("<p:spTgt spid=\"3\"/>", ""),
        ];
        for effect in cases {
            let timing = format!("<p:timing><p:tnLst>{effect}</p:tnLst></p:timing>");
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }

        let duplicate = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            "<p:spTgt spid=\"3\"/>",
            "<p:spTgt spid=\"3\"/><p:spTgt spid=\"4\"/>",
        );
        let timing = format!("<p:timing><p:tnLst>{duplicate}</p:tnLst></p:timing>");
        assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());

        let spoofed = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            "<p:spTgt spid=\"3\"/>",
            "<x:spTgt xmlns:x=\"urn:foreign\" spid=\"3\"/>",
        );
        let timing = format!("<p:timing><p:tnLst>{spoofed}</p:tnLst></p:timing>");
        assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
    }

    #[test]
    fn rejects_excessive_timing_depth() {
        let nested = format!(
            "<p:timing>{}{}</p:timing>",
            "<p:par>".repeat(MAX_TIMING_DEPTH + 1),
            "</p:par>".repeat(MAX_TIMING_DEPTH + 1)
        );
        assert!(Sequence::parse_slide_xml(slide(&nested).as_bytes()).is_err());
    }

    #[test]
    fn preserves_indefinite_duration() {
        let timing = format!(
            "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("dur=\"500\"", "dur=\"indefinite\"")
        );
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.animations[0].duration, Duration::Indefinite);
    }

    #[test]
    fn preserves_unsupported_timing_subtrees_until_typed_data_changes() {
        let timing = format!(
            "<p:timing><p:tnLst>{}</p:tnLst><p:extLst><p:ext uri=\"urn:test\"><x:opaque xmlns:x=\"urn:opaque\" value=\"kept\"><![CDATA[raw]]></x:opaque></p:ext></p:extLst></p:timing>",
            effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
        );
        let mut sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.preserved_timing_xml(), Some(timing.as_str()));
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains("dur=\"750\""));
        assert!(!canonical.contains("x:opaque"));
    }

    #[test]
    fn parses_directional_preset_subtypes() {
        let timing = format!(
            "<p:timing><p:tnLst>{}{}{}{}</p:tnLst></p:timing>",
            effect("3", 2, "entr", "clickEffect", "indefinite", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"3\""),
            effect("4", 22, "entr", "withEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"12\""),
            effect("5", 16, "entr", "afterEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"26\""),
            effect("3", 23, "entr", "withEffect", "0", 0, 500)
                .replace("presetSubtype=\"0\"", "presetSubtype=\"288\"")
        );
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.animations[0].direction, Some(Direction::UpRight));
        assert_eq!(sequence.animations[1].direction, Some(Direction::DownLeft));
        assert_eq!(
            sequence.animations[2].direction,
            Some(Direction::HorizontalIn)
        );
        assert_eq!(
            sequence.animations[3].direction,
            Some(Direction::OutSlightly)
        );
    }

    #[test]
    fn parses_common_time_node_playback_controls() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " fill=\"freeze\" restart=\"whenNotActive\" autoRev=\"1\" repeatCount=\"3500\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(animation.fill, Some(Fill::Freeze));
        assert_eq!(animation.restart, Some(Restart::WhenNotActive));
        assert!(animation.auto_reverse);
        assert_eq!(animation.repeat, Some(Repeat::Finite(3500)));
    }

    #[test]
    fn rejects_invalid_playback_control_values() {
        for attribute in [
            "fill=\"sticky\"",
            "restart=\"sometimes\"",
            "autoRev=\"yes\"",
            "repeatCount=\"2147483626\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_common_time_node_progression_controls() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " spd=\"-50000\" accel=\"25000\" decel=\"10000\" display=\"0\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(
            animation.speed.map(Speed::thousandths_percent),
            Some(-50000)
        );
        assert_eq!(
            animation
                .acceleration
                .map(MotionFraction::thousandths_percent),
            Some(25000)
        );
        assert_eq!(
            animation
                .deceleration
                .map(MotionFraction::thousandths_percent),
            Some(10000)
        );
        assert_eq!(animation.display, Some(false));
    }

    #[test]
    fn rejects_invalid_progression_control_values() {
        for attribute in [
            "spd=\"0\"",
            "spd=\"2147483648\"",
            "accel=\"100001\"",
            "accel=\"-1\"",
            "decel=\"100001\"",
            "display=\"visible\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_repeat_duration_sync_and_after_effect() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500)
            .replace(
                " nodeType=\"clickEffect\"",
                " repeatDur=\"indefinite\" syncBehavior=\"locked\" afterEffect=\"1\" nodeType=\"clickEffect\"",
            );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let animation = &sequence.animations[0];
        assert_eq!(animation.repeat_duration, Some(Duration::Indefinite));
        assert_eq!(animation.sync_behavior, Some(SyncBehavior::Locked));
        assert_eq!(animation.after_effect, Some(true));
    }

    #[test]
    fn rejects_invalid_repeat_duration_sync_and_after_effect() {
        for attribute in [
            "repeatDur=\"2147483626\"",
            "repeatDur=\"forever\"",
            "syncBehavior=\"slippery\"",
            "afterEffect=\"yes\"",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" {attribute} nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_exact_normalized_time_filter_pairs() {
        let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            " nodeType=\"clickEffect\"",
            " tmFilter=\"0.0,0.0; 0.25,0.07; 0.50,0.2; 1.0,1.0\" nodeType=\"clickEffect\"",
        );
        let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
        let sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        let points = sequence.animations[0]
            .time_filter
            .as_ref()
            .unwrap()
            .points();
        assert_eq!(points.len(), 4);
        assert_eq!(
            (
                points[1].local_time.numerator(),
                points[1].local_time.scale()
            ),
            (25, 100)
        );
        assert_eq!(
            (
                points[1].warped_time.numerator(),
                points[1].warped_time.scale()
            ),
            (7, 100)
        );
    }

    #[test]
    fn rejects_malformed_out_of_range_or_unordered_time_filters() {
        for filter in [
            "",
            "0",
            "0,0,0",
            "-0.1,0",
            "0,1.0001",
            "0.5,0;0.5,1",
            "0.75,0;0.25,1",
            "0.1234567890123456789,0",
        ] {
            let configured = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
                " nodeType=\"clickEffect\"",
                &format!(" tmFilter=\"{filter}\" nodeType=\"clickEffect\""),
            );
            let timing = format!("<p:timing><p:tnLst>{configured}</p:tnLst></p:timing>");
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn parses_and_canonicalizes_contextual_cancel_bubble_filter() {
        let interactive = interactive_effect("4", "cancelBubble");
        let timing = format!("<p:timing><p:tnLst>{interactive}</p:tnLst></p:timing>");
        let mut sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.animations[0].sequence_context,
            SequenceContext::Interactive {
                trigger_shape_id: 4,
                event_filter: Some(EventFilter::CancelBubble),
            }
        );
        assert_eq!(sequence.preserved_timing_xml(), Some(timing.as_str()));
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"nodeType="interactiveSeq" evtFilter="cancelBubble""#));
        assert!(
            canonical.contains(r#"<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="4"/>"#)
        );
        let reparsed = Sequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_event_filter_outside_proven_triggered_sequence_context() {
        let on_effect = effect("3", 10, "entr", "clickEffect", "indefinite", 0, 500).replace(
            r#" nodeType="clickEffect""#,
            r#" evtFilter="cancelBubble" nodeType="clickEffect""#,
        );
        let cases = [
            format!("<p:timing><p:tnLst>{on_effect}</p:tnLst></p:timing>"),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "bubble")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble")
                    .replace("evt=\"onClick\"", "evt=\"onNext\"")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble")
                    .replace(r#"<p:tgtEl><p:spTgt spid="4"/></p:tgtEl>"#, "")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("99", "cancelBubble")
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble").replace(
                    r#"<p:spTgt spid="4"/>"#,
                    r#"<p:spTgt spid="4"/><p:spTgt spid="5"/>"#,
                )
            ),
            format!(
                "<p:timing><p:tnLst>{}</p:tnLst></p:timing>",
                interactive_effect("4", "cancelBubble").replace(
                    r#"<p:spTgt spid="4"/>"#,
                    r#"<x:spTgt xmlns:x="urn:foreign" spid="4"/>"#,
                )
            ),
        ];
        for timing in cases {
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn writes_typed_interactive_context_and_validates_trigger_shape() {
        let mut sequence = Sequence::new();
        sequence.add(
            EffectInstance::new(3, Effect::Fade)
                .with_interactive_trigger(4)
                .with_trigger(Trigger::OnClick),
        );
        let targets = HashSet::from([3, 4]);
        let xml = sequence.to_xml_for_slide(&targets).unwrap();
        assert!(xml.contains(r#"nodeType="interactiveSeq" evtFilter="cancelBubble""#));
        let parsed = Sequence::parse_slide_xml(slide(&xml).as_bytes()).unwrap();
        assert_eq!(parsed, sequence);

        sequence.animations[0].sequence_context = SequenceContext::Interactive {
            trigger_shape_id: 99,
            event_filter: Some(EventFilter::CancelBubble),
        };
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_paragraph_build_group_references() {
        let timing = grouped_timing("3", "42");
        let mut sequence = Sequence::parse_slide_xml(slide(&timing).as_bytes()).unwrap();
        assert_eq!(sequence.animations[0].group_id, Some(GroupId::new(42)));
        assert_eq!(
            sequence.paragraph_builds,
            vec![ParagraphBuild::new(3, GroupId::new(42))]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"grpId="42" nodeType="clickEffect""#));
        assert!(canonical.contains(r#"<p:bldLst><p:bldP spid="3" grpId="42"/></p:bldLst>"#));
        let reparsed = Sequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_malformed_dangling_duplicate_or_off_slide_build_groups() {
        let cases = [
            grouped_timing("3", "-1"),
            grouped_timing("3", "4294967296"),
            grouped_timing("99", "42"),
            grouped_timing("3", "42").replace(
                r#" grpId="42" nodeType="clickEffect""#,
                r#" nodeType="clickEffect""#,
            ),
            grouped_timing("3", "42").replace(r#"<p:bldP spid="3" grpId="42"/>"#, ""),
            grouped_timing("3", "42")
                .replace(r#"<p:bldP spid="3" grpId="42"/>"#, r#"<p:bldP spid="3"/>"#),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<p:bldP grpId="42"/>"#,
            ),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<p:bldP spid="3" grpId="42"/><p:bldP spid="3" grpId="42"/>"#,
            ),
            grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                r#"<x:bldP xmlns:x="urn:foreign" spid="3" grpId="42"/>"#,
            ),
        ];
        for timing in cases {
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_build_group_membership_and_targets() {
        let targets = HashSet::from([3]);
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(3, Effect::Fade).with_group_id(7));
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        sequence.add_paragraph_build(ParagraphBuild::new(3, GroupId::new(7)));
        assert!(sequence.to_xml_for_slide(&targets).is_ok());

        sequence.paragraph_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_bldp_schema_defaults_and_powerpoint_auto_advance_semantics() {
        let sequence =
            Sequence::parse_slide_xml(slide(&grouped_timing("3", "42")).as_bytes()).unwrap();
        let build = &sequence.paragraph_builds[0];
        assert!(!build.ui_expand);
        assert_eq!(build.build_type, ParagraphBuildType::Whole);
        assert_eq!(build.build_level, 1);
        assert!(!build.animate_background);
        assert!(build.auto_update_animate_background);
        assert!(!build.reverse);
        assert_eq!(build.auto_advance, Duration::Indefinite);
        assert_eq!(build.powerpoint_auto_advance_milliseconds(), 0);
    }

    #[test]
    fn round_trips_complete_typed_bldp_optional_attribute_grammar() {
        let configured = grouped_timing("3", "42").replace(
            r#"<p:bldP spid="3" grpId="42"/>"#,
            r#"<p:bldP spid="3" grpId="42" uiExpand="true" build="p" bldLvl="3" animBg="1" autoUpdateAnimBg="false" rev="true" advAuto="4294967295"/>"#,
        );
        let mut sequence = Sequence::parse_slide_xml(slide(&configured).as_bytes()).unwrap();
        let build = &sequence.paragraph_builds[0];
        assert!(build.ui_expand);
        assert_eq!(build.build_type, ParagraphBuildType::Paragraph);
        assert_eq!(build.build_level, 3);
        assert!(build.animate_background);
        assert!(!build.auto_update_animate_background);
        assert!(build.reverse);
        assert_eq!(build.auto_advance, Duration::Finite(u32::MAX));

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldP spid="3" grpId="42" uiExpand="1" build="p" bldLvl="3" animBg="1" autoUpdateAnimBg="0" rev="1" advAuto="4294967295"/>"#
        ));
        let reparsed = Sequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_invalid_bldp_optional_attributes_and_cross_field_combinations() {
        for attributes in [
            r#"build="paragraph""#,
            r#"uiExpand="yes""#,
            r#"build="whole" bldLvl="2""#,
            r#"build="whole" rev="1""#,
            r#"bldLvl="-1" build="p""#,
            r#"bldLvl="4294967296" build="p""#,
            r#"animBg="sometimes""#,
            r#"autoUpdateAnimBg="sometimes""#,
            r#"rev="sometimes" build="p""#,
            r#"advAuto="-1""#,
            r#"advAuto="4294967296""#,
            r#"advAuto="forever""#,
        ] {
            let timing = grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                &format!(r#"<p:bldP spid="3" grpId="42" {attributes}/>"#),
            );
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_paragraph_build_cross_field_constraints() {
        let targets = HashSet::from([3]);
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(3, Effect::Fade).with_group_id(7));
        sequence.add_paragraph_build(ParagraphBuild::new(3, GroupId::new(7)).with_build_level(2));
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        sequence.paragraph_builds[0] = sequence.paragraph_builds[0]
            .clone()
            .with_build_type(ParagraphBuildType::Paragraph)
            .with_reverse(true)
            .with_auto_advance(250u32);
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
    }

    #[test]
    fn parses_preserves_and_canonicalizes_complete_paragraph_template_lists() {
        let first = r#"<p:par><p:cTn id="80" dur="500"/></p:par>"#;
        let second = r#"<p:par><p:cTn id="81" dur="indefinite"><p:childTnLst/></p:cTn></p:par>"#;
        let configured = grouped_timing("3", "42").replace(
            r#"<p:bldP spid="3" grpId="42"/>"#,
            &format!(r#"<p:bldP spid="3" grpId="42" build="p"><p:tmplLst><p:tmpl><p:tnLst>{first}</p:tnLst></p:tmpl><p:tmpl lvl="2"><p:tnLst>{second}</p:tnLst></p:tmpl></p:tmplLst></p:bldP>"#),
        );
        let mut sequence = Sequence::parse_slide_xml(slide(&configured).as_bytes()).unwrap();
        let templates = &sequence.paragraph_builds[0].templates;
        assert_eq!(templates.len(), 2);
        assert_eq!(templates[0].level, 0);
        assert_eq!(templates[0].time_node.as_xml(), first);
        assert_eq!(templates[1].level, 2);
        assert_eq!(templates[1].time_node.as_xml(), second);
        assert_eq!(sequence.to_xml(), configured);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(&format!(r#"<p:tmplLst><p:tmpl><p:tnLst>{first}</p:tnLst></p:tmpl><p:tmpl lvl="2"><p:tnLst>{second}</p:tnLst></p:tmpl></p:tmplLst>"#)));
        let reparsed = Sequence::parse_slide_xml(slide(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn rejects_invalid_paragraph_template_cardinality_levels_order_and_namespaces() {
        let par = r#"<p:par><p:cTn id="80"/></p:par>"#;
        let template =
            |level: &str, body: &str| format!(r"<p:tmpl{level}><p:tnLst>{body}</p:tnLst></p:tmpl>");
        let lists = [
            template("", ""),
            r"<p:tmpl/>".to_string(),
            format!(r"<p:tmpl><p:tnLst>{par}{par}</p:tnLst></p:tmpl>"),
            r"<p:tmpl><p:tnLst><p:par><p:cTn/><p:cTn/></p:par></p:tnLst></p:tmpl>".to_string(),
            r"<p:tmpl><p:tnLst><p:seq><p:cTn/></p:seq></p:tnLst></p:tmpl>".to_string(),
            r#"<p:tmpl><p:tnLst><x:par xmlns:x="urn:foreign"><x:cTn/></x:par></p:tnLst></p:tmpl>"#
                .to_string(),
            format!(r"<p:tmpl><p:tnLst>{par}</p:tnLst><p:tnLst>{par}</p:tnLst></p:tmpl>"),
            format!(r#"<p:tmpl lvl="10"><p:tnLst>{par}</p:tnLst></p:tmpl>"#),
            format!(r#"<p:tmpl lvl="nope"><p:tnLst>{par}</p:tnLst></p:tmpl>"#),
            format!(r"{}{}", template("", par), template("", par)),
            (0..10)
                .map(|level| template(&format!(r#" lvl="{level}""#), par))
                .collect::<String>(),
        ];
        for list in lists {
            let timing = grouped_timing("3", "42").replace(
                r#"<p:bldP spid="3" grpId="42"/>"#,
                &format!(r#"<p:bldP spid="3" grpId="42" build="p"><p:tmplLst>{list}</p:tmplLst></p:bldP>"#),
            );
            assert!(Sequence::parse_slide_xml(slide(&timing).as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_paragraph_template_fragments_and_constraints() {
        for xml in [
            "",
            r"<p:seq><p:cTn/></p:seq>",
            r"<p:par/>",
            r"<p:par><p:cTn/><p:cTn/></p:par>",
            r#"<x:par xmlns:x="urn:foreign"><x:cTn/></x:par>"#,
            r"<!DOCTYPE x><p:par><p:cTn/></p:par>",
        ] {
            assert!(TemplateTimeNode::parse(xml).is_err());
        }

        let node = TemplateTimeNode::parse(r#"<p:par><p:cTn id="90"/></p:par>"#).unwrap();
        assert!(ParagraphTemplate::new(10, node.clone()).is_err());
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(3, Effect::Fade).with_group_id(7));
        let duplicate = ParagraphTemplate::new(1, node.clone()).unwrap();
        sequence.add_paragraph_build(
            ParagraphBuild::new(3, GroupId::new(7))
                .with_build_type(ParagraphBuildType::Paragraph)
                .with_template(duplicate.clone())
                .with_template(duplicate),
        );
        assert!(sequence.to_xml_for_slide(&HashSet::from([3])).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_diagram_build_grammar() {
        let timing = diagram_timing("5", "77", r#" uiExpand="true" bld="ccwOut""#);
        let mut sequence = Sequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.diagram_builds,
            vec![DiagramBuild {
                shape_id: 5,
                group_id: GroupId::new(77),
                ui_expand: true,
                build_type: DiagramBuildType::CounterClockwiseOut,
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(r#"<p:bldDgm spid="5" grpId="77" uiExpand="1" bld="ccwOut"/>"#));
        let reparsed = Sequence::parse_slide_xml(slide_with_ole(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_every_diagram_build_enum_and_schema_defaults() {
        for token in [
            "whole",
            "depthByNode",
            "depthByBranch",
            "breadthByNode",
            "breadthByLvl",
            "cw",
            "cwIn",
            "cwOut",
            "ccw",
            "ccwIn",
            "ccwOut",
            "inByRing",
            "outByRing",
            "up",
            "down",
            "allAtOnce",
            "cust",
        ] {
            let timing = diagram_timing("5", "77", &format!(r#" bld="{token}""#));
            assert!(Sequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).is_ok());
        }
        let timing = diagram_timing("5", "77", "");
        let sequence = Sequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.diagram_builds[0].build_type,
            DiagramBuildType::Whole
        );
        assert!(!sequence.diagram_builds[0].ui_expand);
    }

    #[test]
    fn rejects_invalid_diagram_builds_and_non_ole_or_spoofed_targets() {
        let cases = [
            slide_with_ole(&diagram_timing("5", "77", r#" bld="sideways""#)),
            slide_with_ole(&diagram_timing("5", "77", r#" uiExpand="yes""#)),
            slide_with_ole(&diagram_timing("99", "77", "")),
            slide(&diagram_timing("5", "77", "")),
            slide_with_ole(&diagram_timing("3", "77", "")),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5" grpId="77"/><p:bldDgm spid="5" grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<p:bldDgm spid="5" grpId="77"><p:extLst/></p:bldDgm>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", "").replace(
                r#"<p:bldDgm spid="5" grpId="77"/>"#,
                r#"<x:bldDgm xmlns:x="urn:foreign" spid="5" grpId="77"/>"#,
            )),
            slide_with_ole(&diagram_timing("5", "77", ""))
                .replace(r"<p:oleObj/>", r#"<x:oleObj xmlns:x="urn:foreign"/>"#),
        ];
        for xml in cases {
            assert!(Sequence::parse_slide_xml(xml.as_bytes()).is_err());
        }
    }

    #[test]
    fn validates_programmatic_diagram_build_groups_targets_and_duplicates() {
        let targets = HashSet::from([5]);
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(5, Effect::Fade).with_group_id(77));
        sequence.add_diagram_build(
            DiagramBuild::new(5, GroupId::new(77))
                .with_ui_expand(true)
                .with_build_type(DiagramBuildType::BreadthByLevel),
        );
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.diagram_builds.push(sequence.diagram_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.diagram_builds.pop();
        sequence.diagram_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_graphic_chart_build() {
        let timing = graphic_timing(
            "5",
            "88",
            r#" uiExpand="true""#,
            r#"<p:bldSub><a:bldChart bld="seriesEl" animBg="false"/></p:bldSub>"#,
        );
        let mut sequence =
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.graphic_builds,
            vec![GraphicBuild {
                shape_id: 5,
                group_id: GroupId::new(88),
                ui_expand: true,
                mode: GraphicBuildMode::Chart {
                    build_type: GraphicChartBuildType::SeriesElement,
                    animate_background: false,
                },
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldGraphic spid="5" grpId="88" uiExpand="1"><p:bldSub><a:bldChart bld="seriesEl" animBg="0"/></p:bldSub></p:bldGraphic>"#
        ));
        let reparsed =
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_all_graphic_build_modes_tokens_and_schema_defaults() {
        let as_one = graphic_timing("5", "88", "", "<p:bldAsOne/>");
        let sequence =
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&as_one).as_bytes()).unwrap();
        assert_eq!(sequence.graphic_builds[0].mode, GraphicBuildMode::AsOne);

        for token in ["allAtOnce", "one", "lvlOne", "lvlAtOnce"] {
            let timing = graphic_timing(
                "6",
                "88",
                "",
                &format!(r#"<p:bldSub><a:bldDgm bld="{token}"/></p:bldSub>"#),
            );
            assert!(
                Sequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes()).is_ok()
            );
        }
        for token in ["allAtOnce", "series", "category", "seriesEl", "categoryEl"] {
            let timing = graphic_timing(
                "5",
                "88",
                "",
                &format!(r#"<p:bldSub><a:bldChart bld="{token}"/></p:bldSub>"#),
            );
            assert!(
                Sequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes()).is_ok()
            );
        }

        let diagram = graphic_timing("6", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>");
        let sequence =
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&diagram).as_bytes()).unwrap();
        assert_eq!(
            sequence.graphic_builds[0].mode,
            GraphicBuildMode::Diagram {
                build_type: GraphicDiagramBuildType::AllAtOnce,
                reverse: false,
            }
        );
        let chart = graphic_timing("5", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>");
        let sequence =
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&chart).as_bytes()).unwrap();
        assert_eq!(
            sequence.graphic_builds[0].mode,
            GraphicBuildMode::Chart {
                build_type: GraphicChartBuildType::AllAtOnce,
                animate_background: true,
            }
        );
    }

    #[test]
    fn rejects_hostile_graphic_build_grammar_namespaces_and_host_mismatches() {
        let valid_chart = graphic_timing("5", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>");
        let cases = [
            graphic_timing("5", "88", "", ""),
            graphic_timing("5", "88", "", "<p:bldSub/>"),
            graphic_timing("5", "88", "", "<p:bldAsOne><p:extLst/></p:bldAsOne>"),
            graphic_timing("5", "88", "", "<p:bldAsOne/><p:bldAsOne/>"),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart/><a:bldChart/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart><a:ext/></a:bldChart></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><x:bldChart xmlns:x=\"urn:foreign\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<x:bldSub xmlns:x=\"urn:foreign\"><a:bldChart/></x:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart bld=\"rows\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldDgm bld=\"rows\"/></p:bldSub>",
            ),
            graphic_timing(
                "5",
                "88",
                "",
                "<p:bldSub><a:bldChart animBg=\"yes\"/></p:bldSub>",
            ),
            graphic_timing("5", "88", r#" uiExpand="yes""#, "<p:bldAsOne/>"),
            graphic_timing("6", "88", "", "<p:bldSub><a:bldChart/></p:bldSub>"),
            graphic_timing("5", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>"),
            graphic_timing("3", "88", "", "<p:bldAsOne/>"),
            valid_chart.replace(r#" spid="5" grpId="88""#, r#" spid="5""#),
            valid_chart.replace(r#" spid="5" grpId="88""#, r#" grpId="88""#),
        ];
        for timing in cases {
            assert!(
                Sequence::parse_slide_xml(slide_with_graphic_hosts(&timing).as_bytes()).is_err()
            );
        }

        let spoofed_host = slide_with_graphic_hosts(&valid_chart).replace(
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#,
            r#"<x:chart xmlns:x="urn:foreign"/>"#,
        );
        assert!(Sequence::parse_slide_xml(spoofed_host.as_bytes()).is_err());

        let chart_marker =
            r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>"#;
        let marker_elsewhere = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, "")
            .replace(
                r#"</p:nvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
                &format!(r#"</p:nvGraphicFramePr>{chart_marker}<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#),
            );
        assert!(Sequence::parse_slide_xml(marker_elsewhere.as_bytes()).is_err());

        let nested_marker = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, &format!(r"<a:ext>{chart_marker}</a:ext>"));
        assert!(Sequence::parse_slide_xml(nested_marker.as_bytes()).is_err());

        let duplicate_marker = slide_with_graphic_hosts(&valid_chart)
            .replace(chart_marker, &format!(r"{chart_marker}{chart_marker}"));
        assert!(Sequence::parse_slide_xml(duplicate_marker.as_bytes()).is_err());

        let ambiguous_marker = slide_with_graphic_hosts(&valid_chart).replace(
            chart_marker,
            &format!(r#"{chart_marker}<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#),
        );
        assert!(Sequence::parse_slide_xml(ambiguous_marker.as_bytes()).is_err());

        let valid_diagram = graphic_timing("6", "88", "", "<p:bldSub><a:bldDgm/></p:bldSub>");
        let diagram_marker =
            r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>"#;
        let nested_diagram_marker = slide_with_graphic_hosts(&valid_diagram)
            .replace(diagram_marker, &format!(r"<a:ext>{diagram_marker}</a:ext>"));
        assert!(Sequence::parse_slide_xml(nested_diagram_marker.as_bytes()).is_err());

        let duplicate = valid_chart.replace(
            "</p:bldLst>",
            r#"<p:bldGraphic spid="5" grpId="88"><p:bldAsOne/></p:bldGraphic></p:bldLst>"#,
        );
        assert!(
            Sequence::parse_slide_xml(slide_with_graphic_hosts(&duplicate).as_bytes()).is_err()
        );
    }

    #[test]
    fn validates_programmatic_graphic_builds_and_combined_build_boundary() {
        let targets = HashSet::from([5]);
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(5, Effect::Fade).with_group_id(88));
        sequence.add_graphic_build(GraphicBuild::chart(5, GroupId::new(88)));
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.graphic_builds.push(sequence.graphic_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.graphic_builds.pop();
        sequence.graphic_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        let mut oversized = Sequence::new();
        oversized.graphic_builds =
            vec![GraphicBuild::as_one(5, GroupId::new(88)); MAX_ANIMATION_BUILDS + 1];
        assert!(oversized.to_xml_for_slide(&targets).is_err());
    }

    #[test]
    fn parses_preserves_and_writes_complete_ole_chart_build_grammar() {
        let timing = ole_chart_timing(
            "5",
            "91",
            r#" uiExpand="true" bld="categoryEl" animBg="false""#,
        );
        let mut sequence =
            Sequence::parse_slide_xml(slide_with_ole_chart(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.ole_chart_builds,
            vec![OleChartBuild {
                shape_id: 5,
                group_id: GroupId::new(91),
                ui_expand: true,
                build_type: OleChartBuildType::CategoryElement,
                animate_background: false,
            }]
        );
        assert_eq!(sequence.to_xml(), timing);

        sequence.animations[0].duration = Duration::Finite(750);
        let canonical = sequence.to_xml();
        assert!(canonical.contains(
            r#"<p:bldOleChart spid="5" grpId="91" uiExpand="1" bld="categoryEl" animBg="0"/>"#
        ));
        let reparsed =
            Sequence::parse_slide_xml(slide_with_ole_chart(&canonical).as_bytes()).unwrap();
        assert_eq!(reparsed, sequence);
    }

    #[test]
    fn accepts_every_ole_chart_build_token_and_schema_defaults() {
        for token in ["allAtOnce", "series", "category", "seriesEl", "categoryEl"] {
            let timing = ole_chart_timing("5", "91", &format!(r#" bld="{token}""#));
            assert!(Sequence::parse_slide_xml(slide_with_ole_chart(&timing).as_bytes()).is_ok());
        }
        let timing = ole_chart_timing("5", "91", "");
        let sequence = Sequence::parse_slide_xml(slide_with_ole(&timing).as_bytes()).unwrap();
        assert_eq!(
            sequence.ole_chart_builds[0].build_type,
            OleChartBuildType::AllAtOnce
        );
        assert!(!sequence.ole_chart_builds[0].ui_expand);
        assert!(sequence.ole_chart_builds[0].animate_background);
    }

    #[test]
    fn rejects_hostile_invalid_and_non_chart_ole_builds() {
        let valid = ole_chart_timing("5", "91", "");
        let cases = [
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" bld="rows""#)),
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" uiExpand="yes""#)),
            slide_with_ole_chart(&ole_chart_timing("5", "91", r#" animBg="yes""#)),
            slide_with_ole_chart(&ole_chart_timing("99", "91", "")),
            slide(&valid),
            slide_with_ole_chart(&ole_chart_timing("3", "91", "")),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart spid="5"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart grpId="91"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<p:bldOleChart spid="5" grpId="91"><p:extLst/></p:bldOleChart>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:bldOleChart spid="5" grpId="91"/>"#,
                r#"<x:bldOleChart xmlns:x="urn:foreign" spid="5" grpId="91"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<p:oleObj progId="Word.Document.12"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<x:oleObj xmlns:x="urn:foreign" progId="MSGraph.Chart.8"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj progId="MSGraph.Chart.8"/></a:graphicData>"#,
                r#"<p:oleObj progId="MSGraph.Chart.8"/><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"/>"#,
            ),
            slide_with_ole_chart(&valid).replace(
                r#"<p:oleObj progId="MSGraph.Chart.8"/>"#,
                r#"<p:oleObj progId="MSGraph.Chart.8"/><p:oleObj progId="MSGraph.Chart.8"/>"#,
            ),
        ];
        for xml in cases {
            assert!(Sequence::parse_slide_xml(xml.as_bytes()).is_err());
        }

        let duplicate = slide_with_ole_chart(&valid).replace(
            "</p:bldLst>",
            r#"<p:bldOleChart spid="5" grpId="91"/></p:bldLst>"#,
        );
        assert!(Sequence::parse_slide_xml(duplicate.as_bytes()).is_err());
    }

    #[test]
    fn validates_programmatic_ole_chart_builds_and_combined_boundary() {
        let targets = HashSet::from([5]);
        let mut sequence = Sequence::new();
        sequence.add(EffectInstance::new(5, Effect::Fade).with_group_id(91));
        sequence.add_ole_chart_build(
            OleChartBuild::new(5, GroupId::new(91))
                .with_ui_expand(true)
                .with_build_type(OleChartBuildType::Series)
                .with_animate_background(false),
        );
        assert!(sequence.to_xml_for_slide(&targets).is_ok());
        sequence.ole_chart_builds.push(sequence.ole_chart_builds[0]);
        assert!(sequence.to_xml_for_slide(&targets).is_err());
        sequence.ole_chart_builds.pop();
        sequence.ole_chart_builds[0].shape_id = 99;
        assert!(sequence.to_xml_for_slide(&targets).is_err());

        let mut oversized = Sequence::new();
        oversized.ole_chart_builds =
            vec![OleChartBuild::new(5, GroupId::new(91)); MAX_ANIMATION_BUILDS + 1];
        assert!(oversized.to_xml_for_slide(&targets).is_err());
    }
}

// Package relationship validation remains part of the canonical animation boundary.
