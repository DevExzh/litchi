//! Slide-transition values and their bounded `PresentationML` codec.
//!
//! Effect-specific payloads make invalid direction combinations
//! unrepresentable. The common path is concise:
//!
//! ```
//! use litchi_pptx::transition::{Kind, Side, Speed, Transition};
//!
//! let transition = Transition::new(Kind::Push(Side::Left))
//!     .with_speed(Speed::Fast)
//!     .with_click(false);
//! assert_eq!(transition.speed(), Speed::Fast);
//! ```
//!
//! Direction domains are checked by the type system:
//!
//! ```compile_fail
//! use litchi_pptx::transition::{Axis, Kind};
//!
//! // Push accepts a side, never an axis.
//! let _ = Kind::Push(Axis::Horizontal);
//! ```

mod model;
mod reader;
mod writer;

pub use crate::time::Offset;
pub use model::{
    Axis, Corner, FlyThrough, Glitter, GlitterPattern, InOut, Kind, LeftRight, MAX_MS,
    MAX_PRESET_NAME_BYTES, Morph, Ms, Origin, Preset, Prism, Raw, Reveal, Ripple, Shape, Shred,
    ShredPattern, Side, Speed, Spokes, TimeError, Transition,
};
pub use reader::{Limits, read, read_with};
pub use writer::{write, write_to};

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::Error;
    use std::mem::size_of;

    const STANDARD_COVER: &[u8] =
        include_bytes!("../../../../test-data/ooxml/pptx/transitions/standard_cover.xml");
    const STANDARD_EFFECT_OPTIONS: &[u8] =
        include_bytes!("../../../../test-data/ooxml/pptx/transitions/standard_effect_options.xml");

    #[test]
    fn presets_have_checked_durations() {
        assert_eq!(Speed::Fast.duration().get(), 500);
        assert_eq!(Speed::Medium.duration().get(), 1000);
        assert_eq!(Speed::Slow.duration().get(), 1500);
        assert!(Ms::new(MAX_MS).is_ok());
        assert!(Ms::new(MAX_MS + 1).is_err());
    }

    #[test]
    fn concise_builder_keeps_timing_typed() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<Transition>();

        let after = Ms::new(3000).unwrap();
        let value = Transition::new(Kind::Fade { black: None })
            .with_speed(Speed::Fast)
            .with_after(after);

        assert_eq!(value.kind(), &Kind::Fade { black: None });
        assert_eq!(value.speed(), Speed::Fast);
        assert_eq!(value.after(), Some(after));
        assert!(
            size_of::<Transition>() <= 64,
            "the common transition value should fit in one cache line"
        );
    }

    #[test]
    fn custom_duration_uses_compatibility_markup_and_round_trips() {
        let value = Transition::new(Kind::Fade { black: None })
            .with_speed(Speed::Fast)
            .with_duration(Ms::new(750).unwrap())
            .with_click(false)
            .with_after(Ms::new(1250).unwrap());

        let xml = write(&value).unwrap();
        assert!(xml.contains(r#"<mc:Choice Requires="p14">"#));
        assert!(xml.contains(r#"p14:dur="750""#));
        assert!(!xml.contains(r#"<p:transition spd="fast" dur="#));
        assert!(xml.contains(
            r#"<mc:Fallback><p:transition spd="fast" advClick="0" advTm="1250"><p:fade/>"#
        ));
        assert_eq!(xml.matches("<p:fade/>").count(), 2);
        assert!(parse_fragment(&xml).same_semantics(&value));
    }

    #[test]
    fn ripple_uses_a_standard_fade_fallback_and_round_trips() {
        let value = Transition::new(Kind::Ripple(Ripple::LeftDown))
            .with_speed(Speed::Slow)
            .with_duration(Ms::new(1500).unwrap())
            .with_click(false)
            .with_after(Ms::new(4250).unwrap());

        let xml = write(&value).unwrap();
        assert!(xml.contains(r#"<p14:ripple dir="ld"/>"#));
        assert!(xml.contains("<p:fade/>"));
        assert!(parse_fragment(&xml).same_semantics(&value));
    }

    #[test]
    fn morph_uses_powerpoint_2015_choice_and_round_trips() {
        let value = Transition::new(Kind::Morph(Morph::ByWord))
            .with_speed(Speed::Fast)
            .with_duration(Ms::new(750).unwrap())
            .with_click(false)
            .with_after(Ms::new(1250).unwrap());

        let xml = write(&value).unwrap();
        assert!(xml.contains(r#"Requires="p14 p159""#));
        assert!(xml.contains(r#"<p159:morph option="byWord"/>"#));
        assert!(xml.contains("<p:fade/>"));
        assert!(parse_fragment(&xml).same_semantics(&value));
    }

    #[test]
    fn preset_uses_powerpoint_2012_choice_and_preserves_inversion() {
        let preset = Preset::with_options("wind", true, false).unwrap();
        let value = Transition::new(Kind::Preset(preset.clone())).with_speed(Speed::Slow);

        let xml = write(&value).unwrap();
        assert!(xml.contains(r#"Requires="p15""#));
        assert!(xml.contains(r#"<p15:prstTrans prst="wind" invX="1"/>"#));
        assert!(parse_fragment(&xml).same_semantics(&value));
        assert_eq!(preset.name(), Some("wind"));
        assert!(preset.invert_x());
        assert!(!preset.invert_y());
    }

    #[test]
    fn preset_escapes_xml_name_and_round_trips() {
        let preset = Preset::new("wind&<\"\t\n\r").unwrap();
        let value = Transition::new(Kind::Preset(preset));
        let xml = write(&value).unwrap();
        assert!(xml.contains(r#"prst="wind&amp;&lt;&quot;&#x9;&#xA;&#xD;""#));
        assert!(parse_fragment(&xml).same_semantics(&value));
    }

    #[test]
    fn typed_extension_read_preserves_unknown_attributes_until_semantic_change() {
        let xml = transition_xml(
            r#"<p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byChar" future="keep"/>"#,
        );
        let value = read(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(value.kind(), &Kind::Morph(Morph::ByChar));
        let preserved = write(&value).unwrap();
        assert!(preserved.contains(r#"future="keep""#));
        assert!(parse_fragment(&preserved).same_semantics(&value));

        let mut changed = value.clone();
        changed.set_kind(Kind::Morph(Morph::ByObject));
        let changed = write(&changed).unwrap();
        assert!(!changed.contains(r#"future="keep""#));
        assert!(changed.contains(r#"option="byObject""#));
    }

    #[test]
    fn preset_absent_optional_attributes_round_trip_without_normalization() {
        let xml = transition_xml(
            r#"<p15:prstTrans xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" invX="false"/>"#,
        );
        let value = read(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(value.kind(), &Kind::Preset(Preset::without_name()));
        let output = write(&value).unwrap();
        assert!(output.contains(r#"invX="false""#));
        assert!(!output.contains(r#"invY=""#));
    }

    #[test]
    fn preset_default_presence_is_semantic_noop_but_wire_distinct() {
        let absent = read(
            transition_xml(
                r#"<p15:prstTrans xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"/>"#,
            )
            .as_bytes(),
        )
        .unwrap()
        .unwrap();
        let explicit_false = read(
            transition_xml(
                r#"<p15:prstTrans xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" invX="false"/>"#,
            )
            .as_bytes(),
        )
        .unwrap()
        .unwrap();

        assert!(absent.same_semantics(&explicit_false));
        assert_ne!(absent, explicit_false);
        let absent_xml = write(&absent).unwrap();
        let explicit_xml = write(&explicit_false).unwrap();
        assert!(!absent_xml.contains(" invX="));
        assert!(explicit_xml.contains(r#"invX="false""#));
    }

    #[test]
    fn rejects_invalid_morph_option_and_bounded_preset_name() {
        let invalid = transition_xml(
            r#"<p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" option="byShape"/>"#,
        );
        assert!(matches!(
            read(invalid.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("morph transition option")
        ));

        let oversized = "x".repeat(MAX_PRESET_NAME_BYTES + 1);
        assert!(matches!(
            Preset::new(oversized),
            Err(Error::Limit { resource, limit })
                if resource == "preset transition name bytes" && limit == MAX_PRESET_NAME_BYTES
        ));
    }

    #[test]
    fn reads_local_standard_fixtures() {
        let cover = read(STANDARD_COVER).unwrap().unwrap();
        assert_eq!(cover.speed(), Speed::Fast);
        assert!(!cover.click());
        assert_eq!(cover.after().map(Ms::get), Some(750));
        assert_eq!(cover.kind(), &Kind::Cover(Origin::RightDown));

        let fade = read(STANDARD_EFFECT_OPTIONS).unwrap().unwrap();
        assert_eq!(fade.kind(), &Kind::Fade { black: Some(true) });
    }

    #[test]
    fn standard_effect_payloads_are_type_specific() {
        let cases = [
            (Kind::Push(Side::Down), r#"<p:push dir="d"/>"#),
            (
                Kind::Split {
                    axis: Axis::Vertical,
                    toward: Some(InOut::In),
                },
                r#"<p:split orient="vert" dir="in"/>"#,
            ),
            (Kind::Uncover(Origin::LeftUp), r#"<p:pull dir="lu"/>"#),
            (Kind::Cover(Origin::RightDown), r#"<p:cover dir="rd"/>"#),
            (Kind::Blinds(Axis::Vertical), r#"<p:blinds dir="vert"/>"#),
            (
                Kind::RandomBars(Axis::Vertical),
                r#"<p:randomBar dir="vert"/>"#,
            ),
            (Kind::Strips(Corner::LeftDown), r#"<p:strips dir="ld"/>"#),
            (Kind::Comb(Axis::Vertical), r#"<p:comb dir="vert"/>"#),
            (Kind::Wheel(Spokes::Eight), r#"<p:wheel spokes="8"/>"#),
            (Kind::Newsflash, "<p:newsflash/>"),
            (Kind::Shape(Shape::Plus), "<p:plus/>"),
        ];

        for (kind, expected) in cases {
            let value = Transition::new(kind);
            let xml = write(&value).unwrap();
            assert!(xml.contains(expected), "expected {expected:?} in {xml:?}");
            assert!(parse_fragment(&xml).same_semantics(&value));
        }
    }

    #[test]
    fn all_powerpoint_2010_transition_effects_are_typed_and_round_trip() {
        let cases = [
            (
                Kind::Conveyor(LeftRight::Right),
                r#"<p14:conveyor dir="r"/>"#,
            ),
            (Kind::Doors(Axis::Vertical), r#"<p14:doors dir="vert"/>"#),
            (Kind::Ferris(LeftRight::Left), r#"<p14:ferris dir="l"/>"#),
            (Kind::Flash, r#"<p14:flash/>"#),
            (Kind::Flip(LeftRight::Right), r#"<p14:flip dir="r"/>"#),
            (
                Kind::FlyThrough(FlyThrough::new(InOut::Out, true)),
                r#"<p14:flythrough dir="out" hasBounce="1"/>"#,
            ),
            (Kind::Gallery(LeftRight::Left), r#"<p14:gallery dir="l"/>"#),
            (
                Kind::Glitter(Glitter::new(Side::Down, GlitterPattern::Hexagon)),
                r#"<p14:glitter dir="d" pattern="hexagon"/>"#,
            ),
            (Kind::Honeycomb, r#"<p14:honeycomb/>"#),
            (Kind::Pan(Side::Up), r#"<p14:pan dir="u"/>"#),
            (
                Kind::Prism(Prism::new(Side::Right, true, true)),
                r#"<p14:prism dir="r" isContent="1" isInverted="1"/>"#,
            ),
            (
                Kind::Reveal(Reveal::new(LeftRight::Right, true)),
                r#"<p14:reveal dir="r" thruBlk="1"/>"#,
            ),
            (
                Kind::Shred(Shred::new(ShredPattern::Rectangle, InOut::Out)),
                r#"<p14:shred pattern="rectangle" dir="out"/>"#,
            ),
            (Kind::Switch(LeftRight::Right), r#"<p14:switch dir="r"/>"#),
            (Kind::Vortex(Side::Left), r#"<p14:vortex dir="l"/>"#),
            (Kind::Warp(InOut::Out), r#"<p14:warp dir="out"/>"#),
            (
                Kind::WheelReverse(Spokes::Eight),
                r#"<p14:wheelReverse spokes="8"/>"#,
            ),
            (
                Kind::Window(Axis::Horizontal),
                r#"<p14:window dir="horz"/>"#,
            ),
        ];

        for (kind, expected) in cases {
            let value = Transition::new(kind);
            let xml = write(&value).unwrap();
            assert!(xml.contains(r#"<mc:Choice Requires="p14">"#));
            assert!(xml.contains(expected), "expected {expected:?} in {xml:?}");
            assert!(xml.contains("<p:fade/>"));
            assert!(parse_fragment(&xml).same_semantics(&value));
        }
    }

    #[test]
    fn p14_effect_defaults_and_boolean_lexical_forms_are_checked() {
        let cases = [
            (
                r#"<p14:conveyor xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>"#,
                Kind::Conveyor(LeftRight::Unspecified),
            ),
            (
                r#"<p14:doors xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>"#,
                Kind::Doors(Axis::Horizontal),
            ),
            (
                r#"<p14:flythrough xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" dir="in" hasBounce="true"/>"#,
                Kind::FlyThrough(FlyThrough::new(InOut::In, true)),
            ),
            (
                r#"<p14:glitter xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" pattern="diamond"/>"#,
                Kind::Glitter(Glitter::new(Side::Left, GlitterPattern::Diamond)),
            ),
            (
                r#"<p14:prism xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" isContent="false" isInverted="0"/>"#,
                Kind::Prism(Prism::new(Side::Left, false, false)),
            ),
            (
                r#"<p14:reveal xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" thruBlk="false"/>"#,
                Kind::Reveal(Reveal::new(LeftRight::Left, false)),
            ),
            (
                r#"<p14:shred xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>"#,
                Kind::Shred(Shred::new(ShredPattern::Strip, InOut::In)),
            ),
            (
                r#"<p14:wheelReverse xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>"#,
                Kind::WheelReverse(Spokes::Four),
            ),
        ];

        for (effect, expected) in cases {
            let value = read(transition_xml_with_p14(effect).as_bytes())
                .unwrap()
                .unwrap();
            assert_eq!(value.kind(), &expected);
        }
    }

    #[test]
    fn p14_effects_reject_values_outside_their_schema_domains() {
        for effect in [
            r#"<p14:conveyor dir="u"/>"#,
            r#"<p14:doors dir="left"/>"#,
            r#"<p14:flythrough dir="sideways"/>"#,
            r#"<p14:glitter pattern="square"/>"#,
            r#"<p14:prism dir="center"/>"#,
            r#"<p14:reveal dir="u"/>"#,
            r#"<p14:shred pattern="circle"/>"#,
            r#"<p14:wheelReverse spokes="6"/>"#,
        ] {
            assert!(
                matches!(
                    read(transition_xml_with_p14(effect).as_bytes()),
                    Err(Error::Invalid(_))
                ),
                "unexpectedly accepted {effect:?}"
            );
        }
    }

    #[test]
    fn p14_source_attributes_are_preserved_until_a_typed_edit() {
        let value = read(
            transition_xml_with_p14(
                r#"<p14:prism xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" dir="r" isContent="true" future="keep"/>"#,
            )
            .as_bytes(),
        )
        .unwrap()
        .unwrap();
        let unchanged = write(&value).unwrap();
        assert!(unchanged.contains(r#"future="keep""#));

        let mut changed = value.clone();
        changed.set_kind(Kind::Prism(Prism::new(Side::Up, true, false)));
        let changed = write(&changed).unwrap();
        assert!(!changed.contains(r#"future="keep""#));
        assert!(changed.contains(r#"<p14:prism dir="u" isContent="1"/>"#));
        assert!(changed.contains("<p:fade/>"));
    }

    #[test]
    fn rejects_invalid_directions_spokes_and_timing() {
        for effect in [r#"<p:push dir="horz"/>"#, r#"<p:wheel spokes="6"/>"#] {
            let xml = transition_xml(effect);
            assert!(matches!(read(xml.as_bytes()), Err(Error::Invalid(_))));
        }

        let xml = transition_xml(r"<p:fade/>").replacen(
            "<p:transition>",
            r#"<p:transition advTm="2147483648">"#,
            1,
        );
        assert!(matches!(read(xml.as_bytes()), Err(Error::Invalid(_))));
    }

    #[test]
    fn rejects_unit_suffix_on_integer_timed_advance() {
        let xml = transition_xml(r"<p:fade/>").replacen(
            "<p:transition>",
            r#"<p:transition advTm="750ms">"#,
            1,
        );
        assert!(matches!(
            read(xml.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("automatic-advance delay")
        ));

        let xml = transition_xml(r"<p:fade/>").replacen(
            "<p:transition>",
            r#"<p:transition advTm="2147483647">"#,
            1,
        );
        assert_eq!(
            read(xml.as_bytes())
                .unwrap()
                .unwrap()
                .after()
                .unwrap()
                .get(),
            i32::MAX as u32
        );
    }

    #[test]
    fn p14_duration_accepts_exact_units_and_fractional_milliseconds() {
        for (lexical, canonical, legacy) in [
            ("500µs", "0.5", None),
            ("1.25s", "1250", Some(1250)),
            ("100ns", "0.0001", None),
            ("1.5", "1.5", None),
            ("2min", "120000", Some(120_000)),
        ] {
            let xml = transition_xml_with_p14(r#"<p:fade/>"#).replacen(
                "<p:transition>",
                &format!(r#"<p:transition p14:dur="{lexical}">"#),
                1,
            );
            let value = read(xml.as_bytes()).unwrap().unwrap();
            assert_eq!(value.duration_offset().unwrap().as_str(), canonical);
            assert_eq!(value.duration().map(Ms::get), legacy);
            let output = write(&value).unwrap();
            assert!(output.contains(&format!(r#"p14:dur="{canonical}""#)));
            assert!(parse_fragment(&output).same_semantics(&value));
        }

        let invalid = transition_xml_with_p14(r#"<p:fade/>"#).replacen(
            "<p:transition>",
            r#"<p:transition p14:dur="1quarter">"#,
            1,
        );
        assert!(matches!(
            read(invalid.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("transition duration")
        ));

        let invalid = transition_xml_with_p14(r#"<p:fade/>"#).replacen(
            "<p:transition>",
            r#"<p:transition p14:dur=" 1s ">"#,
            1,
        );
        assert!(matches!(
            read(invalid.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("transition duration")
        ));
    }

    #[test]
    fn token_and_boolean_whitespace_is_schema_collapsed() {
        let xml = transition_xml_with_p14(
            r#"<p14:flythrough xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" dir="  out &#x9;" hasBounce=" &#xA; true &#xD; "/>"#,
        );
        let value = read(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(
            value.kind(),
            &Kind::FlyThrough(FlyThrough::new(InOut::Out, true))
        );
    }

    #[test]
    fn typed_effects_reject_nested_children_and_character_data() {
        for effect in [
            r#"<p14:flash xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:child/></p14:flash>"#,
            r#"<p14:honeycomb xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">unexpected</p14:honeycomb>"#,
        ] {
            assert!(matches!(
                read(transition_xml(effect).as_bytes()),
                Err(Error::Invalid(message)) if message.contains("typed transition effects")
            ));
        }
    }

    #[test]
    fn rejects_duplicate_effects() {
        let xml = transition_xml("<p:fade/><p:cut/>");
        assert!(matches!(
            read(xml.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("more than one")
        ));
    }

    #[test]
    fn exact_equality_includes_retained_wire_state() {
        let plain = read(transition_xml("<p:fade/>").as_bytes())
            .unwrap()
            .unwrap();
        let extended = read(transition_xml(r#"<p:fade future="1"/>"#).as_bytes())
            .unwrap()
            .unwrap();
        assert!(plain.same_semantics(&extended));
        assert_ne!(plain, extended);

        let xml = std::sync::Arc::<str>::from("<x:future/>");
        let portable = Raw {
            xml: xml.clone(),
            portable: true,
        };
        let contextual = Raw {
            xml,
            portable: false,
        };
        assert_ne!(portable, contextual);
    }

    #[test]
    fn rejects_doctype_without_expanding_entities() {
        let xml = r#"<!DOCTYPE p:sld [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:transition><p:fade/></p:transition></p:sld>"#;
        assert!(matches!(
            read(xml.as_bytes()),
            Err(Error::Invalid(message)) if message.contains("DOCTYPE")
        ));
    }

    #[test]
    fn retains_unknown_effect_and_extension_children_inertly() {
        let xml = transition_xml(
            r#"<p15:glitter xmlns:p15="urn:example:p15" amount="7"><p15:data/></p15:glitter><p:extLst><p:ext uri="urn:test"/></p:extLst>"#,
        );
        let value = read(xml.as_bytes()).unwrap().unwrap();
        let Kind::Raw(raw) = value.kind() else {
            panic!("unknown effect should remain raw")
        };
        assert!(raw.xml().contains("p15:glitter"));
        assert_eq!(value.preserved().count(), 1);

        let output = write(&value).unwrap();
        assert!(output.contains("p15:glitter"));
        assert!(output.contains("<p:extLst>"));
    }

    #[test]
    fn inspects_but_does_not_emit_raw_children_with_ancestor_only_prefixes() {
        let xml = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p15="urn:example:p15"><p:transition><p15:glitter/></p:transition></p:sld>"#;
        let value = read(xml.as_bytes()).unwrap().unwrap();
        let Kind::Raw(raw) = value.kind() else {
            panic!("unknown effect should remain raw")
        };
        assert!(!raw.is_portable());
        assert!(raw.xml().contains("p15:glitter"));
        let mut output = "unchanged".to_string();
        assert!(matches!(
            write_to(&value, &mut output),
            Err(Error::Invalid(message)) if message.contains("namespace prefix")
        ));
        assert_eq!(output, "unchanged");
    }

    #[test]
    fn input_depth_node_and_retention_limits_are_enforced() {
        let tiny = Limits::new(4096, 2, 10, 16).unwrap();
        let deep = transition_xml("<p:extLst><p:ext/></p:extLst>");
        assert!(matches!(
            read_with(deep.as_bytes(), tiny),
            Err(Error::Limit { .. })
        ));

        let tiny = Limits::new(4096, 16, 2, 4096).unwrap();
        let nodes = transition_xml("<p:fade/>");
        assert!(matches!(
            read_with(nodes.as_bytes(), tiny),
            Err(Error::Limit { .. })
        ));

        let tiny = Limits::new(4096, 16, 20, 8).unwrap();
        let raw = transition_xml(r#"<x:future xmlns:x="urn:x"/>"#);
        assert!(matches!(
            read_with(raw.as_bytes(), tiny),
            Err(Error::Limit { .. })
        ));
    }

    fn transition_xml(effect: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:transition>{effect}</p:transition></p:sld>"#
        )
    }

    fn parse_fragment(xml: &str) -> Transition {
        let xml = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{xml}</p:sld>"#
        );
        read(xml.as_bytes()).unwrap().unwrap()
    }

    fn transition_xml_with_p14(effect: &str) -> String {
        format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:transition>{effect}</p:transition></p:sld>"#
        )
    }
}
