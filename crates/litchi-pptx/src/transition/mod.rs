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

pub use model::{
    Axis, Corner, InOut, Kind, MAX_MS, Ms, Origin, Raw, Ripple, Shape, Side, Speed, Spokes,
    TimeError, Transition,
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
}
