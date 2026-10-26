//! Differential tests against the generated Prost schema.

use std::collections::BTreeMap;

use litchi_iwa_common::WireLimits;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_protos::tswp;
use litchi_iwa_text::appearance::{Background, Outline, Shadow};
use litchi_iwa_text::character::{TextCapitalization, TextPointSize, TextStyle};
use litchi_iwa_text::font::Font;
use litchi_iwa_text::paragraph::format::Alignment;
use prost::Message;

use super::*;

const LIMITS: WireLimits = WireLimits::DEFAULT;

fn named_root(stylesheet: u64) -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: litchi_iwa_protos::tss::StyleArchive {
            name: Some("Body".to_owned()),
            stylesheet: Some(litchi_iwa_protos::tsp::Reference {
                identifier: stylesheet,
                ..Default::default()
            }),
            ..Default::default()
        },
        char_properties: Some(tswp::CharacterStylePropertiesArchive {
            font_size: Some(11.0),
            ..Default::default()
        }),
        para_properties: Some(tswp::ParagraphStylePropertiesArchive::default()),
        ..Default::default()
    }
    .encode_to_vec()
}

fn styles() -> BTreeMap<u64, Vec<u8>> {
    BTreeMap::from([(10, named_root(1))])
}

fn properties() -> Vec<Property> {
    let color = Rgba::new(0.25, 0.5, 0.75, 1.0, RgbColorSpace::Srgb).unwrap();
    vec![
        Property::TextStyle(
            TextStyle::new(TextPointSize::from_points(18.0).unwrap()).with_bold(true),
        ),
        Property::TextFont(Font::named("Helvetica").unwrap()),
        Property::TextColor(color),
        Property::TextCapitalization(TextCapitalization::TitleCase),
        Property::TextOutline(Outline::standard()),
        Property::TextShadow(Shadow::standard()),
        Property::TextBackground(Background::Color(color)),
        Property::Alignment(Alignment::Center),
        Property::Flow(Flow::new().with_keep_with_next(true)),
        Property::DefaultTabInterval(DefaultInterval::from_points(24.0).unwrap()),
    ]
}

#[test]
fn derived_variations_are_canonical_prost_encodings() {
    let styles = styles();
    for property in properties() {
        let plan = plan(
            &styles,
            Target::new(10, false),
            &Change::Set(property.clone()),
            LIMITS,
        )
        .unwrap();
        let Plan::Derive(variation) = plan else {
            panic!("{property:?} must derive a variation, got {plan:?}");
        };
        let decoded = tswp::ParagraphStyleArchive::decode(variation.payload()).unwrap();
        assert_eq!(decoded.encode_to_vec(), variation.payload(), "{property:?}");
        assert_eq!(decoded.super_.parent.unwrap().identifier, 10);
        assert_eq!(variation.references()[0], 10);
        let mut overlay = styles.clone();
        overlay.insert(11, variation.payload().to_vec());
        assert_eq!(
            effective(&overlay, 11, property.kind(), LIMITS).unwrap(),
            property
        );
    }
}

#[test]
fn exclusive_variations_are_rewritten_then_collapsed() {
    let mut styles = styles();
    let first = Property::Alignment(Alignment::Right);
    let Plan::Derive(variation) =
        plan(&styles, Target::new(10, true), &Change::Set(first), LIMITS).unwrap()
    else {
        panic!("first set derives");
    };
    styles.insert(11, variation.into_payload());
    let second = Property::Alignment(Alignment::Left);
    let Plan::Rewrite(rewritten) = plan(
        &styles,
        Target::new(11, true),
        &Change::Set(second.clone()),
        LIMITS,
    )
    .unwrap() else {
        panic!("an exclusive managed variation is rewritten");
    };
    styles.insert(11, rewritten.into_payload());
    assert_eq!(
        effective(&styles, 11, PropertyKind::Alignment, LIMITS).unwrap(),
        second
    );
    assert_eq!(
        plan(&styles, Target::new(11, true), &Change::Set(second), LIMITS).unwrap(),
        Plan::Unchanged
    );
    assert_eq!(
        plan(
            &styles,
            Target::new(11, true),
            &Change::Reset(PropertyKind::Alignment),
            LIMITS
        )
        .unwrap(),
        Plan::Collapse {
            parent: 10,
            stylesheet: 1
        }
    );
    assert_eq!(
        plan(
            &styles,
            Target::new(11, false),
            &Change::Reset(PropertyKind::Alignment),
            LIMITS
        )
        .unwrap(),
        Plan::Unchanged
    );
}

#[test]
fn a_named_style_is_never_rewritten() {
    let styles = styles();
    let plan = plan(
        &styles,
        Target::new(10, true),
        &Change::Set(Property::Alignment(Alignment::Right)),
        LIMITS,
    )
    .unwrap();
    assert!(matches!(plan, Plan::Derive(_)));
    assert_eq!(
        super::plan(
            &styles,
            Target::new(10, true),
            &Change::Reset(PropertyKind::Alignment),
            LIMITS
        )
        .unwrap(),
        Plan::Unchanged
    );
}

#[test]
fn missing_styles_and_cycles_are_invalid_sources() {
    let styles = styles();
    assert!(matches!(
        effective(&styles, 99, PropertyKind::TextFont, LIMITS),
        Err(Error::InvalidSource { .. })
    ));
    let cyclic = BTreeMap::from([(
        5,
        tswp::ParagraphStyleArchive {
            super_: litchi_iwa_protos::tss::StyleArchive {
                parent: Some(litchi_iwa_protos::tsp::Reference {
                    identifier: 5,
                    ..Default::default()
                }),
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec(),
    )]);
    assert!(matches!(
        effective(&cyclic, 5, PropertyKind::TextFont, LIMITS),
        Err(Error::InvalidSource { .. })
    ));
}

#[test]
fn limits_bound_the_chain_walk() {
    let styles = styles();
    let tight = WireLimits::DEFAULT.with_input_bytes(4).unwrap();
    assert!(matches!(
        effective(&styles, 10, PropertyKind::TextFont, tight),
        Err(Error::Limit { .. })
    ));
}
