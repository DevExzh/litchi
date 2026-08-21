use super::*;
use crate::protobuf::tsp::Reference;

fn empty_operation() -> KeynoteOperation {
    KeynoteOperation {
        graph: ObjectGraph {
            objects: HashMap::new(),
            archives: HashMap::new(),
        },
        slide_cache: HashMap::new(),
        drawable_storage_cache: HashMap::new(),
    }
}

fn show_operation(messages: Vec<RawMessage>) -> KeynoteOperation {
    KeynoteOperation {
        graph: ObjectGraph {
            objects: HashMap::from([(2, messages)]),
            archives: HashMap::new(),
        },
        slide_cache: HashMap::new(),
        drawable_storage_cache: HashMap::new(),
    }
}

fn show_payload(slides: &[u64]) -> Vec<u8> {
    kn::ShowArchive {
        theme: Reference {
            identifier: 90,
            ..Default::default()
        },
        slide_tree: kn::SlideTreeArchive {
            slides: slides
                .iter()
                .copied()
                .map(|identifier| Reference {
                    identifier,
                    ..Default::default()
                })
                .collect(),
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: Reference {
            identifier: 91,
            ..Default::default()
        },
        ..Default::default()
    }
    .encode_to_vec()
}

#[test]
fn operation_slide_cache_keeps_compact_summaries_bounded() {
    let mut operation = empty_operation();
    for identifier in 1..=(MAX_OPERATION_CACHED_SLIDES as u64 + 1) {
        let slide = kn::SlideArchive {
            title_placeholder: Some(Reference {
                identifier: identifier + 1,
                ..Default::default()
            }),
            body_placeholder: Some(Reference {
                identifier: identifier + 2,
                ..Default::default()
            }),
            owned_drawables: vec![Reference {
                identifier: identifier + 3,
                ..Default::default()
            }],
            ..Default::default()
        };
        operation.remember_slide(identifier, &slide);
    }

    assert_eq!(operation.slide_cache.len(), MAX_OPERATION_CACHED_SLIDES);
    assert!(
        !operation
            .slide_cache
            .contains_key(&(MAX_OPERATION_CACHED_SLIDES as u64 + 1))
    );
    let cached = operation
        .slide_cache
        .get(&1)
        .expect("first slide is cached");
    assert_eq!(cached.title_placeholder, Some(2));
    assert_eq!(cached.body_placeholder, Some(3));
    assert_eq!(&*cached.owned_drawables, &[4]);
}

#[test]
fn operation_slide_cache_skips_unbounded_slide_summaries() {
    let mut operation = empty_operation();
    let slide = kn::SlideArchive {
        owned_drawables: (0..=MAX_OPERATION_CACHED_DRAWABLES_PER_SLIDE)
            .map(|identifier| Reference {
                identifier: identifier as u64,
                ..Default::default()
            })
            .collect(),
        ..Default::default()
    };

    operation.remember_slide(1, &slide);

    assert!(operation.slide_cache.is_empty());
}

#[test]
fn operation_show_snapshot_keeps_theme_and_node_order() {
    let mut payload = show_payload(&[7, 2, 7]);
    payload.extend([0x98, 0x06, 0x07]);
    let source = payload.clone();
    let operation = show_operation(vec![RawMessage {
        type_: SHOW_MESSAGE_TYPE,
        data: payload,
    }]);
    let snapshot = operation.show_snapshot(2).unwrap();

    assert_eq!(snapshot.theme_identifier(), 90);
    assert_eq!(snapshot.slide_node_identifiers(), [7, 2, 7]);
    assert_eq!(operation.graph.objects[&2][0].data, source);
}

#[test]
fn operation_show_snapshot_rejects_missing_or_repeated_owned_payloads() {
    let payload = show_payload(&[7]);
    let missing = show_operation(vec![RawMessage {
        type_: SHOW_MESSAGE_TYPE + 1,
        data: payload.clone(),
    }]);
    assert!(matches!(
        missing.show_snapshot(2),
        Err(Error::InvalidFormat(message)) if message.contains("has no KN.ShowArchive payload")
    ));

    let repeated = show_operation(vec![
        RawMessage {
            type_: SHOW_MESSAGE_TYPE,
            data: payload.clone(),
        },
        RawMessage {
            type_: SHOW_MESSAGE_TYPE,
            data: payload,
        },
    ]);
    assert!(matches!(
        repeated.show_snapshot(2),
        Err(Error::InvalidFormat(message)) if message.contains("repeats its KN.ShowArchive payload")
    ));

    let missing_object = empty_operation();
    assert!(matches!(
        missing_object.show_snapshot(2),
        Err(Error::InvalidFormat(message)) if message.contains("Object 2 is missing")
    ));
}
