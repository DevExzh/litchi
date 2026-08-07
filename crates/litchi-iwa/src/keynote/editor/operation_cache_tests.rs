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
