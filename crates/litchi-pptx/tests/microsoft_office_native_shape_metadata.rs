//! Explicit inert classification/Designer derivatives of native PowerPoint parents.
use litchi_pptx::{
    Error, Package,
    shape::{
        classification::Outcome,
        designer::{DrawingProperties, Limits, Tag, Tags},
    },
};
use std::{collections::BTreeMap, path::Path};
const SLIDE: &str = "ppt/slides/slide1.xml";
const FILES: [&str; 2] = [
    "powerpoint-navigation-review.pptx",
    "powerpoint-passive-macro.pptm",
];
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(file),
    )
    .unwrap()
}
fn members(p: &Package) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        p.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn preserved(before: &Package, after: &Package, owned: &str) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (n, v) in a {
        if n != owned {
            assert!(b.get(&n) == Some(&v), "changed {n}");
        }
    }
    let av = before.presentation().unwrap();
    let bv = after.presentation().unwrap();
    let a = av.slide(0).unwrap().unwrap();
    let b = bv.slide(0).unwrap().unwrap();
    let a = a.shapes().unwrap();
    let b = b.shapes().unwrap();
    assert_eq!(a.len(), b.len());
    for i in 0..a.len() {
        let x = a.shape(i).unwrap();
        let y = b.shape(i).unwrap();
        assert_eq!(
            (x.id(), x.name(), x.text(), x.transform().unwrap()),
            (y.id(), y.name(), y.text(), y.transform().unwrap())
        );
        if i != 0 || owned != "ppt/slides/slide1.xml" {
            assert_eq!(x.xml().unwrap(), y.xml().unwrap());
        }
    }
}
macro_rules! publish {
    ($source:expr,$commit:expr,$apply:ident) => {{
        let source = &$source;
        let commit = &$commit;
        let f = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
        commit.package().save(f.path()).unwrap();
        let q = Package::open(f.path()).unwrap();
        assert!(members(&q) == members(commit.package()));
        assert!(members(source.$apply(commit.patch()).unwrap().package()) == members(&q));
        assert!(members(q.$apply(&commit.patch().inverse()).unwrap().package()) == members(source));
        q
    }};
}
fn tags(value: &str) -> Tags {
    let mut t = Tags::new();
    t.push(Tag::new("role", value).unwrap()).unwrap();
    t.push(Tag::new("日本語", "A & <B> 🍎").unwrap()).unwrap();
    t
}
#[test]
fn native_pptx_classification_lifecycle_preserves_shapes_and_unowned_parts() {
    for file in FILES {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let source_bytes = p.to_bytes().unwrap();
        assert!(p.shape_classification(0usize, 0usize).unwrap().is_none());
        let mut e = p.edit_shape_classification().unwrap();
        assert!(e.set(0usize, 0usize, Outcome::Header).unwrap().is_none());
        assert!(matches!(
            e.set(0usize, usize::MAX, Outcome::Footer),
            Err(Error::ShapeLookup(
                litchi_pptx::shape::LookupError::IndexOutOfBounds {
                    index: usize::MAX,
                    ..
                }
            ))
        ));
        let c = e.commit().unwrap();
        let mut current = publish!(p, c, apply_shape_classification);
        preserved(&p, &current, SLIDE);
        assert_eq!(
            current
                .shape_classification(0usize, 0usize)
                .unwrap()
                .unwrap()
                .outcome(),
            Some(Outcome::Header)
        );
        assert!(matches!(
            current.apply_shape_classification(c.patch()),
            Err(Error::PatchConflict)
        ));
        for outcome in [Outcome::Footer, Outcome::Watermark, Outcome::None] {
            let mut e = current.edit_shape_classification().unwrap();
            e.set(0usize, 0usize, outcome).unwrap();
            let c = e.commit().unwrap();
            let next = publish!(current, c, apply_shape_classification);
            preserved(&current, &next, SLIDE);
            assert_eq!(
                next.shape_classification(0usize, 0usize)
                    .unwrap()
                    .unwrap()
                    .outcome(),
                Some(outcome)
            );
            current = next;
        }
        let mut e = current.edit_shape_classification().unwrap();
        e.set(0usize, 0usize, Outcome::None).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = current.edit_shape_classification().unwrap();
        assert_eq!(
            e.remove(0usize, 0usize).unwrap().unwrap().outcome(),
            Some(Outcome::None)
        );
        let c = e.commit().unwrap();
        let cleared = publish!(current, c, apply_shape_classification);
        preserved(&current, &cleared, SLIDE);
        assert!(
            cleared
                .shape_classification(0usize, 0usize)
                .unwrap()
                .is_none()
        );
        assert!(p.to_bytes().unwrap() == source_bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_pptx_designer_flags_properties_and_tags_keep_distinct_optional_states() {
    for file in FILES {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let source_bytes = p.to_bytes().unwrap();
        assert!(p.shape_design_element(0usize, 0usize).unwrap().is_none());
        assert!(
            !p.shape_designer_properties(0usize, 0usize)
                .unwrap()
                .is_present()
        );
        let mut e = p.edit_shape_design_element().unwrap();
        e.set(0usize, 0usize, true).unwrap();
        let c = e.commit().unwrap();
        let added = publish!(p, c, apply_shape_design_element);
        preserved(&p, &added, SLIDE);
        assert_eq!(
            added
                .shape_design_element(0usize, 0usize)
                .unwrap()
                .unwrap()
                .value(),
            Some(true)
        );
        let mut e = added.edit_shape_design_element().unwrap();
        e.set(0usize, 0usize, false).unwrap();
        let c = e.commit().unwrap();
        let flagged = publish!(added, c, apply_shape_design_element);
        preserved(&added, &flagged, SLIDE);
        assert_eq!(
            flagged
                .shape_design_element(0usize, 0usize)
                .unwrap()
                .unwrap()
                .value(),
            Some(false)
        );
        let properties = DrawingProperties::new()
            .with_editable(Some(true))
            .with_tags(Some(tags("title")), Limits::new())
            .unwrap();
        let mut e = flagged.edit_shape_designer_properties().unwrap();
        e.set(0usize, 0usize, properties.clone()).unwrap();
        let c = e.commit().unwrap();
        let added = publish!(flagged, c, apply_shape_designer_properties);
        preserved(&flagged, &added, SLIDE);
        assert_eq!(
            added
                .shape_designer_properties(0usize, 0usize)
                .unwrap()
                .properties(),
            Some(&properties)
        );
        assert_eq!(
            added
                .shape_design_element(0usize, 0usize)
                .unwrap()
                .unwrap()
                .value(),
            Some(false)
        );
        let mut e = added.edit_shape_designer_properties().unwrap();
        e.set(0usize, 0usize, properties.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let properties = DrawingProperties::new()
            .with_tags(Some(Tags::new()), Limits::new())
            .unwrap();
        assert_eq!(properties.editable(), None);
        assert!(!properties.effective_editable());
        let mut e = added.edit_shape_designer_properties().unwrap();
        e.set(0usize, 0usize, properties.clone()).unwrap();
        let c = e.commit().unwrap();
        let empty = publish!(added, c, apply_shape_designer_properties);
        preserved(&added, &empty, SLIDE);
        assert_eq!(
            empty
                .shape_designer_properties(0usize, 0usize)
                .unwrap()
                .properties(),
            Some(&properties)
        );
        let mut e = empty.edit_shape_designer_properties().unwrap();
        e.remove(0usize, 0usize).unwrap();
        let c = e.commit().unwrap();
        let cleared = publish!(empty, c, apply_shape_designer_properties);
        preserved(&empty, &cleared, SLIDE);
        assert!(
            !cleared
                .shape_designer_properties(0usize, 0usize)
                .unwrap()
                .is_present()
        );
        let mut e = cleared.edit_shape_design_element().unwrap();
        e.remove(0usize, 0usize).unwrap();
        let c = e.commit().unwrap();
        let done = publish!(cleared, c, apply_shape_design_element);
        preserved(&cleared, &done, SLIDE);
        assert!(done.shape_design_element(0usize, 0usize).unwrap().is_none());
        assert!(p.to_bytes().unwrap() == source_bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_pptx_slide_designer_tags_preserve_stable_slide_bindings() {
    let file = "powerpoint-navigation-review.pptx";
    let bytes = native(file);
    let p = Package::from_bytes(&bytes).unwrap();
    let original = p.slide_designer_tags(0usize).unwrap();
    assert!(original.tags().unwrap().is_none());
    let id = original.slide_id();
    let binding = original.binding().clone();
    let mut e = p.edit_slide_designer_tags().unwrap();
    e.set(0usize, tags("slide")).unwrap();
    let c = e.commit().unwrap();
    let added = publish!(p, c, apply_slide_designer_tags);
    preserved(&p, &added, "ppt/presentation.xml");
    let snap = added.slide_designer_tags(0usize).unwrap();
    assert_eq!(snap.slide_id(), id);
    assert_eq!(snap.binding(), &binding);
    assert_eq!(snap.tags().unwrap(), Some(&tags("slide")));
    let mut e = added.edit_slide_designer_tags().unwrap();
    e.set(0usize, Tags::new()).unwrap();
    let c = e.commit().unwrap();
    let empty = publish!(added, c, apply_slide_designer_tags);
    preserved(&added, &empty, "ppt/presentation.xml");
    assert_eq!(
        empty.slide_designer_tags(0usize).unwrap().tags().unwrap(),
        Some(&Tags::new())
    );
    let mut e = empty.edit_slide_designer_tags().unwrap();
    assert!(e.remove(0usize).unwrap().is_some());
    let c = e.commit().unwrap();
    let cleared = publish!(empty, c, apply_slide_designer_tags);
    preserved(&empty, &cleared, "ppt/presentation.xml");
    assert!(
        cleared
            .slide_designer_tags(0usize)
            .unwrap()
            .tags()
            .unwrap()
            .is_none()
    );
    assert_eq!(
        cleared.slide_designer_tags(0usize).unwrap().binding(),
        &binding
    );
    assert!(native(file) == bytes);
}
