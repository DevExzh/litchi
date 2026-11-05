#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use litchi_pptx::{
    Package,
    shape::{Scene, owned},
};

fn parity(borrowed: &Scene<'_>, owned: &owned::Scene) {
    assert_eq!(borrowed.xml(), owned::raw::xml(owned));
    assert_eq!(borrowed.len(), owned.len());
    for (left, handle) in borrowed.iter().zip(owned.iter()) {
        let right = handle.view();
        assert_eq!(left.name(), right.name());
        assert_eq!(left.id(), right.id());
        assert_eq!(left.text(), right.text());
        assert_eq!(left.bounds(), right.bounds());
        assert_eq!(left.placeholder(), right.placeholder());
        assert_eq!(
            left.non_visual().unwrap().into_parts(),
            right.non_visual().unwrap().into_parts()
        );
        assert_eq!(left.transform().unwrap(), right.transform().unwrap());
        match (left.text_body(), right.text_body()) {
            (Ok(left), Ok(right)) => assert_eq!(left, right),
            (Err(left), Err(right)) => {
                assert_eq!(
                    std::mem::discriminant(&left),
                    std::mem::discriminant(&right)
                );
                assert_eq!(left.to_string(), right.to_string());
            },
            (left, right) => panic!("body parity mismatch: {left:?} / {right:?}"),
        }
        let payload = |shape| {
            litchi_pptx::shape::raw::payload(shape)
                .unwrap()
                .map(|value| {
                    (
                        value.attribute(),
                        value.relationship_id().unwrap().into_owned(),
                    )
                })
        };
        assert_eq!(payload(left), payload(right));
        assert_eq!(left.xml().unwrap(), right.xml().unwrap());
    }
    assert_eq!(
        borrowed
            .roots()
            .map(|s| s.common().index())
            .collect::<Vec<_>>(),
        owned
            .roots()
            .map(|s| s.position().get())
            .collect::<Vec<_>>()
    );
}
#[test]
fn slide_layout_and_master_scenes_match_and_outlive_the_package() {
    let mut retained = Vec::new();
    let mut retained_owners = Vec::new();
    for relative in [
        "test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx",
        "test-data/poi/test-data/slideshow/bug62513.pptx",
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative);
        let package = Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();
        for slide in presentation.slides().unwrap() {
            let scene = slide.scene().unwrap();
            parity(&slide.shapes().unwrap(), &scene);
            if !scene.is_rewritten() {
                assert_eq!(
                    owned::raw::xml(&scene).as_ptr(),
                    slide.part().part().blob().as_ptr()
                );
            }
            retained.extend(scene.iter().map(|handle| {
                let text = handle.view().xml().unwrap().to_vec();
                (handle, text)
            }));
            if let Some(layout) = slide.layout().unwrap() {
                let scene = layout.scene().unwrap();
                parity(&layout.shapes().unwrap(), &scene);
                retained_owners.push((scene.clone(), owned::raw::xml(&scene).to_vec()));
                retained.extend(scene.iter().map(|handle| {
                    let bytes = handle.view().xml().unwrap().to_vec();
                    (handle, bytes)
                }));
                let master = layout.master().unwrap();
                let scene = master.scene().unwrap();
                parity(&master.shapes().unwrap(), &scene);
                retained_owners.push((scene.clone(), owned::raw::xml(&scene).to_vec()));
                retained.extend(scene.iter().map(|handle| {
                    let bytes = handle.view().xml().unwrap().to_vec();
                    (handle, bytes)
                }));
            }
        }
    }
    assert!(retained.len() > 20);
    assert!(retained_owners.len() >= 2);
    for (scene, bytes) in retained_owners {
        assert_eq!(owned::raw::xml(&scene), bytes);
    }
    for (handle, text) in retained {
        assert_eq!(handle.view().xml().unwrap(), text);
    }
}

#[test]
fn source_scenes_match_eager_scenes_on_real_slides_and_survive_owner_drop() {
    let mut held = Vec::new();
    let mut compared = 0;
    for relative in [
        "test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf103477.pptx",
        "test-data/poi/test-data/slideshow/bug62513.pptx",
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(relative);
        let package = Package::open(&path).unwrap();
        let source = litchi_pptx::SourceBackedPresentation::open(&path).unwrap();
        let presentation = package.presentation().unwrap();
        for (position, slide) in presentation.slides().unwrap().iter().enumerate() {
            let scene = source.slide(position).unwrap().scene().unwrap();
            parity(&slide.shapes().unwrap(), &scene);
            compared += scene.len();
            held.push((scene.clone(), owned::raw::xml(&scene).to_vec()));
        }
    }
    assert!(compared > 20);
    for (scene, bytes) in held {
        assert_eq!(owned::raw::xml(&scene), bytes);
    }
}
