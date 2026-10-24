#![allow(clippy::unwrap_used, reason = "test assertions")]

use litchi_pptx::{Package, ShapeTextReplacement, builder::Builder};

#[test]
fn ten_thousand_text_boxes_admit_an_eager_native_edit() {
    let mut builder = Builder::new().unwrap();
    let slide = builder.presentation().unwrap().add_slide().unwrap();
    for index in 0..10_000 {
        slide.add_text_box(&format!("shape-{index}"), 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(edit.slides().unwrap().len(), 1);
    assert_eq!(
        edit.set_shape_texts(
            0_usize,
            &[
                ShapeTextReplacement::at(0, "first changed"),
                ShapeTextReplacement::at(9_999, "last changed"),
            ]
        )
        .unwrap(),
        2
    );
    let after = edit.commit().unwrap().into_package();
    assert_eq!(source.to_bytes().unwrap(), bytes);
    let reopened = Package::from_bytes(&after.to_bytes().unwrap()).unwrap();
    let presentation = reopened.presentation().unwrap();
    let slide = presentation.slide(0).unwrap().unwrap();
    let scene = slide.shapes().unwrap();
    assert_eq!(scene.len(), 10_000);
    assert_eq!(scene.at(0).unwrap().text(), Some("first changed"));
    assert_eq!(scene.at(9_999).unwrap().text(), Some("last changed"));
}
