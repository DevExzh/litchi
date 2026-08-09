#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_pptx::Package;
use litchi_pptx::shape::designer::{DrawingProperties, Limits, Tag, Tags};

fn tags(name: &str, value: &str) -> Tags {
    let mut tags = Tags::new();
    tags.push(Tag::new(name, value).unwrap()).unwrap();
    tags
}

fn find_bytes(haystack: &[u8], needle: &[u8], offset: usize) -> usize {
    haystack[offset..]
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| offset + position)
        .unwrap()
}

fn swap_first_two_slide_ids(xml: &[u8]) -> Vec<u8> {
    let list = find_bytes(xml, b"<p:sldIdLst", 0);
    let first = find_bytes(xml, b"<p:sldId ", list);
    let first_tag_end = find_bytes(xml, b">", first) + 1;
    let first_end = if xml[first_tag_end - 2] == b'/' {
        first_tag_end
    } else {
        find_bytes(xml, b"</p:sldId>", first_tag_end) + b"</p:sldId>".len()
    };
    let second = find_bytes(xml, b"<p:sldId ", first_end);
    let second_tag_end = find_bytes(xml, b">", second) + 1;
    let second_end = if xml[second_tag_end - 2] == b'/' {
        second_tag_end
    } else {
        find_bytes(xml, b"</p:sldId>", second_tag_end) + b"</p:sldId>".len()
    };
    let mut reordered = Vec::with_capacity(xml.len());
    reordered.extend_from_slice(&xml[..first]);
    reordered.extend_from_slice(&xml[second..second_end]);
    reordered.extend_from_slice(&xml[first_end..second]);
    reordered.extend_from_slice(&xml[first..first_end]);
    reordered.extend_from_slice(&xml[second_end..]);
    reordered
}

#[test]
fn package_designer_facades_round_trip_and_retire_only_on_change() {
    let properties = DrawingProperties::new().with_editable(Some(true));
    let slide_tags = tags("layout", "hero");
    let mut package = Package::new().unwrap();
    let slide = package.presentation_mut().unwrap().add_slide().unwrap();
    slide
        .add_text_box("Designer", 0, 0, 1_000_000, 500_000)
        .set_designer_properties(properties.clone())
        .unwrap();
    slide.set_designer_tags(slide_tags.clone()).unwrap();
    package.to_bytes().unwrap();

    assert_eq!(
        package
            .shape_designer_properties(0usize, 0usize)
            .unwrap()
            .properties(),
        Some(&properties)
    );
    assert_eq!(
        package.slide_designer_tags(0usize).unwrap().tags().unwrap(),
        Some(&slide_tags)
    );

    package
        .put_shape_designer_properties(0usize, 0usize, properties.clone())
        .unwrap();
    package
        .put_slide_designer_tags(0usize, slide_tags.clone())
        .unwrap();
    assert!(package.presentation_mut().is_ok());

    let changed = DrawingProperties::new().with_editable(Some(false));
    package
        .put_shape_designer_properties_with_limits(
            0usize,
            0usize,
            changed.clone(),
            Limits::default(),
        )
        .unwrap();
    assert!(package.presentation_mut().is_err());
    assert_eq!(
        package
            .presentation()
            .unwrap()
            .slide(0)
            .unwrap()
            .unwrap()
            .shape_designer_properties(0usize)
            .unwrap()
            .properties(),
        Some(&changed)
    );
}

#[test]
fn reopened_facades_support_names_limits_removal_and_stable_slide_ids() {
    let mut authored = Package::new().unwrap();
    authored
        .presentation_mut()
        .unwrap()
        .add_slide()
        .unwrap()
        .add_text_box("Named shape", 0, 0, 1_000_000, 500_000);
    let bytes = authored.to_bytes().unwrap();
    let mut package = Package::from_bytes(&bytes).unwrap();
    let (slide_name, shape_name) = {
        let slide = package.presentation().unwrap().slide(0).unwrap().unwrap();
        let slide_name = slide.name().unwrap();
        let scene = slide.shapes().unwrap();
        let shape_name = scene.shape(0usize).unwrap().name().unwrap().to_owned();
        (slide_name, shape_name)
    };

    let properties = DrawingProperties::new().with_editable(Some(true));
    package
        .put_shape_designer_properties(slide_name.as_str(), shape_name.as_str(), properties.clone())
        .unwrap();
    package
        .put_slide_designer_tags(slide_name.as_str(), tags("kind", "title"))
        .unwrap();
    assert_eq!(
        package
            .shape_designer_properties_with_limits(
                slide_name.as_str(),
                shape_name.as_str(),
                Limits::default(),
            )
            .unwrap()
            .properties(),
        Some(&properties)
    );
    assert_eq!(
        package
            .slide_designer_tags_with_limits(slide_name.as_str(), Limits::default())
            .unwrap()
            .slide_id(),
        package.presentation().unwrap().slide_references().unwrap()[0].id()
    );

    assert!(
        package
            .remove_shape_designer_properties_with_limits(
                slide_name.as_str(),
                shape_name.as_str(),
                Limits::default(),
            )
            .unwrap()
            .properties()
            .is_none()
    );
    assert!(
        package
            .remove_slide_designer_tags_with_limits(slide_name.as_str(), Limits::default())
            .unwrap()
            .tags()
            .unwrap()
            .is_none()
    );
    assert!(matches!(
        package.shape_designer_properties_with_limits(
            slide_name.as_str(),
            shape_name.as_str(),
            Limits::default().with_xml_bytes(1),
        ),
        Err(litchi_pptx::Error::Limit { .. })
    ));
}

#[test]
fn slide_designer_tags_follow_stable_id_reorder_and_refuse_broken_binding() {
    let mut authored = Package::new().unwrap();
    authored.presentation_mut().unwrap().add_slide().unwrap();
    authored.presentation_mut().unwrap().add_slide().unwrap();
    let bytes = authored.to_bytes().unwrap();
    let mut package = Package::from_bytes(&bytes).unwrap();
    let expected = tags("stable", "first");
    let slide_id = package
        .put_slide_designer_tags(0usize, expected.clone())
        .unwrap()
        .slide_id();

    package
        .edit_opc(|opc| {
            let part_name = opc.main_document_part()?.partname().clone();
            let reordered = swap_first_two_slide_ids(opc.get_part(&part_name)?.blob());
            opc.get_part_mut(&part_name)?.set_blob(reordered);
            Ok(())
        })
        .unwrap();
    let moved = package.slide_designer_tags(1usize).unwrap();
    assert_eq!(moved.slide_id(), slide_id);
    assert_eq!(moved.tags().unwrap(), Some(&expected));

    let relationship_id = package.presentation().unwrap().slide_references().unwrap()[1]
        .relationship_id()
        .to_owned();
    package
        .edit_opc(|opc| {
            let part_name = opc.main_document_part()?.partname().clone();
            opc.get_part_mut(&part_name)?
                .rels_mut()
                .remove(&relationship_id);
            Ok(())
        })
        .unwrap();
    assert!(package.slide_designer_tags(1usize).is_err());
}
