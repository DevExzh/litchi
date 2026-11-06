//! Native placeholder ownership and source-backed SVG attachment operations.

use litchi_pptx::{Package, master_layout::PlaceholderKind, source};
use std::{collections::BTreeMap, path::PathBuf};

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint")
        .join(name)
}

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

fn changed_members(before: &[u8], after: &[u8]) -> Vec<String> {
    let a = members(before);
    let b = members(after);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|name| a.get(*name) != b.get(*name))
        .cloned()
        .collect()
}

fn element<'a>(xml: &'a [u8], tag: &str) -> &'a str {
    let xml = std::str::from_utf8(xml).unwrap();
    let start = xml.find(&format!("<{tag}")).unwrap();
    let close = format!("</{tag}>");
    let end = start + xml[start..].find(&close).unwrap() + close.len();
    &xml[start..end]
}

fn assert_media_metadata_unchanged(before: &[u8], after: &[u8], slide: usize) {
    let a = members(before);
    let b = members(after);
    let slide_name = format!("ppt/slides/slide{}.xml", slide + 1);
    for tag in ["p:nvPicPr", "p:timing"] {
        assert_eq!(
            element(&a[&slide_name], tag),
            element(&b[&slide_name], tag),
            "{tag}"
        );
    }
    let original = Package::from_bytes(before).unwrap();
    let updated = Package::from_bytes(after).unwrap();
    let uri = litchi_opc::PackURI::new(format!("/{slide_name}")).unwrap();
    let original_part = original.opc().unwrap().get_part(&uri).unwrap();
    let updated_part = updated.opc().unwrap().get_part(&uri).unwrap();
    for relationship in original_part.rels().iter() {
        let retained = updated_part.rels().get(relationship.r_id()).unwrap();
        assert_eq!(retained.reltype(), relationship.reltype());
        assert_eq!(retained.target_ref(), relationship.target_ref());
        assert_eq!(retained.is_external(), relationship.is_external());
    }
}

#[test]
fn native_placeholders_resolve_slide_layout_and_master_owners() {
    let package = Package::open(fixture("powerpoint-master-motion.pptx")).unwrap();
    let presentation = package.presentation().unwrap();
    let slide = presentation.slide(1).unwrap().unwrap();
    let layout = slide.layout().unwrap().unwrap();
    assert_eq!(
        layout.part().part().partname().as_str(),
        "/ppt/slideLayouts/slideLayout1.xml"
    );
    assert_eq!(layout.name().unwrap(), "Fixture Title Layout 006");
    let master = layout.master().unwrap();
    assert_eq!(
        master.part().part().partname().as_str(),
        "/ppt/slideMasters/slideMaster1.xml"
    );
    assert_eq!(master.layouts().unwrap().len(), 17);
    let slide_scene = slide.shapes().unwrap();
    let layout_scene = layout.shapes().unwrap();
    let master_scene = master.shapes().unwrap();
    for (kind, index, text) in [
        ("ctrTitle", 0, "PPT_MASTER_MOTION_006"),
        ("dt", 10, "2026-01-15"),
        ("ftr", 11, "PPT_FOOTER_003"),
        ("sldNum", 12, "2"),
    ] {
        let shape = slide_scene
            .placeholders()
            .find(|s| s.placeholder().unwrap().kind() == Some(kind))
            .unwrap();
        assert_eq!(shape.placeholder().unwrap().index(), index);
        assert_eq!(shape.text(), Some(text));
        assert!(shape.bounds().is_none());
        let slot = layout_scene
            .placeholders()
            .find(|s| s.placeholder().unwrap().index() == index)
            .unwrap();
        assert_eq!(slot.placeholder().unwrap().kind(), Some(kind));
    }
    let subtitle = slide_scene
        .placeholders()
        .find(|s| s.placeholder().unwrap().kind() == Some("subTitle"))
        .unwrap();
    assert_eq!(subtitle.placeholder().unwrap().index(), 1);
    assert!(
        subtitle
            .text()
            .unwrap()
            .contains("Native master, layout, animation and Zoom")
    );
    let layout_subtitle = layout_scene
        .placeholders()
        .find(|s| s.placeholder().unwrap().index() == 1)
        .unwrap();
    assert_eq!(
        layout_subtitle.placeholder().unwrap().kind(),
        Some("subTitle")
    );
    let bounds = layout_subtitle.bounds().unwrap();
    assert_eq!(
        (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
        (1_154_955, 4_777_380, 8_825_658, 861_420)
    );
    // Master indices differ from the layout's; these are declared owners,
    // not a synthesized rendering or a flattened effective shape scene.
    for (kind, index) in [
        ("title", 0),
        ("body", 1),
        ("dt", 2),
        ("ftr", 3),
        ("sldNum", 4),
    ] {
        let shape = master_scene
            .placeholders()
            .find(|s| s.placeholder().unwrap().kind() == Some(kind))
            .unwrap();
        assert_eq!(shape.placeholder().unwrap().index(), index);
        assert!(shape.bounds().is_some());
    }
    for (scene, text, id, y) in [
        (&layout_scene, "TITLE_LAYOUT_INHERITED_006", 9, 5_943_600),
        (&master_scene, "MASTER_INHERITED_006", 13, 6_355_080),
    ] {
        let marker = scene.iter().find(|s| s.text() == Some(text)).unwrap();
        assert_eq!(marker.id(), Some(id));
        assert!(marker.placeholder().is_none());
        let bounds = marker.bounds().unwrap();
        assert_eq!(
            (bounds.x(), bounds.y(), bounds.width(), bounds.height()),
            (548_640, y, 6_099_858, 369_332)
        );
        assert!(!slide_scene.iter().any(|s| s.text() == Some(text)));
    }
}

#[test]
fn native_layout_retarget_preserves_shapes_and_other_package_members() {
    let package = Package::open(fixture("powerpoint-master-motion.pptx")).unwrap();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    assert!(
        edit.set_slide_layout(6_usize, "Fixture Title Layout 006")
            .unwrap()
    );
    assert!(
        !edit
            .set_slide_layout(6_usize, "Fixture Title Layout 006")
            .unwrap()
    );
    assert!(matches!(
        edit.set_slide_layout(6_usize, "missing native layout"),
        Err(litchi_pptx::Error::SlideLayoutNameNotFound(_))
    ));
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    assert_eq!(
        changed_members(&before, &after),
        ["ppt/slides/_rels/slide7.xml.rels"]
    );
    let output = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    commit.package().save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let presentation = reopened.presentation().unwrap();
    let slide = presentation.slide(6).unwrap().unwrap();
    assert_eq!(
        slide.layout().unwrap().unwrap().name().unwrap(),
        "Fixture Title Layout 006"
    );
    assert_eq!(slide.shapes().unwrap().len(), 5);
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(
        members(&restored.package().to_bytes().unwrap()),
        members(&before)
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn native_placeholder_recreation_keeps_title_and_subtitle_layout_indices() {
    let package = Package::open(fixture("powerpoint-master-motion.pptx")).unwrap();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit().unwrap();
    assert!(
        edit.set_slide_layout(6_usize, "Fixture Title Layout 006")
            .unwrap()
    );
    assert!(edit.remove_shape(6_usize, "Title 1").unwrap());
    assert!(edit.remove_shape(6_usize, "Content Placeholder 2").unwrap());
    edit.add_placeholder(6_usize, PlaceholderKind::Title, "Replacement native title")
        .unwrap();
    edit.add_placeholder(
        6_usize,
        PlaceholderKind::Subtitle,
        "Replacement native subtitle",
    )
    .unwrap();
    edit.put_placeholder_text(
        6_usize,
        PlaceholderKind::Title,
        None,
        "Updated native title",
    )
    .unwrap();
    assert!(matches!(
        edit.add_placeholder(6_usize, PlaceholderKind::Title, "duplicate"),
        Err(litchi_pptx::Error::Refused(_))
    ));
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    assert_eq!(
        changed_members(&before, &after),
        ["ppt/slides/_rels/slide7.xml.rels", "ppt/slides/slide7.xml"]
    );
    let output = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    commit.package().save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let presentation = reopened.presentation().unwrap();
    let slide = presentation.slide(6).unwrap().unwrap();
    let scene = slide.shapes().unwrap();
    assert_eq!(scene.len(), 5);
    for (kind, index, text) in [
        ("ctrTitle", 0, "Updated native title"),
        ("subTitle", 1, "Replacement native subtitle"),
    ] {
        let shape = scene
            .placeholders()
            .find(|s| s.placeholder().unwrap().kind() == Some(kind))
            .unwrap();
        assert_eq!(shape.placeholder().unwrap().index(), index);
        assert_eq!(shape.text(), Some(text));
        assert!(shape.bounds().is_none());
    }
    for text in ["2026-01-15", "PPT_FOOTER_003", "7"] {
        assert!(scene.iter().any(|s| s.text() == Some(text)));
    }
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(
        members(&restored.package().to_bytes().unwrap()),
        members(&before)
    );
}

#[test]
fn native_png_media_pictures_attach_replace_and_detach_svg_without_changing_playback() {
    let svg = include_bytes!("../../../test-data/microsoft-office/assets/svg-blue-yellow-007.svg");
    for slide in [1_usize, 2] {
        let package = source::Package::open(fixture("powerpoint-media-objects.pptx")).unwrap();
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit_svg_images().unwrap();
        let initial = edit.attachment(slide, 0).unwrap();
        assert!(initial.svg().is_none());
        assert_eq!(
            initial.raster_part_uri().as_str(),
            format!("/ppt/media/image{slide}.png")
        );
        let raster_id = initial.raster_relationship_id().to_owned();
        assert!(!edit.detach(slide, 0).unwrap());
        assert!(
            edit.attach(
                slide,
                0,
                &source::svg::AttachmentReplacement::new(svg.to_vec())
            )
            .unwrap()
        );
        let attached = edit.attachment(slide, 0).unwrap();
        assert_eq!(attached.raster_relationship_id(), raster_id);
        let vector = attached.svg().unwrap();
        assert_eq!(vector.bytes(), svg);
        let svg_member = vector.part_uri().membername().to_owned();
        let svg_id = vector.relationship_id().to_owned();
        assert_ne!(svg_id, raster_id);
        let (updated, patch) = edit.commit().unwrap().into_parts();
        let after = updated.to_bytes().unwrap();
        let mut expected = vec![
            "[Content_Types].xml".to_owned(),
            format!("ppt/slides/_rels/slide{}.xml.rels", slide + 1),
            format!("ppt/slides/slide{}.xml", slide + 1),
            svg_member.clone(),
        ];
        expected.sort();
        assert_eq!(changed_members(&before, &after), expected);
        assert_eq!(members(&after)[&svg_member], svg);
        assert_media_metadata_unchanged(&before, &after, slide);
        let output = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
        updated.save(output.path()).unwrap();
        let reopened = source::Package::open(output.path()).unwrap();
        let mut replace = reopened.edit_svg_images().unwrap();
        let snapshot = replace.attachment(slide, 0).unwrap();
        assert_eq!(snapshot.svg().unwrap().relationship_id(), svg_id);
        assert_eq!(snapshot.svg().unwrap().bytes(), svg);
        let replacement = svg
            .iter()
            .copied()
            .chain(b"\n<!-- native-source replacement -->".iter().copied())
            .collect::<Vec<_>>();
        let raster = members(&before)[initial.raster_part_uri().membername()].clone();
        assert!(
            replace
                .replace(
                    slide,
                    0,
                    source::svg::Replacement::new(replacement.clone(), raster)
                )
                .unwrap()
        );
        let replaced = replace.commit().unwrap().into_package();
        let replaced_bytes = replaced.to_bytes().unwrap();
        assert_eq!(
            changed_members(&after, &replaced_bytes),
            std::slice::from_ref(&svg_member)
        );
        assert_eq!(
            replaced
                .edit_svg_images()
                .unwrap()
                .attachment(slide, 0)
                .unwrap()
                .svg()
                .unwrap()
                .bytes(),
            replacement
        );
        let mut detach = replaced.edit_svg_images().unwrap();
        assert!(detach.detach(slide, 0).unwrap());
        let detached = detach.commit().unwrap().into_package();
        assert!(
            detached
                .edit_svg_images()
                .unwrap()
                .attachment(slide, 0)
                .unwrap()
                .svg()
                .is_none()
        );
        let detached_members = members(&detached.to_bytes().unwrap());
        assert_media_metadata_unchanged(&before, &detached.to_bytes().unwrap(), slide);
        assert!(!detached_members.contains_key(&svg_member));
        // The writer may retain an unused SVG content-type declaration.
        for (name, bytes) in members(&before) {
            if name != "[Content_Types].xml"
                && name != format!("ppt/slides/_rels/slide{}.xml.rels", slide + 1)
                && name != format!("ppt/slides/slide{}.xml", slide + 1)
            {
                assert_eq!(detached_members[&name], bytes, "{name}");
            }
        }
        // This inverse is authorized against the retained committed source.
        assert_eq!(
            updated
                .apply_svg_images(&patch.inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            before
        );
        assert_eq!(package.to_bytes().unwrap(), before);
    }
}

#[test]
fn native_svg_attachment_rejects_non_png_fallbacks_without_staging_changes() {
    for (name, slide, actual) in [
        ("powerpoint-modern-media.pptx", 4_usize, "image/svg+xml"),
        ("powerpoint-visual-objects.pptx", 1_usize, "image/jpeg"),
    ] {
        let package = source::Package::open(fixture(name)).unwrap();
        let before = package.to_bytes().unwrap();
        let mut edit = package.edit_svg_images().unwrap();
        assert!(
            matches!(edit.attach(slide, 0, &source::svg::AttachmentReplacement::new(b"<svg/>".to_vec())), Err(litchi_pptx::Error::ContentType { expected, actual: found }) if expected == "image/png" && found == actual)
        );
        let commit = edit.commit().unwrap();
        assert!(commit.patch().is_noop());
        assert_eq!(commit.package().to_bytes().unwrap(), before);
    }
}

#[test]
fn native_svg_import_rejects_malformed_or_dtd_payloads_before_output() {
    let package = source::Package::open(fixture("powerpoint-media-objects.pptx")).unwrap();
    let before = package.to_bytes().unwrap();
    for invalid in [
        b"<svg><g></svg>".as_slice(),
        b"<!DOCTYPE svg [<!ENTITY x 'hidden'>]><svg>&x;</svg>",
        b"<svg/><svg/>",
    ] {
        let mut edit = package.edit_svg_images().unwrap();
        edit.attach(
            1,
            0,
            &source::svg::AttachmentReplacement::new(invalid.to_vec()),
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let result = commit.package().write_to(&mut output);
        assert!(
            matches!(
                result,
                Err(litchi_pptx::Error::Opc(
                    litchi_opc::OpcError::XmlPublication { .. }
                ))
            ),
            "{result:?}"
        );
        assert!(output.is_empty());
        assert_eq!(package.to_bytes().unwrap(), before);
    }
}

#[test]
fn native_svg_replacement_can_import_a_new_media_part_without_normalization() {
    let package = source::Package::open(fixture("powerpoint-media-objects.pptx")).unwrap();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_svg_images().unwrap();
    edit.attach(
        1,
        0,
        &source::svg::AttachmentReplacement::new(b"<svg/>".to_vec()),
    )
    .unwrap();
    let attached = edit.commit().unwrap().into_package();
    let mut edit = attached.edit_svg_images().unwrap();
    let snapshot = edit.attachment(1, 0).unwrap();
    let old = snapshot.svg().unwrap().part_uri().membername().to_owned();
    let id = snapshot.svg().unwrap().relationship_id().to_owned();
    let raster = members(&before)[snapshot.raster_part_uri().membername()].clone();
    let svg = include_bytes!("../../../test-data/microsoft-office/assets/svg-blue-yellow-007.svg");
    edit.replace(
        1,
        0,
        source::svg::Replacement::new(svg.to_vec(), raster).with_part_uris(
            litchi_opc::PackURI::new("/ppt/media/native-replacement.svg").unwrap(),
            snapshot.raster_part_uri().clone(),
        ),
    )
    .unwrap();
    let (updated, patch) = edit.commit().unwrap().into_parts();
    let after = updated.to_bytes().unwrap();
    let attached_bytes = attached.to_bytes().unwrap();
    let mut changed = vec![
        "[Content_Types].xml".to_owned(),
        "ppt/slides/_rels/slide2.xml.rels".to_owned(),
        "ppt/media/native-replacement.svg".to_owned(),
        old.clone(),
    ];
    changed.sort();
    assert_eq!(changed_members(&attached_bytes, &after), changed);
    let data = members(&after);
    assert!(!data.contains_key(&old));
    assert_eq!(data["ppt/media/native-replacement.svg"], svg);
    assert_eq!(
        data["ppt/media/image1.png"],
        members(&before)["ppt/media/image1.png"]
    );
    let state = updated.edit_svg_images().unwrap().attachment(1, 0).unwrap();
    assert_eq!(state.svg().unwrap().relationship_id(), id);
    assert_eq!(state.svg().unwrap().bytes(), svg);
    assert_eq!(
        updated
            .apply_svg_images(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        attached_bytes
    );
}
