//! Native source-backed image slide copies; no rendering or chart regeneration.
use litchi_pptx::{Error, Package, SourceBackedPresentation, source};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn package(bytes: &[u8]) -> source::Package {
    source::Package::from_read_at(std::sync::Arc::new(litchi_core::OwnedSource::new(
        bytes.to_vec(),
    )))
    .unwrap()
}
fn bytes(p: &source::Package) -> Vec<u8> {
    let mut out = Vec::new();
    p.write_to(&mut out).unwrap();
    out
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let a = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    a.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = a.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn without_first_slide_ref(xml: &[u8]) -> Vec<u8> {
    use quick_xml::events::Event;
    let mut r = quick_xml::Reader::from_reader(xml);
    loop {
        let start = usize::try_from(r.buffer_position()).unwrap();
        match r.read_event().unwrap() {
            Event::Empty(e) if e.local_name().as_ref() == b"sldId" => {
                let end = usize::try_from(r.buffer_position()).unwrap();
                return [&xml[..start], &xml[end..]].concat();
            },
            Event::Eof => panic!("missing copied slide reference"),
            _ => {},
        }
    }
}
fn manifest(xml: &[u8]) -> BTreeMap<String, BTreeMap<String, String>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut out = BTreeMap::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e)
                if matches!(
                    e.local_name().as_ref(),
                    b"Relationship" | b"Default" | b"Override"
                ) =>
            {
                let attrs = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect::<BTreeMap<_, _>>();
                let key = attrs
                    .get("Id")
                    .or_else(|| attrs.get("PartName"))
                    .or_else(|| attrs.get("Extension"))
                    .unwrap()
                    .clone();
                assert!(out.insert(key, attrs).is_none());
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out
}
#[test]
fn native_picture_slides_copy_with_metadata_media_and_exact_member_preservation() {
    for ext in ["pptx", "ppsx", "potx"] {
        let original = native(&format!("powerpoint-visual-objects.{ext}"));
        let source = SourceBackedPresentation::from_reader(Cursor::new(&original)).unwrap();
        let target = package(&original);
        let before = members(&original);
        let jpeg = &before["ppt/media/image1.jpg"];
        assert_eq!(jpeg.len(), 83204);
        assert_eq!(
            Sha256::digest(jpeg)
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect::<String>(),
            "b6f8d168d41e0fc296eb1e0e3855d9669584b7528c2ef4966817443d9bc9b929"
        );
        assert!(
            bytes(
                target
                    .edit_cross_slide_copy()
                    .unwrap()
                    .commit()
                    .unwrap()
                    .package()
            ) == original
        );
        let plan = target
            .plan_cross_slide_copy(&source, 1usize, 1usize, 0usize)
            .unwrap();
        let mut edit = target.edit_cross_slide_copy().unwrap();
        let copied = edit.stage(&plan).unwrap();
        assert_eq!(
            (
                copied.source_position(),
                copied.destination_slide_position(),
                copied.insertion_position(),
                copied.destination_slide_count()
            ),
            (1, 1, 0, 6)
        );
        assert_eq!(copied.name(), "");
        let commit = edit.commit().unwrap();
        let out = bytes(commit.package());
        assert!(bytes(commit.package()) == out);
        let file = tempfile::NamedTempFile::with_suffix(format!(".{ext}")).unwrap();
        commit.package().save(file.path()).unwrap();
        assert!(std::fs::read(file.path()).unwrap() == out);
        let after = members(&out);
        assert_eq!(
            after
                .keys()
                .filter(|k| !before.contains_key(*k))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            [
                "ppt/media/image1-copy1.jpg",
                "ppt/slides/_rels/slide6.xml.rels",
                "ppt/slides/slide6.xml"
            ]
        );
        assert!(after["ppt/media/image1-copy1.jpg"] == *jpeg);
        assert!(after["ppt/slides/slide6.xml"] == before["ppt/slides/slide2.xml"]);
        assert!(
            without_first_slide_ref(&after["ppt/presentation.xml"])
                == before["ppt/presentation.xml"]
        );
        for (name, data) in &before {
            if !matches!(
                name.as_str(),
                "[Content_Types].xml" | "ppt/presentation.xml" | "ppt/_rels/presentation.xml.rels"
            ) {
                assert!(after[name] == *data, "{ext}: changed {name}");
            }
        }
        for part in ["[Content_Types].xml", "ppt/_rels/presentation.xml.rels"] {
            let old = manifest(&before[part]);
            let new = manifest(&after[part]);
            for (id, entry) in &old {
                assert_eq!(
                    new.get(id),
                    Some(entry),
                    "changed existing manifest entry {part}: {id}"
                );
            }
            let added = new
                .iter()
                .filter(|(id, _)| !old.contains_key(*id))
                .collect::<Vec<_>>();
            if part == "[Content_Types].xml" {
                assert!(added.iter().all(|(id, _)| matches!(
                    id.as_str(),
                    "/ppt/slides/slide6.xml" | "/ppt/media/image1-copy1.jpg"
                )));
                assert!(new.contains_key("/ppt/slides/slide6.xml"));
            } else {
                assert_eq!(added.len(), 1);
                assert_eq!(
                    added[0].1["Type"],
                    litchi_opc::constants::relationship_type::SLIDE
                );
                assert_eq!(added[0].1["Target"], "slides/slide6.xml");
            }
        }
        let read = Package::from_reader(Cursor::new(&out)).unwrap();
        let p = read.presentation().unwrap();
        let slides = p.slides().unwrap();
        assert_eq!(slides.len(), 6);
        assert_eq!(slides[0].text().unwrap(), "EMBEDDED_PICTURE_002");
        let states = read.edit_change_tracking().unwrap();
        let state = states.change_tracking(0usize).unwrap();
        assert_eq!(state.state().creation_id().unwrap().value(), 2058337588);
        let rels = read
            .opc()
            .unwrap()
            .get_part(&litchi_opc::PackURI::new("/ppt/slides/slide6.xml").unwrap())
            .unwrap();
        assert_eq!(rels.rels().len(), 2);
        let image = rels.rels().get("rId2").unwrap();
        assert_eq!(
            image.reltype(),
            litchi_opc::constants::relationship_type::IMAGE
        );
        assert!(!image.is_external());
        assert_eq!(image.target_ref(), "../media/image1-copy1.jpg");
        let fresh = package(&out);
        assert!(
            bytes(
                fresh
                    .apply_cross_slide_copy(&commit.patch().inverse())
                    .unwrap()
                    .package()
            ) == original
        );
        assert!(
            bytes(
                target
                    .apply_cross_slide_copy(commit.patch())
                    .unwrap()
                    .package()
            ) == out
        );
        assert!(bytes(&target) == original);
        assert!(bytes(&source::Package::from_presentation(&source).unwrap()) == original);
    }
}
#[test]
fn native_copy_rejects_dependent_graphs_without_discarding_a_prior_copy() {
    let original = native("powerpoint-visual-objects.pptx");
    let source = SourceBackedPresentation::from_reader(Cursor::new(&original)).unwrap();
    let target = package(&original);
    let mut valid = target.edit_cross_slide_copy().unwrap();
    valid.copy(&source, 1usize, 1usize, 0usize).unwrap();
    let expected = bytes(valid.commit().unwrap().package());
    let mut edit = target.edit_cross_slide_copy().unwrap();
    edit.copy(&source, 1usize, 1usize, 0usize).unwrap();
    // The native chart owns style/color parts and an embedded workbook.
    let error = edit.copy(&source, 3usize, 2usize, 0usize).unwrap_err();
    assert!(
        matches!(error,Error::Refused(r) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && r.reason()=="the selected slide cannot be copied losslessly because it has an unsupported relationship")
    );
    let after = bytes(edit.commit().unwrap().package());
    assert!(after == expected);
    assert!(bytes(&target) == original);
    for (from, to, slide, layout, reason) in [
        (
            "powerpoint-modern-media.pptx",
            "powerpoint-modern-media.pptx",
            4,
            4,
            "the selected slide cannot be copied losslessly because its semantic surface is unsupported",
        ),
        (
            "powerpoint-visual-objects.pptx",
            "powerpoint-visual-objects.ppsx",
            1,
            1,
            "the selected slide cannot be copied losslessly because it has a shared owner",
        ),
        (
            "powerpoint-visual-objects.pptx",
            "powerpoint-text-table.pptx",
            1,
            1,
            "the selected slide cannot be copied losslessly because it contains markup-compatibility content",
        ),
    ] {
        let input = native(from);
        let src = SourceBackedPresentation::from_reader(Cursor::new(&input)).unwrap();
        let destination = native(to);
        let p = package(&destination);
        let mut edit = p.edit_cross_slide_copy().unwrap();
        assert!(
            matches!(edit.copy(&src,slide,layout,0usize),Err(Error::Refused(r)) if r.kind()==litchi_core::edit::RefusalKind::UnsupportedContent && r.reason()==reason)
        );
        assert!(bytes(edit.commit().unwrap().package()) == destination);
        assert!(bytes(&p) == destination);
    }
}
#[test]
fn native_creation_and_modification_ids_retain_their_distinct_owners() {
    for ext in ["pptx", "ppsx", "potx"] {
        let original = native(&format!("powerpoint-visual-objects.{ext}"));
        let p = Package::from_reader(Cursor::new(&original)).unwrap();
        let edit = p.edit_change_tracking().unwrap();
        for (index, creation, modification) in [
            (0, 3852577993, None),
            (1, 2058337588, None),
            (2, 4165099922, Some(1341187905)),
            (3, 199848644, Some(4245603230)),
            (4, 425733195, None),
        ] {
            let snapshot = edit.change_tracking(index).unwrap();
            let state = snapshot.state();
            assert_eq!(state.creation_id().unwrap().value(), creation);
            assert_eq!(
                state
                    .shapes()
                    .iter()
                    .filter_map(|s| s.modification_id().map(|id| id.value()))
                    .collect::<Vec<_>>(),
                modification.into_iter().collect::<Vec<_>>()
            );
            assert_eq!(state.shapes()[0].name(), Some("Title 1"));
        }
        assert!(edit.commit().unwrap().package().to_bytes().unwrap() == original);
    }
}
#[test]
fn native_identifier_edits_preserve_the_chart_graph_and_unknown_creation_guids() {
    let original = native("powerpoint-visual-objects.pptx");
    let p = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = p.edit_change_tracking().unwrap();
    edit.set_creation_id(3usize, 4000000000u32).unwrap();
    edit.set_shape_modification_id(3usize, "Content Placeholder 3", 4000000001u32)
        .unwrap();
    assert!(matches!(
        edit.set_shape_modification_id(3usize, "Title 1", 4000000001u32),
        Err(Error::Invalid(_))
    ));
    let c = edit.commit().unwrap();
    let file = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    c.package().save(file.path()).unwrap();
    let after = std::fs::read(file.path()).unwrap();
    let a = members(&original);
    let b = members(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in &a {
        if name == "ppt/slides/slide4.xml" {
            let xml = std::str::from_utf8(data).unwrap();
            assert_eq!(xml.matches("val=\"199848644\"").count(), 1);
            assert_eq!(xml.matches("val=\"4245603230\"").count(), 1);
            assert!(
                b[name]
                    == xml
                        .replace("val=\"199848644\"", "val=\"4000000000\"")
                        .replace("val=\"4245603230\"", "val=\"4000000001\"")
                        .as_bytes()
            );
        } else {
            assert!(b[name] == *data, "changed {name}");
        }
    }
    let read = Package::from_reader(Cursor::new(&after)).unwrap();
    let state = read
        .edit_change_tracking()
        .unwrap()
        .change_tracking(3usize)
        .unwrap();
    assert_eq!(state.state().creation_id().unwrap().value(), 4000000000);
    assert_eq!(
        state.state().shapes()[1].modification_id().unwrap().value(),
        4000000001
    );
    let restored = read
        .apply_change_tracking(&c.patch().inverse())
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    let restored = members(&restored);
    for (n, data) in a {
        assert!(restored[&n] == data);
    }
    assert!(p.to_bytes().unwrap() == original);
}
