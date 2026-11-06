//! Native PowerPoint InkML discovery, inert metadata and source preservation.
use litchi_opc::PackURI;
use litchi_pptx::{Package, presentation::embedded::ink};

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn annotations(package: &Package) -> Vec<ink::Annotation> {
    let p = package.presentation().unwrap();
    let slides = p.slides().unwrap();
    assert_eq!(slides.len(), 1);
    ink::load_slide(
        package.opc().unwrap(),
        0,
        slides[0].part().part(),
        &mut ink::Limits::default(),
    )
    .unwrap()
}
fn sha256(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .map(|b| format!("{b:02x}"))
        .collect()
}

#[test]
fn native_inkml_annotations_discover_the_active_compatibility_choice() {
    for name in ["powerpoint-ink-stroke.pptx", "powerpoint-ink-replay.pptx"] {
        let package = native(name);
        let found = annotations(&package);
        assert_eq!(found.len(), 1, "{name}");
        let ink = &found[0];
        assert_eq!((ink.slide_index(), ink.index()), (0, 0));
        assert_eq!(ink.relationship_id(), "rId2");
        assert_eq!(ink.part_name().as_str(), "/ppt/ink/ink1.xml");
        assert_eq!((ink.trace_count(), ink.trace_group_count()), (1, 0));
        let payload = package.opc().unwrap().get_part(ink.part_name()).unwrap();
        assert_eq!(payload.blob().len(), 2424);
        assert_eq!(payload.content_type(), ink::CONTENT_TYPE);
        assert_eq!(
            sha256(payload.blob()),
            "153cd907463545144a3cdf362afbc52a6308f4db40a49c5590cfde5233dc883c"
        );
        let preview = package
            .opc()
            .unwrap()
            .get_part(&PackURI::new("/ppt/media/image1.png").unwrap())
            .unwrap();
        assert_eq!(preview.blob().len(), 1645);
        assert_eq!(
            sha256(preview.blob()),
            "f233c87a6418eb3f73de40c9a4ed3a93a45dbb5f3d0f0a5497dafb675d9fa183"
        );
        let actions = package.ink_actions().unwrap();
        assert_eq!(actions.len(), 1);
        assert!(actions[0].anchors().is_empty());
        assert!(actions[0].source_anchors().is_empty());
        let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
        assert_eq!(annotations(&reopened), found);
        assert_eq!(
            reopened
                .opc()
                .unwrap()
                .get_part(ink.part_name())
                .unwrap()
                .blob(),
            payload.blob()
        );
    }
}

#[test]
fn native_inkml_shared_projection_keeps_brush_trace_references_and_raw_data() {
    use litchi_drawingml::ink::{self, BrushPropertyName};
    let package = native("powerpoint-ink-stroke.pptx");
    let part = package
        .opc()
        .unwrap()
        .get_part(&PackURI::new("/ppt/ink/ink1.xml").unwrap())
        .unwrap();
    let doc = ink::read(part.blob(), ink::Limits::DEFAULT).unwrap();
    assert_eq!(doc.trace_count(), 1);
    let trace = &doc.traces()[0];
    assert_eq!(trace.context_ref(), Some("#ctx0"));
    assert_eq!(trace.brush_ref(), Some("#br0"));
    assert_eq!(trace.data(&doc).len(), 1263);
    assert_eq!(
        sha256(trace.data(&doc)),
        "0e3b1078e313f32b9edbddd0aab535533bb42166a8df11a4d79e1bc90469d99c"
    );
    assert_eq!(
        doc.brush_properties()
            .iter()
            .map(|p| (p.name(), p.value(), p.units()))
            .collect::<Vec<_>>(),
        [
            (&BrushPropertyName::Width, "0.035", Some("cm")),
            (&BrushPropertyName::Height, "0.035", Some("cm")),
            (&BrushPropertyName::Color, "#E71224", None)
        ]
    );
    assert_eq!(ink::write(&doc, ink::Limits::DEFAULT).unwrap(), part.blob());
}

#[test]
fn native_ink_replay_keeps_shared_target_duration_and_source_timing() {
    use litchi_pptx::animations::{Duration, Effect, Trigger};
    let baseline = native("powerpoint-ink-stroke.pptx");
    assert!(
        baseline.presentation().unwrap().slides().unwrap()[0]
            .animations()
            .unwrap()
            .is_empty()
    );
    let package = native("powerpoint-ink-replay.pptx");
    let p = package.presentation().unwrap();
    let slides = p.slides().unwrap();
    let sequence = slides[0].animations().unwrap();
    assert_eq!(sequence.animations.len(), 1);
    let effect = &sequence.animations[0];
    assert_eq!(effect.shape_id, 4);
    assert_eq!(effect.effect, Effect::Custom("entr:63".into()));
    assert_eq!(effect.trigger, Trigger::OnClick);
    assert_eq!(effect.duration, Duration::Finite(2000));
    assert_eq!(effect.delay, 0);
    assert!(sequence.timing_tree().is_some());
    let xml = sequence.preserved_timing_xml().unwrap();
    assert!(xml.contains("drawProgress"));
    assert!(xml.contains("style.visibility"));
    let emitted = sequence
        .to_xml(litchi_pptx::animations::Limits::DEFAULT)
        .unwrap();
    assert_eq!(emitted, xml);
    let reread = litchi_pptx::animations::Sequence::parse_timing_xml(
        &emitted,
        litchi_pptx::animations::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(reread.animations, sequence.animations);
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = archive.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}
fn assert_only_changed(before: &[u8], after: &[u8], changed: &[&str]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    assert_eq!(
        a.iter()
            .filter(|(name, data)| **data != b[*name])
            .map(|(name, _)| name.as_str())
            .collect::<Vec<_>>(),
        changed
    );
}

#[test]
fn native_ink_owner_mode_and_payload_edits_keep_fallback_and_replay_bytes() {
    use litchi_pptx::presentation::embedded::content_parts::BlackWhiteMode;
    for name in ["powerpoint-ink-stroke.pptx", "powerpoint-ink-replay.pptx"] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let uri = PackURI::new("/ppt/ink/ink1.xml").unwrap();
        let original = package.opc().unwrap().get_part(&uri).unwrap().blob();
        let text = std::str::from_utf8(original).unwrap();
        assert_eq!(text.matches("#E71224").count(), 1);
        let revised = text.replace("#E71224", "#112233").into_bytes();
        // Payload replacement is inert. Its existing raster preview is kept,
        // not regenerated or presented as a rendering of the changed color.
        let parsed =
            litchi_drawingml::ink::read(&revised, litchi_drawingml::ink::Limits::DEFAULT).unwrap();
        assert_eq!(parsed.brush_properties()[2].value(), "#112233");
        let mut edit = package.edit_content_parts().unwrap();
        edit.update(0_usize, |draft| {
            assert_eq!(draft.parts().len(), 1);
            assert_eq!(draft.parts()[0].relationship_id(), "rId2");
            assert_eq!(draft.parts()[0].bw_mode(), Some(BlackWhiteMode::Auto));
            assert!(draft.set_black_white_mode(0, Some(BlackWhiteMode::Gray))?);
            assert!(draft.replace_payload(0, ink::CONTENT_TYPE, revised.clone())?);
            Ok(())
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        let output = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
        commit.package().save(output.path()).unwrap();
        let reopened = Package::open(output.path()).unwrap();
        assert_eq!(annotations(&reopened).len(), 1);
        let parts = reopened.presentation().unwrap().content_parts().unwrap();
        assert_eq!(parts.len(), 1);
        assert_eq!(parts[0].bw_mode(), Some(BlackWhiteMode::Gray));
        assert_eq!(parts[0].payload().unwrap().bytes(), revised);
        let after = reopened.to_bytes().unwrap();
        assert_only_changed(
            &before,
            &after,
            &["ppt/ink/ink1.xml", "ppt/slides/slide1.xml"],
        );
        let a = members(&before);
        let b = members(&after);
        let slide = std::str::from_utf8(&a["ppt/slides/slide1.xml"]).unwrap();
        assert_eq!(slide.matches("p14:bwMode=\"auto\"").count(), 1);
        assert!(
            b["ppt/slides/slide1.xml"]
                == slide
                    .replace("p14:bwMode=\"auto\"", "p14:bwMode=\"gray\"")
                    .as_bytes(),
            "only the source bwMode value may change"
        );
        assert_eq!(
            commit
                .package()
                .apply_content_parts(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            before
        );
        assert_eq!(package.to_bytes().unwrap(), before);
    }
}

#[test]
fn native_ink_addition_allocates_fresh_owners_and_preserves_the_original_stroke() {
    let package = native("powerpoint-ink-stroke.pptx");
    let before = package.to_bytes().unwrap();
    let a = members(&before);
    let ink_bytes = &a["ppt/ink/ink1.xml"];
    let mut edit = package.edit_ink().unwrap();
    let added = edit.add(0_usize, ink_bytes).unwrap();
    assert_eq!(added.relationship_id, "rId4");
    assert_eq!(added.part_name.as_str(), "/ppt/ink/ink2.xml");
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let found = annotations(&reopened);
    assert_eq!(found.len(), 2);
    assert_eq!(found[1].relationship_id(), "rId4");
    assert_eq!(found[1].trace_count(), 1);
    let b = members(&after);
    let added_names = b
        .keys()
        .filter(|n| !a.contains_key(*n))
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(added_names, ["ppt/ink/ink2.xml"]);
    assert_eq!(&b["ppt/ink/ink2.xml"], ink_bytes);
    for (name, data) in &a {
        if ![
            "[Content_Types].xml",
            "ppt/slides/slide1.xml",
            "ppt/slides/_rels/slide1.xml.rels",
        ]
        .contains(&name.as_str())
        {
            assert_eq!(data, &b[name], "{name}");
        }
    }
    assert_eq!(
        commit
            .package()
            .apply_ink(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}

#[test]
fn native_ink_choice_relationship_failures_are_reported_without_empty_success() {
    use litchi_pptx::Error;
    for mode in [0, 1, 2, 3, 4] {
        let package = native("powerpoint-ink-stroke.pptx");
        let mut opc = litchi_opc::OpcPackage::from_bytes(
            &package.to_bytes().unwrap(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        let slide = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let payload = PackURI::new("/ppt/ink/ink1.xml").unwrap();
        match mode {
            0 => {
                opc.get_part_mut(&slide).unwrap().rels_mut().remove("rId2");
            },
            1 => {
                opc.remove_part(&payload);
            },
            2 => {
                opc.get_part_mut(&payload)
                    .unwrap()
                    .set_blob(b"<ink xmlns=\"urn:foreign\"/>".to_vec());
            },
            _ => {
                let rels = opc.get_part_mut(&slide).unwrap().rels_mut();
                rels.remove("rId2");
                let (kind, target, external) = if mode == 3 {
                    (
                        litchi_opc::constants::relationship_type::IMAGE,
                        "../ink/ink1.xml",
                        false,
                    )
                } else {
                    (
                        litchi_opc::constants::relationship_type::CUSTOM_XML,
                        "https://example.invalid/ink.xml",
                        true,
                    )
                };
                rels.add_relationship(kind.into(), target.into(), "rId2".into(), external);
            },
        }
        let error = ink::load_slide(
            &opc,
            0,
            opc.get_part(&slide).unwrap(),
            &mut ink::Limits::default(),
        )
        .unwrap_err();
        match mode {
            0 => assert!(matches!(error, Error::Relationship(_))),
            1 => assert!(matches!(
                error,
                Error::Opc(litchi_opc::OpcError::PartNotFound(_))
            )),
            2 => assert!(matches!(error, Error::Invalid(_))),
            _ => assert!(matches!(error, Error::Relationship(_))),
        }
    }
}

#[test]
fn native_ink_removal_collects_the_orphan_payload_and_inverses_exactly() {
    // The baseline has no replay timing that would refer to the removed ink.
    let package = native("powerpoint-ink-stroke.pptx");
    let before = package.to_bytes().unwrap();
    let a = members(&before);
    let mut edit = package.edit_content_parts().unwrap();
    edit.update(0_usize, |draft| {
        let removed = draft.remove(0)?;
        assert_eq!(removed.relationship_id(), "rId2");
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert!(annotations(&reopened).is_empty());
    assert!(
        reopened
            .presentation()
            .unwrap()
            .content_parts()
            .unwrap()
            .is_empty()
    );
    let opc = reopened.opc().unwrap();
    let slide = opc
        .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap();
    assert!(slide.rels().get("rId2").is_none());
    assert!(slide.rels().get("rId3").is_some());
    assert!(!opc.contains_part(&PackURI::new("/ppt/ink/ink1.xml").unwrap()));
    let b = members(&after);
    assert_eq!(
        a.keys()
            .filter(|n| !b.contains_key(*n))
            .map(String::as_str)
            .collect::<Vec<_>>(),
        ["ppt/ink/ink1.xml"]
    );
    assert!(b.keys().all(|n| a.contains_key(n)));
    for (name, data) in &b {
        if ![
            "[Content_Types].xml",
            "ppt/slides/slide1.xml",
            "ppt/slides/_rels/slide1.xml.rels",
        ]
        .contains(&name.as_str())
        {
            assert_eq!(data, &a[name], "{name}");
        }
    }
    assert_eq!(
        commit
            .package()
            .apply_content_parts(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}
