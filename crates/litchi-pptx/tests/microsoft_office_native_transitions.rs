//! Native transition CRUD must leave unrelated Zoom MCE owners byte exact.
use litchi_core::Position;
use litchi_pptx::{
    Package,
    transition::{self, Kind, Morph, Ms, Side, Speed, Transition},
};
use std::{collections::BTreeMap, sync::Arc};

fn native_bytes() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint/powerpoint-master-motion.pptx"),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    archive
        .member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = archive.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn value(package: &Package, index: usize) -> Option<Transition> {
    let p = package.presentation().unwrap();
    let slides = p.slides().unwrap();
    transition::read(
        slides[index].part().part().blob(),
        transition::Limits::DEFAULT,
    )
    .unwrap()
}
fn desired(index: usize) -> Transition {
    Transition::new(match index {
        0 => Kind::Fade { black: None },
        1 => Kind::Push(Side::Left),
        _ => Kind::Morph(Morph::ByWord),
    })
    .with_duration(Ms::new(1750).unwrap())
    .with_after(Ms::new(4500).unwrap())
    .with_click(false)
    .with_speed(Speed::Fast)
}
fn assert_preserved(before: &[u8], after: &[u8], index: usize) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let selected = format!("ppt/slides/slide{}.xml", index + 1);
    for (name, bytes) in &a {
        if name != &selected {
            assert!(bytes == &b[name], "changed unrelated member {name}");
        }
    }
    let source = std::str::from_utf8(&a[&selected]).unwrap();
    let changed = std::str::from_utf8(&b[&selected]).unwrap();
    let start = source.find("<p:cSld").unwrap();
    let end = source.find("</p:cSld>").unwrap() + "</p:cSld>".len();
    assert!(
        changed.contains(&source[start..end]),
        "native Zoom and shape XML changed"
    );
    if let Some(start) = source.find("<p:timing") {
        let end = source.find("</p:timing>").unwrap() + "</p:timing>".len();
        assert!(
            changed.contains(&source[start..end]),
            "animation timing changed"
        );
    }
    // The native Morph owner follows cSld; its fallback is separate from the
    // Zoom fallbacks inside cSld and must survive every transition edit.
    let tail = source.split_once("</p:cSld>").unwrap().1;
    if let Some(start) = tail.find("<mc:Fallback") {
        let end = tail[start..].find("</mc:Fallback>").unwrap() + start + "</mc:Fallback>".len();
        assert!(
            changed.contains(&tail[start..end]),
            "inactive transition fallback changed"
        );
    }
    let before = Package::from_bytes(before).unwrap();
    let after = Package::from_bytes(after).unwrap();
    let p = before.presentation().unwrap();
    let q = after.presentation().unwrap();
    let a = p.slides().unwrap();
    let b = q.slides().unwrap();
    let a = a[index].zooms().unwrap();
    let b = b[index].zooms().unwrap();
    assert!(a.iter().eq(b.iter()), "native Zoom values changed");
}

#[test]
fn native_eager_transition_crud_preserves_zoom_shapes_timing_and_other_members() {
    let bytes = native_bytes();
    let source = Package::from_bytes(&bytes).unwrap();
    let encoded = source.to_bytes().unwrap();
    assert!(
        members(&encoded) == members(&bytes),
        "eager open must keep native member bytes"
    );
    let bytes = encoded;
    for index in 0..3 {
        let original = value(&source, index);
        let mut noop = source.edit().unwrap();
        assert!(!noop.set_slide_transition(index, original.as_ref()).unwrap());
        assert!(
            noop.commit().unwrap().package().to_bytes().unwrap() == bytes,
            "eager no-op bytes"
        );
        let target = desired(index);
        let mut edit = source.edit().unwrap();
        assert!(edit.set_slide_transition(index, Some(&target)).unwrap());
        let commit = edit.commit().unwrap();
        let after = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert!(value(&reopened, index).unwrap().same_semantics(&target));
        assert_preserved(&bytes, &after, index);
        let mut clear = reopened.edit().unwrap();
        assert!(clear.set_slide_transition(index, None).unwrap());
        let cleared = clear.commit().unwrap();
        assert!(value(cleared.package(), index).is_none());
        assert_preserved(&bytes, &cleared.package().to_bytes().unwrap(), index);
        let cleared_reopened = Package::from_bytes(&cleared.package().to_bytes().unwrap()).unwrap();
        let mut restore = cleared_reopened.edit().unwrap();
        assert!(restore.set_slide_transition(index, Some(&target)).unwrap());
        let restored = restore.commit().unwrap();
        assert!(
            value(restored.package(), index)
                .unwrap()
                .same_semantics(&target)
        );
        assert!(
            commit
                .package()
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == bytes,
            "inverse source bytes"
        );
    }
    assert!(source.to_bytes().unwrap() == bytes, "source bytes changed");
}

#[test]
fn native_source_transition_edits_match_eager_semantics_and_inverse() {
    let bytes = native_bytes();
    let source = litchi_pptx::source::Package::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(bytes.clone()),
    ))
    .unwrap();
    for index in 0..3 {
        let target = desired(index);
        let direct = source
            .presentation()
            .unwrap()
            .slide(index)
            .unwrap()
            .transition()
            .unwrap();
        let eager_original = value(&Package::from_bytes(&bytes).unwrap(), index);
        assert!(match (&direct, &eager_original) {
            (None, None) => true,
            (Some(a), Some(b)) => a.same_semantics(b),
            _ => false,
        });
        let mut edit = source.edit().unwrap();
        assert!(
            edit.set_slide_transition(Position::new(index), Some(&target))
                .unwrap()
        );
        let (changed, patch) = edit.commit().unwrap().into_parts();
        let after = changed.to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert!(value(&reopened, index).unwrap().same_semantics(&target));
        assert_preserved(&bytes, &after, index);
        let mut no_change = changed.edit().unwrap();
        assert!(
            !no_change
                .set_slide_transition(Position::new(index), Some(&target))
                .unwrap()
        );
        assert!(no_change.commit().unwrap().into_parts().1.is_noop());
        let mut clear = changed.edit().unwrap();
        assert!(
            clear
                .set_slide_transition(Position::new(index), None)
                .unwrap()
        );
        let cleared = clear.commit().unwrap().into_package();
        assert!(
            value(
                &Package::from_bytes(&cleared.to_bytes().unwrap()).unwrap(),
                index
            )
            .is_none()
        );
        assert_preserved(&bytes, &cleared.to_bytes().unwrap(), index);
        assert!(matches!(
            cleared.apply(&patch.inverse()),
            Err(litchi_pptx::Error::PatchConflict)
        ));
        let cleared_reopened = litchi_pptx::source::Package::from_read_at(Arc::new(
            litchi_core::OwnedSource::new(cleared.to_bytes().unwrap()),
        ))
        .unwrap();
        let mut restore = cleared_reopened.edit().unwrap();
        assert!(
            restore
                .set_slide_transition(Position::new(index), Some(&target))
                .unwrap()
        );
        let restored = restore.commit().unwrap().into_package();
        assert!(
            value(
                &Package::from_bytes(&restored.to_bytes().unwrap()).unwrap(),
                index
            )
            .unwrap()
            .same_semantics(&target)
        );
        let eager = Package::from_bytes(&bytes).unwrap();
        let mut edit = eager.edit().unwrap();
        edit.set_slide_transition(index, Some(&target)).unwrap();
        let eager = edit.commit().unwrap();
        let a = members(&after);
        let b = members(&eager.package().to_bytes().unwrap());
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        for (name, data) in a {
            // Changed eager XML parts use the compact serializer, which
            // removes the native declaration's CRLF. Source overlays keep it.
            let expected = if name == format!("ppt/slides/slide{}.xml", index + 1) {
                String::from_utf8(data)
                    .unwrap()
                    .replacen("?>\r\n", "?>", 1)
                    .into_bytes()
            } else {
                data
            };
            assert!(expected == b[&name], "source/eager mismatch in {name}");
        }
        assert!(
            changed
                .apply(&patch.inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
                == bytes,
            "inverse source bytes"
        );
    }
    assert!(source.to_bytes().unwrap() == bytes, "source bytes changed");
}

#[test]
fn native_zoom_transition_admission_keeps_shape_edit_guards_and_failure_atomicity() {
    let bytes = native_bytes();
    let source = litchi_pptx::source::Package::from_read_at(Arc::new(
        litchi_core::OwnedSource::new(bytes.clone()),
    ))
    .unwrap();
    let mut edit = source.edit().unwrap();
    for result in [
        edit.set_shape_text(0, 0_usize, "must stay untouched"),
        edit.remove_shape(0, 0_usize),
    ] {
        assert!(
            matches!(result,Err(litchi_pptx::Error::Refused(refusal)) if refusal.kind()==litchi_core::edit::RefusalKind::UnsupportedContent)
        );
    }
    assert!(edit.commit().unwrap().into_package().to_bytes().unwrap() == bytes);
    for case in 0..3 {
        let package = Package::from_bytes(&bytes).unwrap();
        let mut opc = package.opc().unwrap().clone();
        let uri = litchi_opc::PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = std::str::from_utf8(opc.get_part(&uri).unwrap().blob()).unwrap();
        let changed=match case {
            0=>xml.replacen("Requires=", "MissingRequires=",1),
            1=>xml.replacen("<p:spTree>","<p:spTree><p:transition/>",1),
            _=>xml.replacen("<p:cSld>",r#"<p:cSld><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future"><mc:Choice Requires="x"/><mc:Fallback/></mc:AlternateContent>"#,1),
        };
        assert_ne!(changed, xml);
        opc.get_part_mut(&uri)
            .unwrap()
            .set_blob(changed.into_bytes());
        let eager = Package::from_opc_package(opc).unwrap();
        let before = eager.to_bytes().unwrap();
        let source = litchi_pptx::source::Package::from_read_at(Arc::new(
            litchi_core::OwnedSource::new(before.clone()),
        ))
        .unwrap();
        let mut eager_edit = eager.edit().unwrap();
        let mut source_edit = source.edit().unwrap();
        for error in [
            eager_edit
                .set_slide_transition(0_usize, Some(&desired(0)))
                .unwrap_err(),
            source_edit
                .set_slide_transition(0, Some(&desired(0)))
                .unwrap_err(),
        ] {
            if case == 0 {
                assert!(
                    matches!(error, litchi_pptx::Error::MarkupCompatibility(_)),
                    "{error:?}"
                );
            } else {
                assert!(
                    matches!(error,litchi_pptx::Error::Refused(refusal) if refusal.kind()==litchi_core::edit::RefusalKind::UnsupportedContent),
                    "{error:?}"
                );
            }
        }
        if case == 0 {
            // A coherent eager snapshot cannot publish this malformed MCE
            // source, even after a rejected mutation. The source stays exact.
            assert!(matches!(
                eager_edit.commit(),
                Err(litchi_pptx::Error::MarkupCompatibility(_))
            ));
        } else {
            assert!(eager_edit.commit().unwrap().package().to_bytes().unwrap() == before);
        }
        assert!(eager.to_bytes().unwrap() == before);
        assert!(
            source_edit
                .commit()
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap()
                == before
        );
    }
}

#[test]
fn native_morph_namespace_alias_variant_can_clear_and_restore_its_active_branch() {
    let package = Package::from_bytes(&native_bytes()).unwrap();
    let mut opc = package.opc().unwrap().clone();
    let uri = litchi_opc::PackURI::new("/ppt/slides/slide3.xml").unwrap();
    let xml = std::str::from_utf8(opc.get_part(&uri).unwrap().blob()).unwrap();
    // A lexical in-memory variant tests inherited prefix resolution. The
    // tracked Office fixture and all namespace URIs remain unchanged.
    let alias = xml
        .replace("<p:", "<q:")
        .replace("</p:", "</q:")
        .replace("xmlns:p=", "xmlns:q=");
    opc.get_part_mut(&uri).unwrap().set_blob(alias.into_bytes());
    let package = Package::from_opc_package(opc).unwrap();
    let mut edit = package.edit().unwrap();
    assert!(edit.set_slide_transition(2_usize, None).unwrap());
    let cleared = edit.commit().unwrap();
    assert!(value(cleared.package(), 2).is_none());
    let mut edit = cleared.package().edit().unwrap();
    let target = desired(2);
    assert!(edit.set_slide_transition(2_usize, Some(&target)).unwrap());
    let restored = edit.commit().unwrap();
    assert!(
        value(restored.package(), 2)
            .unwrap()
            .same_semantics(&target)
    );
}

#[test]
fn native_transition_edits_preserve_all_unencrypted_variants_and_resources() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    for (name, count) in [
        ("powerpoint-text-table.pptx", 3),
        ("powerpoint-visual-objects.pptx", 5),
        ("powerpoint-visual-objects.potx", 5),
        ("powerpoint-visual-objects.ppsx", 5),
        ("powerpoint-navigation-review.pptx", 6),
        ("powerpoint-media-objects.pptx", 4),
        ("powerpoint-passive-macro.pptm", 1),
        ("powerpoint-passive-macro.potm", 1),
        ("powerpoint-passive-macro.ppsm", 1),
        ("powerpoint-master-motion.pptx", 7),
        ("powerpoint-modern-media.pptx", 8),
        ("powerpoint-fonts-math-handout.pptx", 4),
        ("powerpoint-ink-stroke.pptx", 1),
        ("powerpoint-ink-replay.pptx", 1),
    ] {
        let bytes = std::fs::read(root.join(name)).unwrap();
        let eager = Package::from_bytes(&bytes).unwrap();
        let eager_bytes = eager.to_bytes().unwrap();
        assert_eq!(eager.presentation().unwrap().slide_count().unwrap(), count);
        let source = litchi_pptx::source::Package::from_read_at(Arc::new(
            litchi_core::OwnedSource::new(bytes.clone()),
        ))
        .unwrap();
        let target = Transition::new(Kind::Fade { black: Some(true) })
            .with_duration(Ms::new(1750).unwrap())
            .with_after(Ms::new(12345).unwrap())
            .with_click(false);
        for index in 0..count {
            let mut edit = eager.edit().unwrap();
            assert!(
                edit.set_slide_transition(index, Some(&target))
                    .unwrap_or_else(|e| panic!("{name} slide{index} eager: {e}"))
            );
            let committed = edit.commit().unwrap();
            let after = committed.package().to_bytes().unwrap();
            assert_preserved(&bytes, &after, index);
            assert!(
                value(&Package::from_bytes(&after).unwrap(), index)
                    .unwrap()
                    .same_semantics(&target)
            );
            assert!(
                committed
                    .package()
                    .apply(&committed.patch().inverse())
                    .unwrap()
                    .package()
                    .to_bytes()
                    .unwrap()
                    == eager_bytes
            );
            let mut edit = source.edit().unwrap();
            assert!(
                edit.set_slide_transition(Position::new(index), Some(&target))
                    .unwrap_or_else(|e| panic!("{name} slide{index} source: {e}"))
            );
            let (changed, patch) = edit.commit().unwrap().into_parts();
            let after = changed.to_bytes().unwrap();
            assert_preserved(&bytes, &after, index);
            assert!(
                value(&Package::from_bytes(&after).unwrap(), index)
                    .unwrap()
                    .same_semantics(&target)
            );
            assert!(
                changed
                    .apply(&patch.inverse())
                    .unwrap()
                    .into_package()
                    .to_bytes()
                    .unwrap()
                    == bytes
            );
        }
        assert!(source.to_bytes().unwrap() == bytes);
        assert!(eager.to_bytes().unwrap() == eager_bytes);
    }
}

#[test]
fn native_presentation_structure_variants_keep_the_saved_slide_catalog() {
    use litchi_pptx::presentation_properties::metadata::structure;
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/powerpoint");
    for (name, ids) in [
        (
            "powerpoint-visual-objects.pptx",
            [256, 257, 258, 259, 260].as_slice(),
        ),
        (
            "powerpoint-visual-objects.potx",
            [256, 257, 258, 259, 260].as_slice(),
        ),
        (
            "powerpoint-visual-objects.ppsx",
            [256, 257, 258, 259, 260].as_slice(),
        ),
        ("powerpoint-passive-macro.pptm", [256].as_slice()),
        ("powerpoint-passive-macro.potm", [256].as_slice()),
        ("powerpoint-passive-macro.ppsm", [256].as_slice()),
    ] {
        let package = Package::open(root.join(name)).unwrap();
        let opc = package.opc().unwrap();
        let graph = structure::load(opc, structure::Limits::DEFAULT).unwrap();
        assert_eq!(
            graph.slides.iter().map(|s| s.slide_id).collect::<Vec<_>>(),
            ids
        );
        for (index, slide) in graph.slides.iter().enumerate() {
            assert_eq!(
                slide.part_name,
                format!("/ppt/slides/slide{}.xml", index + 1)
            );
            let relationship = opc
                .main_document_part()
                .unwrap()
                .rels()
                .get(&slide.relationship_id)
                .unwrap();
            assert!(!relationship.is_external());
            assert_eq!(
                relationship.target_partname().unwrap().as_str(),
                slide.part_name
            );
        }
        let mut invalid = opc.clone();
        let main = invalid.main_document_part().unwrap().partname().clone();
        invalid
            .get_part_mut(&main)
            .unwrap()
            .set_content_type(litchi_opc::constants::content_type::WML_DOCUMENT_MAIN.into())
            .unwrap();
        assert!(matches!(
            structure::load(&invalid, structure::Limits::DEFAULT),
            Err(litchi_pptx::Error::Invalid(_))
        ));
    }
}
