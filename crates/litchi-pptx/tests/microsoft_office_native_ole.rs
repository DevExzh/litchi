//! Inert native Excel OLE graph discovery and source-checked package edits.
use litchi_opc::PackURI;
use litchi_pptx::{
    Package,
    presentation::embedded::ole::{self, Frame, Kind, Mode, slide},
};
use std::collections::BTreeMap;
const EMBEDDED: &str = "/ppt/embeddings/Microsoft_Excel_Worksheet.xlsx";
const SLIDE: &str = "/ppt/slides/slide4.xml";
fn native() -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint/powerpoint-media-objects.pptx"),
    )
    .unwrap()
}
fn snapshot(package: &Package) -> slide::Snapshot {
    slide::load(package.opc().unwrap(), 3, &PackURI::new(SLIDE).unwrap()).unwrap()
}
fn sha256(b: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(b)
        .iter()
        .map(|v| format!("{v:02x}"))
        .collect()
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
fn assert_changed(before: &[u8], after: &[u8], names: &[&str]) {
    let a = members(before);
    let b = members(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    assert_eq!(
        a.iter()
            .filter(|(n, v)| **v != b[*n])
            .map(|(n, _)| n.as_str())
            .collect::<Vec<_>>(),
        names
    );
}

#[test]
fn native_excel_ole_inventory_keeps_frame_payload_type_and_preview() {
    let package = native();
    let p = package.presentation().unwrap();
    let slides = p.slides().unwrap();
    let found = ole::load_slide(
        package.opc().unwrap(),
        3,
        slides[3].part().part(),
        &mut ole::Limits::default(),
    )
    .unwrap();
    assert_eq!(found.len(), 1);
    let object = &found[0];
    assert_eq!((object.slide_index(), object.index()), (3, 0));
    assert_eq!(object.shape_id(), Some(4));
    assert_eq!(object.shape_name(), Some("Content Placeholder 3"));
    assert_eq!(object.name(), Some("Worksheet"));
    assert_eq!(object.program_id(), Some("Excel.Sheet.12"));
    assert_eq!(object.mode(), Mode::Embedded);
    assert_eq!(object.kind(), Some(Kind::Package));
    assert_eq!(object.relationship_id(), Some("rId2"));
    assert_eq!(
        object.anchor(),
        Some(Frame::new(4_851_400, 3_486_150, 2_489_200, 1_028_700))
    );
    assert_eq!(object.preview_width(), Some(2_489_200));
    assert_eq!(object.preview_height(), Some(1_028_700));
    assert_eq!(object.show_as_icon(), None);
    // The active VML-capable branch has no picture. The EMF preview below
    // belongs to the retained fallback and remains an inert resource.
    assert_eq!(object.preview_relationship_id(), None);
    let target = object.target().unwrap();
    assert_eq!(target.part_name().unwrap().as_str(), EMBEDDED);
    assert_eq!(
        target.content_type(),
        Some(litchi_opc::constants::content_type::SML_SHEET)
    );
    let snapshot = snapshot(&package);
    assert_eq!(snapshot.objects(), found);
    let payload = snapshot.payload(0).unwrap().unwrap();
    assert_eq!(payload.len(), 18372);
    assert_eq!(
        sha256(payload),
        "07fbdf1bd05d9f080356bbc10bc9cc897ffd794a166b2c141f45bad707051bbc"
    );
    let preview = package
        .opc()
        .unwrap()
        .get_part(&PackURI::new("/ppt/media/image3.emf").unwrap())
        .unwrap();
    assert_eq!(preview.content_type(), "image/x-emf");
    assert_eq!(preview.blob().len(), 81992);
    assert_eq!(
        sha256(preview.blob()),
        "d2a459ae9eec497ac562a4b96569b9dabe2e161f9d961c74192979ab95a7597b"
    );
}

#[test]
fn native_excel_ole_metadata_edit_keeps_inactive_branch_and_unrelated_members() {
    let package = native();
    let before = package.to_bytes().unwrap();
    let original = snapshot(&package);
    let mut edit = package.edit_slide_ole_objects().unwrap();
    edit.update(3_usize, |draft| {
        assert!(draft.set_name(0, Some("Reviewed worksheet".into()))?);
        assert!(draft.set_show_as_icon(0, Some(true))?);
        Ok(())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".pptx").unwrap();
    commit.package().save(output.path()).unwrap();
    let reopened = Package::open(output.path()).unwrap();
    let changed = snapshot(&reopened);
    assert_eq!(changed.objects().len(), 1);
    let object = &changed.objects()[0];
    assert_eq!(object.name(), Some("Reviewed worksheet"));
    assert_eq!(object.show_as_icon(), Some(true));
    assert_eq!(object.anchor(), original.objects()[0].anchor());
    assert_eq!(changed.payload(0).unwrap(), original.payload(0).unwrap());
    assert_changed(&before, &after, &["ppt/slides/slide4.xml"]);
    let a = members(&before);
    let b = members(&after);
    let raw = std::str::from_utf8(&a["ppt/slides/slide4.xml"]).unwrap();
    let fallback = raw
        .split_once("<mc:Fallback>")
        .unwrap()
        .1
        .split_once("</mc:Fallback>")
        .unwrap()
        .0;
    assert!(
        std::str::from_utf8(&b["ppt/slides/slide4.xml"])
            .unwrap()
            .contains(fallback)
    );
    assert!(
        commit
            .package()
            .apply_slide_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
    assert!(package.to_bytes().unwrap() == before);
}

#[test]
fn native_excel_ole_payload_replacement_keeps_specific_mime_frame_and_preview() {
    let replacement = include_bytes!(
        "../../../test-data/microsoft-office/excel/excel-values-formulas-chart.xlsx"
    );
    let package = native();
    let before = package.to_bytes().unwrap();
    let original = snapshot(&package);
    assert_ne!(
        original.payload(0).unwrap().unwrap().len(),
        replacement.len()
    );
    let mut edit = package.edit_slide_ole_objects().unwrap();
    edit.update(3_usize, |draft| {
        draft.replace_payload(0, replacement.to_vec())
    })
    .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let changed = snapshot(&reopened);
    assert_eq!(changed.objects(), original.objects());
    assert!(changed.payload(0).unwrap().unwrap() == replacement);
    assert_eq!(
        changed.objects()[0].target().unwrap().content_type(),
        Some(litchi_opc::constants::content_type::SML_SHEET)
    );
    assert_changed(
        &before,
        &after,
        &["ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"],
    );
    assert!(
        commit
            .package()
            .apply_slide_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
    assert!(package.to_bytes().unwrap() == before);
}

#[test]
fn native_excel_ole_remove_and_detach_collect_payload_but_keep_unowned_preview() {
    for detach in [false, true] {
        let package = native();
        let before = package.to_bytes().unwrap();
        let a = members(&before);
        let mut edit = package.edit_slide_ole_objects().unwrap();
        edit.update(3_usize, |draft| {
            let removed = if detach {
                draft.detach(0)?
            } else {
                draft.remove(0)?
            };
            assert_eq!(removed.shape_id(), Some(4));
            Ok(())
        })
        .unwrap();
        let commit = edit.commit().unwrap();
        let after = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert!(snapshot(&reopened).objects().is_empty());
        let b = members(&after);
        assert_eq!(
            a.keys()
                .filter(|n| !b.contains_key(*n))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            ["ppt/embeddings/Microsoft_Excel_Worksheet.xlsx"]
        );
        let original = std::str::from_utf8(&a["ppt/slides/slide4.xml"]).unwrap();
        let start = original.find("<p:graphicFrame>").unwrap();
        let end = original.find("</p:graphicFrame>").unwrap() + "</p:graphicFrame>".len();
        assert!(
            b["ppt/slides/slide4.xml"]
                == format!("{}{}", &original[..start], &original[end..]).as_bytes()
        );
        for (name, data) in &b {
            if ![
                "[Content_Types].xml",
                "ppt/slides/slide4.xml",
                "ppt/slides/_rels/slide4.xml.rels",
            ]
            .contains(&name.as_str())
            {
                assert!(data == &a[name], "changed {name}");
            }
        }
        let slide = reopened
            .opc()
            .unwrap()
            .get_part(&PackURI::new(SLIDE).unwrap())
            .unwrap();
        assert!(slide.rels().get("rId2").is_none());
        assert!(slide.rels().get("rId3").is_some());
        assert!(
            commit
                .package()
                .apply_slide_ole_objects(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == before
        );
        assert!(package.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_excel_ole_source_frame_removal_has_durable_forward_and_inverse() {
    use litchi_pptx::source::{Package as Source, Patch};
    use std::sync::Arc;
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint/powerpoint-media-objects.pptx"),
    )
    .unwrap();
    let open = |bytes: Vec<u8>| {
        Source::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes))).unwrap()
    };
    let source = open(bytes.clone());
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_shape(3, "Content Placeholder 3").unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let after = changed.to_bytes().unwrap();
    assert_changed(&bytes, &after, &["ppt/slides/slide4.xml"]);
    let a = members(&bytes);
    let b = members(&after);
    let xml = std::str::from_utf8(&a["ppt/slides/slide4.xml"]).unwrap();
    let start = xml.find("<p:graphicFrame>").unwrap();
    let end = xml.find("</p:graphicFrame>").unwrap() + "</p:graphicFrame>".len();
    assert!(b["ppt/slides/slide4.xml"] == format!("{}{}", &xml[..start], &xml[end..]).as_bytes());
    // Source shape removal retains relationships/resources by contract; the
    // graph-owning remove/detach API above performs payload collection.
    let forward = Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let replay = open(bytes.clone())
        .apply(&forward)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert!(members(&replay) == b);
    let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let restored = open(after)
        .apply(&inverse)
        .unwrap()
        .into_package()
        .to_bytes()
        .unwrap();
    assert!(members(&restored) == a);
    assert!(
        changed
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(source.to_bytes().unwrap() == bytes);
}

#[test]
fn native_excel_ole_invalid_and_shared_payload_edits_leave_source_unchanged() {
    let package = native();
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit_slide_ole_objects().unwrap();
    let error = edit
        .update(3_usize, |draft| {
            draft.set_name(0, Some("must roll back".into()))?;
            draft.replace_payload(0, Vec::new())
        })
        .unwrap_err();
    assert!(matches!(error, litchi_pptx::Error::Invalid(_)));
    assert!(edit.commit().unwrap().package().to_bytes().unwrap() == before);
    let mut opc = package.opc().unwrap().clone();
    let other = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    opc.get_part_mut(&other)
        .unwrap()
        .rels_mut()
        .add_relationship(
            litchi_opc::constants::relationship_type::PACKAGE.into(),
            "../embeddings/Microsoft_Excel_Worksheet.xlsx".into(),
            "rIdShared".into(),
            false,
        );
    let shared = Package::from_opc_package(opc).unwrap();
    let original = shared.to_bytes().unwrap();
    let mut edit = shared.edit_slide_ole_objects().unwrap();
    let error = edit
        .update(3_usize, |draft| {
            draft.replace_payload(0, b"inert replacement".to_vec())
        })
        .unwrap_err();
    assert!(
        matches!(error,litchi_pptx::Error::Invalid(message) if message.contains("shared OLE payload"))
    );
    assert!(edit.commit().unwrap().package().to_bytes().unwrap() == original);
}

#[test]
fn native_excel_ole_addition_allocates_a_distinct_frame_relationship_and_payload() {
    let package = native();
    let before = package.to_bytes().unwrap();
    let source = snapshot(&package);
    let payload = source.payload(0).unwrap().unwrap();
    let frame = Frame::new(100_000, 200_000, 2_000_000, 900_000);
    let definition = slide::Definition::embedded(Kind::Package, frame, payload.to_vec())
        .set_name("Additional workbook")
        .set_program_id("Excel.Sheet.12");
    let mut edit = package.edit_slide_ole_objects().unwrap();
    let index = edit.update(3_usize, |draft| draft.add(definition)).unwrap();
    assert_eq!(index, 1);
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let result = snapshot(&reopened);
    assert_eq!(result.objects().len(), 2);
    assert_eq!(result.objects()[0], source.objects()[0]);
    let added = &result.objects()[1];
    assert_eq!(added.shape_id(), Some(5));
    assert_eq!(added.name(), Some("Additional workbook"));
    assert_eq!(added.program_id(), Some("Excel.Sheet.12"));
    assert_eq!(added.anchor(), Some(frame));
    assert_eq!(added.relationship_id(), Some("rId4"));
    assert_eq!(
        added.target().unwrap().part_name().unwrap().as_str(),
        "/ppt/embeddings/package1.bin"
    );
    assert_eq!(
        added.target().unwrap().content_type(),
        Some(litchi_opc::constants::content_type::OFC_PACKAGE)
    );
    assert!(result.payload(1).unwrap().unwrap() == payload);
    let a = members(&before);
    let b = members(&after);
    assert_eq!(
        b.keys()
            .filter(|n| !a.contains_key(*n))
            .map(String::as_str)
            .collect::<Vec<_>>(),
        ["ppt/embeddings/package1.bin"]
    );
    for (name, data) in &a {
        if ![
            "[Content_Types].xml",
            "ppt/slides/slide4.xml",
            "ppt/slides/_rels/slide4.xml.rels",
        ]
        .contains(&name.as_str())
        {
            assert!(data == &b[name], "changed {name}");
        }
    }
    assert!(
        commit
            .package()
            .apply_slide_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}
