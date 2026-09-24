use super::*;
use litchi_docx::ink::{
    AnchorGeometry, BaseProfile, BrushDraft, ContextDraft, ContextKind, Destination, Draft,
    EditLimits, FallbackImage, Geometry, Location, Placement, Prepared, Style, TraceDraft,
};
use soapberry_zip::office::ArchiveReader;
use std::collections::BTreeMap;

fn prepared() -> Prepared {
    Draft::default().finish().unwrap()
}

fn strokes() -> Prepared {
    let mut draft = Draft::default()
        .context(
            ContextDraft::new(ContextKind::InkDrawing)
                .with_xml_id("ctx")
                .unwrap(),
        )
        .unwrap()
        .brush(BrushDraft::new("brush").unwrap())
        .unwrap();
    for data in ["0 0, 10 20, -10 30", "5 6, 7 8"] {
        draft = draft
            .trace(
                TraceDraft::new(data)
                    .unwrap()
                    .with_context_ref("#ctx")
                    .unwrap()
                    .with_brush_ref("#brush")
                    .unwrap(),
            )
            .unwrap();
    }
    draft.finish().unwrap()
}
fn image() -> FallbackImage {
    FallbackImage::from_bytes(PNG.to_vec()).unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}
const PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
    0, 0, 31, 21, 196, 137, 0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 16, 80, 48, 248, 15, 0, 2,
    4, 1, 96, 141, 188, 187, 113, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

#[test]
fn creates_base_ink_in_all_story_roles_and_restores_exact_members_after_reopen() {
    for strict in [false, true] {
        for profile in [BaseProfile::WordTextXml, BaseProfile::InkContent] {
            let original = PackageWriter::to_bytes(&source(strict, true)).unwrap();
            let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
            let mut edit = package.edit_ink().unwrap();
            let locations: Vec<_> = edit
                .snapshot()
                .annotations()
                .iter()
                .map(|annotation| annotation.location())
                .collect();
            for location in locations {
                edit.insert(
                    Destination::new(location, Position::new(0)),
                    prepared(),
                    Style::Base(profile),
                )
                .unwrap();
            }
            let commit = edit.commit().unwrap();
            assert_eq!(package.ink().unwrap().annotations().len(), 8);
            assert_eq!(commit.snapshot().annotations().len(), 16);
            package.apply_ink_patch(commit.patch()).unwrap();
            let published = save(&mut package);
            let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
            let inventory = reopened.ink().unwrap();
            assert_eq!(inventory.distinct_payload_count(), 9);
            for pair in inventory.annotations().chunks_exact(2) {
                assert_eq!(pair[0].trace_count(), 1);
                assert_eq!(pair[1].trace_count(), 0);
                assert_eq!(pair[0].location(), pair[1].location());
            }
            reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
            assert_eq!(members(&save(&mut reopened)), members(&original));
        }
    }
}

#[test]
fn selected_replacement_clones_shared_payload_and_preserves_unselected_annotations() {
    for strict in [false, true] {
        let original = PackageWriter::to_bytes(&source(strict, true)).unwrap();
        let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
        let mut edit = package.edit_ink().unwrap();
        assert!(edit.replace(Position::new(0), prepared(), None).unwrap());
        assert!(!edit.replace(Position::new(0), prepared(), None).unwrap());
        let commit = package.publish_ink_edit(edit).unwrap();
        let annotations = package.ink().unwrap();
        assert_eq!(annotations.annotations()[0].trace_count(), 0);
        assert!(
            annotations.annotations()[1..]
                .iter()
                .all(|annotation| annotation.trace_count() == 1)
        );
        assert_eq!(annotations.distinct_payload_count(), 2);
        assert_eq!(
            package
                .opc_package()
                .get_part(&PackURI::new("/payload/handwriting.xml").unwrap())
                .unwrap()
                .blob(),
            INK
        );
        let published = save(&mut package);
        let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
        reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
        assert_eq!(members(&save(&mut reopened)), members(&original));
    }
}

#[test]
fn insertion_expands_empty_paragraph_and_preserves_call_order() {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/word/document.xml",
        ct::WML_DOCUMENT_MAIN,
        format!("<w:document xmlns:w=\"{W}\"><w:body><w:p/></w:body></w:document>").as_bytes(),
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::WordTextXml),
    )
    .unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert_eq!(package.ink().unwrap().annotations().len(), 2);
    let document = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let targets: Vec<_> = document
        .rels()
        .iter()
        .filter(|edge| edge.reltype() == rt::CUSTOM_XML)
        .map(|edge| edge.target_partname().unwrap())
        .collect();
    assert_eq!(targets.len(), 2);
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

#[test]
fn drawing_canvas_group_creation_has_fallback_and_reversible_resource_closure() {
    let geometry = Geometry::new(914400, 914400).unwrap();
    for placement in [
        Placement::Inline(geometry),
        Placement::Anchor(AnchorGeometry::new(geometry)),
    ] {
        for (name, style) in [
            ("drawing", Style::drawing(placement, image())),
            ("canvas", Style::canvas(placement, image())),
            ("group", Style::group(placement, image())),
        ] {
            let description = format!("{name} {placement:?}");
            let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
            let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
            let mut edit = package.edit_ink().unwrap();
            let payload = strokes();
            edit.insert(Destination::main(Position::new(0)), payload.clone(), style)
                .unwrap();
            let commit = package.publish_ink_edit(edit).unwrap();
            assert_eq!(package.ink().unwrap().annotations().len(), 2);
            let published = save(&mut package);
            let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
            assert_eq!(reopened.ink().unwrap().annotations().len(), 2);
            assert_eq!(reopened.ink().unwrap().annotations()[1].trace_count(), 2);
            assert!(
                reopened
                    .opc_package()
                    .try_iter_parts()
                    .any(|part| part.expect("decode part payload").blob() == payload.as_bytes())
            );
            let mut unchanged = reopened.edit_ink().unwrap();
            assert!(
                !unchanged
                    .replace(Position::new(1), payload.clone(), None)
                    .unwrap()
            );
            assert!(
                unchanged
                    .replace(Position::new(1), prepared(), None)
                    .unwrap()
            );
            assert!(unchanged.replace(Position::new(1), payload, None).unwrap());
            let unchanged = reopened.publish_ink_edit(unchanged).unwrap();
            assert!(unchanged.patch().is_empty());
            assert_eq!(members(&save(&mut reopened)), members(&published));
            assert!(reopened.opc_package().try_iter_parts().any(|part| {
                let part = part.expect("decode part payload");
                part.content_type() == "image/png" && part.blob() == PNG
            }));
            let mut replacement = reopened.edit_ink().unwrap();
            replacement
                .replace(Position::new(1), prepared(), Some(image()))
                .unwrap();
            let replacement = reopened
                .publish_ink_edit(replacement)
                .unwrap_or_else(|error| panic!("{description}: {error}"));
            assert_eq!(reopened.ink().unwrap().annotations()[1].trace_count(), 0);
            let mut removal = reopened.edit_ink().unwrap();
            assert!(removal.remove(Position::new(1)).unwrap());
            let removal = reopened.publish_ink_edit(removal).unwrap();
            assert_eq!(reopened.ink().unwrap().annotations().len(), 1);
            reopened
                .apply_ink_patch(&removal.patch().inverse())
                .unwrap();
            reopened
                .apply_ink_patch(&replacement.patch().inverse())
                .unwrap();
            reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
            assert_eq!(members(&save(&mut reopened)), members(&original));
        }
    }
}

#[test]
fn invalid_destinations_strict_drawings_and_conflicting_intents_refuse_atomically() {
    for destination in [
        Destination::main(Position::new(999)),
        Destination::new(
            Location::new(StoryKind::Header, Position::new(999)),
            Position::new(0),
        ),
    ] {
        let mut package = open(&source(false, false));
        let original = members(&save(&mut package));
        let mut edit = package.edit_ink().unwrap();
        edit.insert(
            destination,
            prepared(),
            Style::Base(BaseProfile::InkContent),
        )
        .unwrap();
        assert!(package.publish_ink_edit(edit).is_err());
        assert_eq!(members(&save(&mut package)), original);
    }
    let mut package = open(&source(true, false));
    let original = members(&save(&mut package));
    let mut edit = package.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::drawing(Placement::Inline(Geometry::new(1, 1).unwrap()), image()),
    )
    .unwrap();
    assert!(package.publish_ink_edit(edit).is_err());
    assert_eq!(members(&save(&mut package)), original);
    let mut edit = package.edit_ink().unwrap();
    edit.replace(Position::new(0), prepared(), None).unwrap();
    assert!(edit.remove(Position::new(0)).is_err());
    assert_eq!(edit.commit().unwrap().snapshot().annotations().len(), 1);
}

#[test]
fn drawing_replacement_updates_fallback_preserves_geometry_and_undoes_after_reopen() {
    let original = PackageWriter::to_bytes(&removal::drawing_source("", "")).unwrap();
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut missing = package.edit_ink().unwrap();
    missing.replace(Position::new(0), prepared(), None).unwrap();
    assert!(package.publish_ink_edit(missing).is_err());
    assert_eq!(members(&save(&mut package)), members(&original));
    let mut edit = package.edit_ink().unwrap();
    edit.replace(Position::new(0), prepared(), Some(image()))
        .unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert_eq!(package.ink().unwrap().annotations()[0].trace_count(), 0);
    let published = save(&mut package);
    let published_members = members(&published);
    let document = std::str::from_utf8(&published_members["word/document.xml"]).unwrap();
    assert!(document.contains("<wp:extent cx=\"127000\" cy=\"127000\"/>"));
    assert!(document.contains("style=\"width:10pt;height:10pt\""));
    assert!(document.contains("<w:t>before</w:t>"));
    assert!(document.contains("<w:t>after</w:t>"));
    assert!(!published_members.contains_key("payload/handwriting.xml"));
    assert!(!published_members.contains_key("payload/fallback.png"));
    assert!(published_members.values().any(|bytes| bytes == PNG));
    let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
    reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn authoring_limits_refuse_without_changing_prior_intents_or_package() {
    let mut package = open(&source(false, false));
    let original = members(&save(&mut package));
    let mut edit = package
        .edit_ink_with_limits(EditLimits {
            max_operations: 1,
            ..EditLimits::default()
        })
        .unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    assert!(edit.replace(Position::new(0), prepared(), None).is_err());
    let commit = package.publish_ink_edit(edit).unwrap();
    assert_eq!(package.ink().unwrap().annotations().len(), 2);
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), original);

    let payload = strokes();
    let mut edit = package
        .edit_ink_with_limits(EditLimits {
            max_staged_bytes: payload.as_bytes().len() - 1,
            ..EditLimits::default()
        })
        .unwrap();
    assert!(
        edit.insert(
            Destination::main(Position::new(0)),
            payload,
            Style::Base(BaseProfile::InkContent)
        )
        .is_err()
    );
    assert!(edit.commit().unwrap().patch().is_empty());
    assert_eq!(members(&save(&mut package)), original);

    let payload = strokes();
    let mut edit = package
        .edit_ink_with_limits(EditLimits {
            max_staged_bytes: payload.as_bytes().len(),
            ..EditLimits::default()
        })
        .unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        payload,
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    assert!(package.publish_ink_edit(edit).is_err());
    assert_eq!(members(&save(&mut package)), original);
}

#[test]
fn mixed_insert_replace_remove_uses_base_selectors_and_one_reversible_commit() {
    let original = PackageWriter::to_bytes(&source(false, true)).unwrap();
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    edit.replace(Position::new(1), prepared(), None).unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    let snapshot = package.ink().unwrap();
    assert_eq!(snapshot.annotations().len(), 8);
    assert_eq!(snapshot.annotations()[0].trace_count(), 0);
    assert_eq!(snapshot.annotations()[1].trace_count(), 0);
    assert!(
        snapshot.annotations()[2..]
            .iter()
            .all(|annotation| annotation.trace_count() == 1)
    );
    let published = save(&mut package);
    let mut reopened = Package::from_reader(Cursor::new(&published)).unwrap();
    reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn authored_patch_guards_graph_and_payload_and_requires_explicit_unsign() {
    let original = PackageWriter::to_bytes(&source(false, false)).unwrap();
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let mut changed = source(false, false);
    part(
        &mut changed,
        "/payload/handwriting.xml",
        INK_TYPE,
        b"<ink xmlns=\"http://www.w3.org/2003/InkML\"/>",
    );
    let mut changed = open(&changed);
    let unchanged = members(&save(&mut changed));
    assert!(changed.apply_ink_patch(commit.patch()).is_err());
    assert_eq!(members(&save(&mut changed)), unchanged);
    let mut signed = source(false, false);
    part(
        &mut signed,
        "/_xmlsignatures/origin.sigs",
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        b"",
    );
    signed.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let mut signed = open(&signed);
    let original = members(&save(&mut signed));
    let mut edit = signed.edit_ink().unwrap();
    edit.insert(
        Destination::main(Position::new(0)),
        prepared(),
        Style::Base(BaseProfile::InkContent),
    )
    .unwrap();
    assert!(signed.publish_ink_edit(edit).is_err());
    assert_eq!(members(&save(&mut signed)), original);
    // The unrelated package still accepts its source-bound prepared patch.
    package.apply_ink_patch(commit.patch()).unwrap();
    assert_eq!(package.ink().unwrap().annotations().len(), 2);
}
