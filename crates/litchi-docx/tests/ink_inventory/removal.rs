use std::collections::BTreeMap;

use litchi_docx::ink::EditLimits;
use soapberry_zip::office::ArchiveReader;

use super::*;

fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive = ArchiveReader::new(bytes).unwrap();
    archive
        .file_names()
        .map(|name| (name.to_owned(), archive.read(name).unwrap()))
        .collect()
}

fn has_part(package: &Package, name: &str) -> bool {
    package
        .opc_package()
        .get_part(&PackURI::new(name).unwrap())
        .is_ok()
}

#[test]
fn base_removal_across_all_stories_retains_shared_target_until_last_owner_and_undoes_exactly() {
    for strict in [false, true] {
        let original = PackageWriter::to_bytes(&source(strict, true)).unwrap();
        let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
        let mut first = package.edit_ink().unwrap();
        assert!(first.remove(Position::new(0)).unwrap());
        assert!(!first.remove(Position::new(0)).unwrap());
        assert!(first.remove(Position::new(8)).is_err());
        let first = first.commit().unwrap();
        assert!(first.changed());
        assert_eq!(first.snapshot().annotations().len(), 7);
        assert_eq!(package.ink().unwrap().annotations().len(), 8);
        package.apply_ink_patch(first.patch()).unwrap();
        assert!(has_part(&package, "/payload/handwriting.xml"));
        let intermediate = save(&mut package);
        let mut package = Package::from_reader(Cursor::new(&intermediate)).unwrap();
        let mut rest = package.edit_ink().unwrap();
        for position in 0..7 {
            rest.remove(Position::new(position)).unwrap();
        }
        let rest = package.publish_ink_edit(rest).unwrap();
        assert!(!has_part(&package, "/payload/handwriting.xml"));
        let removed = save(&mut package);
        let mut package = Package::from_reader(Cursor::new(&removed)).unwrap();
        package.apply_ink_patch(&rest.patch().inverse()).unwrap();
        assert_eq!(members(&save(&mut package)), members(&intermediate));
        package.apply_ink_patch(&first.patch().inverse()).unwrap();
        assert_eq!(members(&save(&mut package)), members(&original));
        assert_eq!(package.ink().unwrap().annotations().len(), 8);
    }
}

#[test]
fn same_id_and_inactive_or_unknown_attribute_uses_prevent_relationship_cleanup() {
    for remaining in [
        "<u:contentPart z:id=\"ink\"/>",
        "<x:opaque xmlns:x=\"urn:producer\" x:refs=\"other #ink tail\"/>",
        "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:producer\"><mc:Choice Requires=\"x\"><u:contentPart z:id=\"ink\"/></mc:Choice><mc:Fallback/></mc:AlternateContent>",
    ] {
        let mut opc = source(false, false);
        let original = format!(
            "<u:document xmlns:u=\"{W}\" xmlns:z=\"{R}\"><u:body><u:p><u:r><u:t>before</u:t><u:contentPart z:id=\"ink\"/><!--keep-->{remaining}<u:t>after</u:t></u:r></u:p></u:body></u:document>"
        );
        opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .set_blob(original.as_bytes().to_vec());
        let mut package = open(&opc);
        let mut edit = package.edit_ink().unwrap();
        edit.remove(Position::new(0)).unwrap();
        package.publish_ink_edit(edit).unwrap();
        assert!(has_part(&package, "/payload/handwriting.xml"));
        let main = package
            .opc_package()
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap();
        assert!(main.rels().get("ink").is_some());
        assert_eq!(
            main.blob(),
            original
                .replacen("<u:contentPart z:id=\"ink\"/>", "", 1)
                .as_bytes()
        );
    }
}

#[test]
fn selectors_skip_generic_content_and_batch_removal_handles_distinct_ids_sharing_a_target() {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:contentPart r:id="math"/><w:contentPart r:id="ink"/><!--retain--><w:contentPart r:id="second"/></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.as_bytes().to_vec());
    part(
        &mut opc,
        "/payload/math.xml",
        "application/mathml+xml",
        b"<math xmlns=\"http://www.w3.org/1998/Math/MathML\"/>",
    );
    edge(
        &mut opc,
        "/word/document.xml",
        "math",
        "../payload/math.xml",
        rt::CUSTOM_XML,
    );
    edge(
        &mut opc,
        "/word/document.xml",
        "second",
        "../payload/handwriting.xml",
        rt::CUSTOM_XML,
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    assert_eq!(edit.snapshot().annotations().len(), 2);
    edit.remove(Position::new(1)).unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert!(!has_part(&package, "/payload/handwriting.xml"));
    assert!(has_part(&package, "/payload/math.xml"));
    let current = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    assert_eq!(
        current.blob(),
        main.replace("<w:contentPart r:id=\"ink\"/>", "")
            .replace("<w:contentPart r:id=\"second\"/>", "")
            .as_bytes()
    );
    assert_eq!(current.rels().len(), 1);
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

pub(super) fn drawing_source(extra_choice: &str, extra_fallback: &str) -> OpcPackage {
    let mut opc = source(false, false);
    let main = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" mc:Ignorable="wi w14"><w:body><w:p><w:r><w:t>before</w:t><mc:AlternateContent><mc:Choice Requires="wi"><w:drawing><wp:inline><wp:extent cx="127000" cy="127000"/><wp:docPr id="1" name="Ink"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"><w14:contentPart r:id="ink"/></a:graphicData></a:graphic></wp:inline></w:drawing>{extra_choice}</mc:Choice><mc:Fallback><w:pict><v:shape id="fallback" style="width:10pt;height:10pt"><v:imagedata r:id="fallbackImage"/></v:shape>{extra_fallback}</w:pict></mc:Fallback></mc:AlternateContent><w:t>after</w:t></w:r></w:p></w:body></w:document>"#
    );
    opc.get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap()
        .set_blob(main.as_bytes().to_vec());
    part(
        &mut opc,
        "/payload/fallback.png",
        "image/png",
        b"opaque image fixture",
    );
    edge(
        &mut opc,
        "/word/document.xml",
        "fallbackImage",
        "../payload/fallback.png",
        rt::IMAGE,
    );
    opc
}

#[test]
fn drawing_removal_deletes_complete_alternative_and_exclusive_fallback_then_restores_members() {
    let opc = drawing_source("", "");
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert!(!has_part(&package, "/payload/fallback.png"));
    assert!(!has_part(&package, "/payload/handwriting.xml"));
    let main = package
        .opc_package()
        .get_part(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    let main = std::str::from_utf8(main.blob()).unwrap();
    assert!(main.contains("<w:t>before</w:t><w:t>after</w:t>"));
    assert!(!main.contains("<mc:AlternateContent>"));
    let mut reopened = Package::from_reader(Cursor::new(save(&mut package))).unwrap();
    reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn removal_and_reopened_inverse_preserve_relationship_and_content_type_lexical_extensions() {
    use soapberry_zip::office::StreamingArchiveWriter;
    let original = PackageWriter::to_bytes(&drawing_source("", "")).unwrap();
    let mut writer = StreamingArchiveWriter::new();
    for (name, bytes) in members(&original) {
        let bytes = if name == "[Content_Types].xml" || name.ends_with(".rels") {
            let xml = String::from_utf8(bytes).unwrap();
            let close = if name.ends_with(".rels") {
                "</Relationships>"
            } else {
                "</Types>"
            };
            xml.replace(
                close,
                &format!("\n<!-- untouched producer spelling &amp; order -->\n{close}"),
            )
            .into_bytes()
        } else {
            bytes
        };
        writer.write_stored(&name, &bytes).unwrap();
    }
    let original = writer.finish_to_bytes().unwrap();
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    let modified = save(&mut package);
    let modified_members = members(&modified);
    for name in ["[Content_Types].xml", "word/_rels/document.xml.rels"] {
        assert!(
            std::str::from_utf8(&modified_members[name])
                .unwrap()
                .contains("<!-- untouched producer spelling &amp; order -->")
        );
    }
    let mut reopened = Package::from_reader(Cursor::new(modified)).unwrap();
    reopened.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut reopened)), members(&original));
}

#[test]
fn drawing_removal_retains_fallback_with_foreign_incoming_and_guards_its_bytes() {
    let mut opc = drawing_source("", "");
    opc.relate_to("payload/fallback.png", "urn:producer:keep-image");
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = edit.commit().unwrap();
    let mut changed = opc.clone();
    part(
        &mut changed,
        "/payload/fallback.png",
        "image/png",
        b"changed image",
    );
    let mut changed = open(&changed);
    let unchanged = save(&mut changed);
    assert!(changed.apply_ink_patch(commit.patch()).is_err());
    assert_eq!(save(&mut changed), unchanged);
    package.apply_ink_patch(commit.patch()).unwrap();
    assert!(has_part(&package, "/payload/fallback.png"));
    assert!(!has_part(&package, "/payload/handwriting.xml"));
    let mut package = Package::from_reader(Cursor::new(save(&mut package))).unwrap();
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

#[test]
fn external_fallback_edge_is_removed_and_undone_without_fetching() {
    let mut opc = drawing_source("", "");
    let main = opc
        .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
        .unwrap();
    main.rels_mut().remove("fallbackImage");
    main.rels_mut().add_relationship(
        rt::IMAGE.to_owned(),
        "https://invalid.example/never-fetch.png".to_owned(),
        "fallbackImage".to_owned(),
        true,
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert!(has_part(&package, "/payload/fallback.png")); // The unrelated orphan stays untouched.
    let mut package = Package::from_reader(Cursor::new(save(&mut package))).unwrap();
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

#[test]
fn unknown_fallback_resource_dependencies_refuse_atomically() {
    let mut opc = drawing_source("", "");
    part(
        &mut opc,
        "/payload/sidecar.bin",
        "application/octet-stream",
        b"sidecar",
    );
    edge(
        &mut opc,
        "/payload/fallback.png",
        "sidecar",
        "sidecar.bin",
        "urn:producer:sidecar",
    );
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    assert!(package.publish_ink_edit(edit).is_err());
    assert_eq!(save(&mut package), original);
}

#[test]
fn unmodeled_drawing_siblings_refuse_without_publication() {
    for (choice, fallback) in [
        ("<w:t>unselected text</w:t>", ""),
        ("unselected text", ""),
        ("<![CDATA[unselected text]]>", ""),
        ("", "<v:shape id=\"unselected\"/>"),
    ] {
        let opc = drawing_source(choice, fallback);
        let original = PackageWriter::to_bytes(&opc).unwrap();
        let mut package = open(&opc);
        let mut edit = package.edit_ink().unwrap();
        edit.remove(Position::new(0)).unwrap();
        assert!(package.publish_ink_edit(edit).is_err());
        assert_eq!(save(&mut package), original);
    }
}

#[test]
fn generic_text_xml_ink_removes_and_restores_its_exact_xml_member() {
    let mut opc = source(false, false);
    part(&mut opc, "/payload/handwriting.xml", "text/xml", INK);
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let commit = package.publish_ink_edit(edit).unwrap();
    assert!(!has_part(&package, "/payload/handwriting.xml"));
    package.apply_ink_patch(&commit.patch().inverse()).unwrap();
    assert_eq!(members(&save(&mut package)), members(&original));
}

#[test]
fn signed_noop_is_exact_changed_edit_requires_explicit_unsigned_source() {
    let mut opc = source(false, false);
    part(
        &mut opc,
        "/_xmlsignatures/origin.sigs",
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        b"",
    );
    opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
    let original = PackageWriter::to_bytes(&opc).unwrap();
    let mut package = open(&opc);
    let noop = package.edit_ink().unwrap().commit().unwrap();
    assert!(!noop.changed());
    package.apply_ink_patch(noop.patch()).unwrap();
    assert_eq!(save(&mut package), original);
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    assert!(package.publish_ink_edit(edit).is_err());
    assert_eq!(save(&mut package), original);
    package.unsign();
    let mut edit = package.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    package.publish_ink_edit(edit).unwrap();
    assert!(!package.is_signed());
}

#[test]
fn stale_graph_and_payload_reject_changed_and_noop_patches_without_touching_target() {
    let mut original_opc = source(false, false);
    part(
        &mut original_opc,
        "/unrelated.bin",
        "application/octet-stream",
        b"original",
    );
    let original = open(&original_opc);
    let noop = original.edit_ink().unwrap().commit().unwrap();
    let mut edit = original.edit_ink().unwrap();
    edit.remove(Position::new(0)).unwrap();
    let changed = edit.commit().unwrap();
    for graph_only in [false, true] {
        let mut opc = original_opc.clone();
        if graph_only {
            opc.relate_to("unrelated.bin", "urn:foreign:edge");
        } else {
            part(
                &mut opc,
                "/unrelated.bin",
                "application/octet-stream",
                b"changed",
            );
        }
        let expected = PackageWriter::to_bytes(&opc).unwrap();
        let mut target = open(&opc);
        assert!(target.apply_ink_patch(noop.patch()).is_err());
        assert!(target.apply_ink_patch(changed.patch()).is_err());
        assert_eq!(save(&mut target), expected);
    }
}

#[test]
fn edit_limits_and_dirty_state_refuse_before_publication() {
    let mut package = open(&source(false, true));
    let original = save(&mut package);
    let mut too_small = package
        .edit_ink_with_limits(EditLimits {
            max_staged_bytes: 1,
            ..EditLimits::default()
        })
        .unwrap();
    too_small.remove(Position::new(0)).unwrap();
    assert!(matches!(
        package.publish_ink_edit(too_small),
        Err(Error::InkLimit {
            resource: "edit staged bytes",
            ..
        })
    ));
    assert_eq!(save(&mut package), original);
    assert!(
        package
            .edit_ink_with_limits(EditLimits {
                max_package_bytes: 1,
                ..EditLimits::default()
            })
            .is_err()
    );
    assert!(
        package
            .edit_ink_with_limits(EditLimits {
                max_operations: 0,
                ..EditLimits::default()
            })
            .is_err()
    );
    let mut edit = package
        .edit_ink_with_limits(EditLimits {
            max_operations: 1,
            ..EditLimits::default()
        })
        .unwrap();
    edit.remove(Position::new(0)).unwrap();
    assert!(edit.remove(Position::new(1)).is_err());
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().annotations().len(), 7);
    package
        .document_mut()
        .unwrap()
        .add_paragraph()
        .add_run_with_text("pending");
    assert!(matches!(package.edit_ink(), Err(Error::UnsafeEdit { .. })));
    assert!(matches!(
        package.apply_ink_patch(commit.patch()),
        Err(Error::UnsafeEdit { .. })
    ));
}
