use litchi_odf_common::package::raw_identical_members;
use litchi_odt::{
    Document,
    core::{OwnedPackage, PackageWriter, Profile},
    package::resource_gc::{
        EmbeddedResourceGcCandidate as Candidate, EmbeddedResourceGcDecision as Decision,
        EmbeddedResourceGcRefusal as Refusal, MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES,
    },
    transaction::Snapshot,
};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";

fn content(body: &str) -> String {
    format!(
        "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xmlns:foo=\"urn:example:unknown\" office:version=\"1.3\"><office:body><office:text>{body}</office:text></office:body></office:document-content>"
    )
}

fn package(body: &str, files: &[(&str, &[u8], &str)], directories: &[(&str, &str)]) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .add_file("content.xml", content(body).as_bytes())
        .unwrap();
    for (path, media_type) in directories {
        writer.add_manifest_directory(path, media_type).unwrap();
    }
    for (path, bytes, media_type) in files {
        writer
            .add_file_with_media_type(path, bytes, media_type)
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn snapshot(bytes: Vec<u8>) -> Snapshot {
    Snapshot::from_bytes(bytes).unwrap()
}

fn decision<'a>(
    plan: &'a litchi_odt::package::resource_gc::EmbeddedResourceGcPlan,
    path: &str,
) -> &'a Decision {
    plan.entries()
        .iter()
        .find(|entry| entry.candidate().path() == path)
        .map(|entry| entry.decision())
        .unwrap()
}

#[test]
fn plans_first_middle_last_orphans_and_retains_a_shared_owner() {
    let body = concat!(
        "<text:p>resources</text:p>",
        "<draw:frame><draw:image xlink:href=\"Pictures/shared.png\"/></draw:frame>",
        "<draw:frame><draw:image xlink:href=\"./Pictures/shared.png\"/></draw:frame>"
    );
    let bytes = package(
        body,
        &[
            ("Pictures/first.png", b"first", "image/png"),
            ("Pictures/shared.png", b"shared", "image/png"),
            ("Pictures/middle.png", b"middle", "image/png"),
            ("Pictures/last.png", b"last", "image/png"),
            ("Thumbnails/thumbnail.png", b"untouched", "image/png"),
        ],
        &[("Pictures/", ""), ("Thumbnails/", "")],
    );
    let source = snapshot(bytes.clone());
    let plan = source
        .plan_embedded_resource_gc(&[
            Candidate::package_file("Pictures/last.png"),
            Candidate::package_file("Pictures/shared.png"),
            Candidate::package_file("Pictures/first.png"),
            Candidate::package_file("Pictures/middle.png"),
        ])
        .unwrap();

    assert_eq!(decision(&plan, "Pictures/first.png"), &Decision::Delete);
    assert_eq!(decision(&plan, "Pictures/middle.png"), &Decision::Delete);
    assert_eq!(decision(&plan, "Pictures/last.png"), &Decision::Delete);
    assert_eq!(
        decision(&plan, "Pictures/shared.png"),
        &Decision::RetainReferenced {
            supported_owner_count: 2
        }
    );

    let commit = plan.apply(&source).unwrap();
    let reopened = commit.snapshot().document().unwrap();
    for path in [
        "Pictures/first.png",
        "Pictures/middle.png",
        "Pictures/last.png",
    ] {
        assert!(reopened.get_file(path).is_err());
    }
    assert_eq!(reopened.get_file("Pictures/shared.png").unwrap(), b"shared");
    assert_eq!(
        reopened.get_file("Thumbnails/thumbnail.png").unwrap(),
        b"untouched"
    );
    let identical = raw_identical_members(&bytes, commit.snapshot().as_bytes()).unwrap();
    for path in [
        "mimetype",
        "content.xml",
        "Pictures/shared.png",
        "Thumbnails/thumbnail.png",
    ] {
        assert!(identical.contains(path), "{path}");
    }
    assert_eq!(
        commit.patch().apply(&source).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        bytes
    );
}

#[test]
fn subdocument_gc_removes_complete_directory_and_manifest_closure() {
    let nested = content("<text:p>nested</text:p>");
    let bytes = package(
        "<text:p>detached object</text:p>",
        &[
            ("Object_1/content.xml", nested.as_bytes(), "text/xml"),
            ("Object_1/Pictures/p.png", b"payload", "image/png"),
            ("Pictures/keep.png", b"keep", "image/png"),
        ],
        &[
            ("Object_1/", MIMETYPE),
            ("Object_1/Pictures/", ""),
            ("Pictures/", ""),
        ],
    );
    let source = snapshot(bytes);
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_subdocument("Object_1/")])
        .unwrap();
    assert_eq!(decision(&plan, "Object_1/"), &Decision::Delete);
    let entry = &plan.entries()[0];
    assert_eq!(
        entry.archive_paths(),
        &["Object_1/Pictures/p.png", "Object_1/content.xml"]
    );
    assert_eq!(
        entry.manifest_paths(),
        &[
            "Object_1/",
            "Object_1/Pictures/",
            "Object_1/Pictures/p.png",
            "Object_1/content.xml"
        ]
    );
    let commit = plan.apply(&source).unwrap();
    let owned = OwnedPackage::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    let package = owned.package().unwrap();
    assert!(!package.has_file("Object_1/content.xml"));
    assert!(!package.manifest().has_path("Object_1/"));
    assert!(package.has_file("Pictures/keep.png"));
}

#[test]
fn unknown_extension_reference_and_unsafe_paths_are_typed_refusals() {
    let bytes = package(
        "<text:p foo:payload=\"Pictures/orphan.png\">unknown owner</text:p>",
        &[("Pictures/orphan.png", b"payload", "image/png")],
        &[("Pictures/", "")],
    );
    let source = snapshot(bytes.clone());
    let plan = source
        .plan_embedded_resource_gc(&[
            Candidate::package_file("Pictures/orphan.png"),
            Candidate::package_file("../../outside.bin"),
        ])
        .unwrap();
    assert!(matches!(
        decision(&plan, "Pictures/orphan.png"),
        Decision::Refuse(Refusal::UnknownReference { part }) if part == "content.xml"
    ));
    assert_eq!(
        decision(&plan, "../../outside.bin"),
        &Decision::Refuse(Refusal::UnsafePath)
    );
    assert!(!plan.is_applicable());
    assert!(plan.apply(&source).is_err());
    assert_eq!(source.as_bytes(), bytes);
}

#[test]
fn stale_apply_exact_noop_and_durable_replay_are_source_checked() {
    let bytes = package(
        "<text:p>orphan</text:p>",
        &[("Pictures/orphan.png", b"payload", "image/png")],
        &[("Pictures/", "")],
    );
    let source = snapshot(bytes.clone());
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_file("Pictures/orphan.png")])
        .unwrap();
    let foreign = snapshot(package("<text:p>foreign</text:p>", &[], &[]));
    assert!(plan.apply(&foreign).is_err());

    let commit = plan.apply(&source).unwrap();
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded = litchi_odt::transaction::DurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(
        decoded.apply(&source).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert!(decoded.apply(&foreign).is_err());
    let decoded_inverse = decoded.inverse();
    assert_eq!(
        decoded_inverse.apply(commit.snapshot()).unwrap().as_bytes(),
        bytes
    );
    assert!(decoded_inverse.apply(&source).is_err());
    assert!(decoded_inverse.apply(&foreign).is_err());

    let noop = source.plan_embedded_resource_gc(&[]).unwrap();
    let noop_commit = noop.apply(&source).unwrap();
    assert_eq!(noop_commit.snapshot().as_bytes(), bytes);
}

#[test]
fn candidate_limit_and_overlaps_refuse_without_partial_deletion() {
    let bytes = package(
        "<text:p>orphan</text:p>",
        &[("Object_1/content.xml", content("").as_bytes(), "text/xml")],
        &[("Object_1/", MIMETYPE)],
    );
    let source = snapshot(bytes.clone());
    let overlap = source
        .plan_embedded_resource_gc(&[
            Candidate::package_subdocument("Object_1/"),
            Candidate::package_file("Object_1/content.xml"),
        ])
        .unwrap();
    assert!(
        overlap
            .entries()
            .iter()
            .all(|entry| { entry.decision() == &Decision::Refuse(Refusal::OverlappingCandidate) })
    );
    assert!(overlap.apply(&source).is_err());
    assert_eq!(source.as_bytes(), bytes);

    let candidates = (0..257)
        .map(|index| Candidate::package_file(format!("Pictures/{index}.png")))
        .collect::<Vec<_>>();
    assert!(source.plan_embedded_resource_gc(&candidates).is_err());
}

#[test]
fn candidate_paths_enforce_the_utf8_byte_bound_before_planning() {
    let source = snapshot(package("<text:p>bounded</text:p>", &[], &[]));
    let exact = "a".repeat(MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES);
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_file(exact.clone())])
        .unwrap();
    assert_eq!(
        decision(&plan, &exact),
        &Decision::Refuse(Refusal::MissingArchiveEntry)
    );

    let above = "é".repeat(MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES / 2 + 1);
    assert_eq!(above.len(), MAX_EMBEDDED_RESOURCE_GC_PATH_BYTES + 2);
    let error = source
        .plan_embedded_resource_gc(&[Candidate::package_file(above)])
        .err()
        .unwrap();
    assert!(
        error
            .to_string()
            .contains("candidate path exceeds 4096 UTF-8 bytes")
    );
}

#[test]
fn object_ole_and_styles_owners_are_part_of_the_shared_reference_closure() {
    let styles = concat!(
        "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" ",
        "xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\">",
        "<office:styles><draw:frame><draw:image xlink:href=\"Pictures/style.png\"/>",
        "</draw:frame></office:styles></office:document-styles>"
    );
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .add_file(
            "content.xml",
            content("<draw:frame><draw:object-ole xlink:href=\"Object_1.bin\"/></draw:frame>")
                .as_bytes(),
        )
        .unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer.add_manifest_directory("Pictures/", "").unwrap();
    writer
        .add_file_with_media_type("Pictures/style.png", b"style", "image/png")
        .unwrap();
    writer
        .add_file_with_media_type("Object_1.bin", b"ole", "application/vnd.sun.star.oleobject")
        .unwrap();
    let source = snapshot(writer.finish_to_bytes().unwrap());
    let plan = source
        .plan_embedded_resource_gc(&[
            Candidate::package_file("Pictures/style.png"),
            Candidate::package_file("Object_1.bin"),
        ])
        .unwrap();
    for path in ["Pictures/style.png", "Object_1.bin"] {
        assert_eq!(
            decision(&plan, path),
            &Decision::RetainReferenced {
                supported_owner_count: 1
            }
        );
    }
    assert_eq!(
        plan.apply(&source).unwrap().snapshot().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn signatures_and_document_protection_are_typed_refusals() {
    const SETTINGS: &[u8] = br#"<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings><config:config-item-set config:name="ooo:configuration-settings"><config:config-item config:name="LoadReadonly" config:type="boolean">true</config:config-item></config:config-item-set></office:settings></office:document-settings>"#;
    let build = |protected: bool, signed: bool| {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(MIMETYPE).unwrap();
        writer
            .add_file("content.xml", content("<text:p>orphan</text:p>").as_bytes())
            .unwrap();
        writer.add_manifest_directory("Pictures/", "").unwrap();
        writer
            .add_file_with_media_type("Pictures/orphan.png", b"payload", "image/png")
            .unwrap();
        if protected {
            writer.add_file("settings.xml", SETTINGS).unwrap();
        }
        if signed {
            writer
                .add_file(
                    "META-INF/documentsignatures.xml",
                    br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"/>"#,
                )
                .unwrap();
        }
        snapshot(writer.finish_to_bytes().unwrap())
    };
    let candidate = [Candidate::package_file("Pictures/orphan.png")];
    let protected = build(true, false);
    assert_eq!(
        protected
            .plan_embedded_resource_gc(&candidate)
            .unwrap()
            .entries()[0]
            .decision(),
        &Decision::Refuse(Refusal::ProtectedDocument)
    );
    let signed = build(false, true);
    assert_eq!(
        signed
            .plan_embedded_resource_gc(&candidate)
            .unwrap()
            .entries()[0]
            .decision(),
        &Decision::Refuse(Refusal::SignedPackage)
    );
    let signed_noop = signed
        .plan_embedded_resource_gc(&[])
        .unwrap()
        .apply(&signed)
        .unwrap();
    assert_eq!(signed_noop.snapshot().as_bytes(), signed.as_bytes());
    let durable = signed_noop.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&signed).unwrap().as_bytes(),
        signed.as_bytes()
    );
}

#[test]
fn live_selected_subdocument_cannot_hide_a_reference_to_another_candidate() {
    let nested =
        content("<draw:frame><draw:image xlink:href=\"Pictures/nested.png\"/></draw:frame>");
    let bytes = package(
        "<draw:frame><draw:object xlink:href=\"Object_1\"/></draw:frame>",
        &[
            ("Object_1/content.xml", nested.as_bytes(), "text/xml"),
            ("Pictures/nested.png", b"nested", "image/png"),
        ],
        &[("Object_1/", MIMETYPE), ("Pictures/", "")],
    );
    let source = snapshot(bytes);
    let plan = source
        .plan_embedded_resource_gc(&[
            Candidate::package_subdocument("Object_1/"),
            Candidate::package_file("Pictures/nested.png"),
        ])
        .unwrap();
    assert_eq!(
        decision(&plan, "Object_1/"),
        &Decision::RetainReferenced {
            supported_owner_count: 1
        }
    );
    assert!(matches!(
        decision(&plan, "Pictures/nested.png"),
        Decision::Refuse(Refusal::UnknownReference { part }) if part == "Object_1/content.xml"
    ));
    assert!(!plan.is_applicable());
    assert!(plan.apply(&source).is_err());
}

#[test]
fn inherited_xml_base_refuses_deletion_instead_of_misresolving_a_live_owner() {
    let body = concat!(
        "<draw:frame xml:base=\"Pictures/\">",
        "<draw:image xlink:href=\"based.png\"/>",
        "</draw:frame>"
    );
    let bytes = package(
        body,
        &[("Pictures/based.png", b"based", "image/png")],
        &[("Pictures/", "")],
    );
    let source = snapshot(bytes);
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_file("Pictures/based.png")])
        .unwrap();
    assert!(matches!(
        decision(&plan, "Pictures/based.png"),
        Decision::Refuse(Refusal::UnknownReference { part }) if part == "content.xml"
    ));
    assert!(plan.apply(&source).is_err());
}

#[test]
fn nested_xml_references_are_resolved_relative_to_their_package_part() {
    let nested =
        content("<draw:frame><draw:image xlink:href=\"Pictures/local.png\"/></draw:frame>");
    let bytes = package(
        "<draw:frame><draw:object xlink:href=\"Object_1\"/></draw:frame>",
        &[
            ("Object_1/content.xml", nested.as_bytes(), "text/xml"),
            ("Object_1/Pictures/local.png", b"local", "image/png"),
        ],
        &[("Object_1/", MIMETYPE), ("Object_1/Pictures/", "")],
    );
    let source = snapshot(bytes);
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_file("Object_1/Pictures/local.png")])
        .unwrap();
    assert!(matches!(
        decision(&plan, "Object_1/Pictures/local.png"),
        Decision::Refuse(Refusal::UnknownReference { part }) if part == "Object_1/content.xml"
    ));
    assert!(plan.apply(&source).is_err());
}

#[test]
fn encrypted_snapshot_planning_returns_a_typed_refusal_without_decryption() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIMETYPE).unwrap();
    writer
        .set_encryption("secret", Profile::compatible())
        .unwrap();
    writer
        .add_file(
            "content.xml",
            content("<text:p>encrypted</text:p>").as_bytes(),
        )
        .unwrap();
    writer.add_manifest_directory("Pictures/", "").unwrap();
    writer
        .add_file_with_media_type("Pictures/orphan.png", b"payload", "image/png")
        .unwrap();
    let bytes = writer.finish_to_bytes().unwrap();
    let document = Document::from_bytes_with_password(bytes, "secret").unwrap();
    let source = document.snapshot().unwrap();
    let plan = source
        .plan_embedded_resource_gc(&[Candidate::package_file("Pictures/orphan.png")])
        .unwrap();
    assert_eq!(
        decision(&plan, "Pictures/orphan.png"),
        &Decision::Refuse(Refusal::EncryptedPackage)
    );
    assert!(plan.apply(&source).is_err());
}
