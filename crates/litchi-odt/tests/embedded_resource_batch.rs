use litchi_odf_common::package::raw_identical_members;
use litchi_odt::{
    Document,
    mutable::MutableDocument,
    package::embedded::{
        EmbeddedResource, EmbeddedResourceChange, EmbeddedResourceFile, EmbeddedResourceKind,
        EmbeddedResourceSource,
    },
    transaction::{OperationResult, Position},
};

fn source() -> litchi_odt::transaction::Snapshot {
    let mut document = MutableDocument::new();
    document.add_paragraph("resource batch").unwrap();
    Document::from_bytes(document.to_bytes().unwrap())
        .unwrap()
        .snapshot()
        .unwrap()
}

fn image(name: &str, path: &str, bytes: &[u8]) -> EmbeddedResource {
    EmbeddedResource {
        kind: EmbeddedResourceKind::Image,
        source: EmbeddedResourceSource::PackageFile {
            bytes: bytes.to_vec(),
            media_type: "image/png".to_string(),
            preferred_path: Some(path.to_string()),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    }
}

fn object(name: &str, root: &str) -> EmbeddedResource {
    EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::PackageSubdocument {
            files: vec![EmbeddedResourceFile {
                path: "content.xml".to_string(),
                bytes: br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#.to_vec(),
                media_type: "text/xml".to_string(),
            }],
            media_type: "application/vnd.oasis.opendocument.text".to_string(),
            preferred_root: Some(root.to_string()),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    }
}

fn linked_object(name: &str) -> EmbeddedResource {
    EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::Linked {
            href: format!("https://example.invalid/{name}"),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    }
}

fn package_object(name: &str, path: &str) -> EmbeddedResource {
    EmbeddedResource {
        kind: EmbeddedResourceKind::Object,
        source: EmbeddedResourceSource::PackageFile {
            bytes: b"opaque object".to_vec(),
            media_type: "application/octet-stream".to_string(),
            preferred_path: Some(path.to_string()),
        },
        frame_name: Some(name.to_string()),
        xml_id: None,
        class_id: None,
    }
}

fn with_content(
    snapshot: &litchi_odt::transaction::Snapshot,
    content: &str,
) -> litchi_odt::transaction::Snapshot {
    let package = litchi_odt::core::OwnedPackage::from_bytes(snapshot.as_bytes().to_vec()).unwrap();
    let bytes = litchi_odf_common::package::replace_content_xml(&package, content).unwrap();
    Document::from_bytes(bytes).unwrap().snapshot().unwrap()
}

#[test]
fn batch_publishes_media_rich_create_update_delete_once_and_replays_durably() {
    let base = source();
    let additions = [
        EmbeddedResourceChange::add(&image("first", "Pictures/first.png", b"first")),
        EmbeddedResourceChange::add(&image("middle", "Pictures/middle.png", b"middle")),
        EmbeddedResourceChange::add(&image("last", "Pictures/last.png", b"last")),
        EmbeddedResourceChange::add(&object("object", "Objects/Writer")),
    ];
    let mut setup = base.edit();
    setup.edit_embedded_resources(&additions).unwrap();
    let setup = setup.commit().unwrap();
    assert_eq!(
        setup.results(),
        &[OperationResult::Indices(vec![0, 1, 2, 0])]
    );
    let setup = setup.into_snapshot();
    let setup_document = setup.document().unwrap();
    assert_eq!(setup_document.images().unwrap().len(), 3);
    assert_eq!(setup_document.embedded_objects().unwrap().len(), 1);

    let changes = [
        EmbeddedResourceChange::remove_image(Position::new(0)),
        EmbeddedResourceChange::replace_image(
            Position::new(1),
            &image("middle", "Pictures/middle.png", b"replacement"),
        ),
        EmbeddedResourceChange::remove_image(Position::new(2)),
        EmbeddedResourceChange::replace_object(Position::new(0), &linked_object("object")),
        EmbeddedResourceChange::add(&image("new", "Pictures/new.png", b"new")),
    ];
    let mut edit = setup.edit();
    edit.edit_embedded_resources(&changes).unwrap();
    let committed = edit.commit().unwrap();
    assert_eq!(committed.results(), &[OperationResult::Indices(vec![1])]);
    let reopened = committed.snapshot().document().unwrap();
    assert_eq!(reopened.images().unwrap().len(), 2);
    assert_eq!(reopened.embedded_objects().unwrap().len(), 1);
    assert_eq!(
        reopened.get_file("Pictures/middle.png").unwrap(),
        b"replacement"
    );
    assert_eq!(reopened.get_file("Pictures/new.png").unwrap(), b"new");
    assert_eq!(reopened.get_file("Pictures/first.png").unwrap(), b"first");
    assert_eq!(reopened.get_file("Pictures/last.png").unwrap(), b"last");
    assert!(reopened.get_file("Objects/Writer/content.xml").is_ok());

    let durable = committed.patch().durable().unwrap();
    let json = durable.to_deterministic_json().unwrap();
    assert!(
        std::str::from_utf8(&json)
            .unwrap()
            .contains("resource.embedded.batch")
    );
    let replayed = litchi_odt::transaction::DurablePatch::from_deterministic_json(&json)
        .unwrap()
        .apply(&setup)
        .unwrap();
    assert_eq!(replayed.as_bytes(), committed.snapshot().as_bytes());
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(committed.snapshot())
            .unwrap()
            .as_bytes(),
        setup.as_bytes()
    );
    assert!(committed.patch().apply(&base).is_err());

    let identical = raw_identical_members(setup.as_bytes(), committed.snapshot().as_bytes())
        .expect("ordinary ZIP preservation audit");
    for untouched in ["mimetype", "styles.xml", "meta.xml"] {
        assert!(identical.contains(untouched), "{untouched}");
    }
}

#[test]
fn batch_preflight_is_bounded_duplicate_safe_and_atomic_on_late_failure() {
    let source = source();
    let mut setup = source.edit();
    setup
        .edit_embedded_resources(&[
            EmbeddedResourceChange::add(&image("one", "Pictures/one.png", b"one")),
            EmbeddedResourceChange::add(&image("two", "Pictures/two.png", b"two")),
        ])
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();

    let mut semantic_noop = setup.edit();
    semantic_noop
        .edit_embedded_resources(&[EmbeddedResourceChange::replace_image(
            Position::new(0),
            &image("one", "Pictures/one.png", b"one"),
        )])
        .unwrap();
    assert_eq!(
        semantic_noop.commit().unwrap().snapshot().as_bytes(),
        setup.as_bytes()
    );

    let duplicate = [
        EmbeddedResourceChange::remove_image(Position::new(0)),
        EmbeddedResourceChange::replace_image(
            Position::new(0),
            &image("again", "Pictures/again.png", b"again"),
        ),
    ];
    let mut duplicate_edit = setup.edit();
    duplicate_edit.edit_embedded_resources(&duplicate).unwrap();
    assert!(duplicate_edit.commit().is_err());

    let late_failure = [
        EmbeddedResourceChange::replace_image(
            Position::new(0),
            &image("valid", "Pictures/one.png", b"changed"),
        ),
        EmbeddedResourceChange::add(&image("unsafe", "../escape.png", b"unsafe")),
    ];
    let mut failed = setup.edit();
    failed.edit_embedded_resources(&late_failure).unwrap();
    assert!(failed.commit().is_err());
    assert_eq!(
        setup
            .document()
            .unwrap()
            .get_file("Pictures/one.png")
            .unwrap(),
        b"one"
    );

    let changes = vec![EmbeddedResourceChange::add(&linked_object("bounded")); 257];
    let mut oversized = setup.edit();
    assert!(oversized.edit_embedded_resources(&changes).is_err());

    let before = setup.as_bytes().to_vec();
    let mut exact_noop = setup.edit();
    exact_noop.edit_embedded_resources(&[]).unwrap();
    assert_eq!(exact_noop.commit().unwrap().snapshot().as_bytes(), before);
}

#[test]
fn batch_removes_first_middle_last_and_only_owners_without_leaving_empty_frames() {
    for selected in 0..3 {
        let source = source();
        let mut setup = source.edit();
        setup
            .edit_embedded_resources(&[
                EmbeddedResourceChange::add(&image("first", "Pictures/first.png", b"first")),
                EmbeddedResourceChange::add(&image("middle", "Pictures/middle.png", b"middle")),
                EmbeddedResourceChange::add(&image("last", "Pictures/last.png", b"last")),
            ])
            .unwrap();
        let setup = setup.commit().unwrap().into_snapshot();
        let mut edit = setup.edit();
        edit.edit_embedded_resources(&[EmbeddedResourceChange::remove_image(Position::new(
            selected,
        ))])
        .unwrap();
        let committed = edit.commit().unwrap();
        let document = committed.snapshot().document().unwrap();
        assert_eq!(document.images().unwrap().len(), 2);
        let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
        assert_eq!(content.matches("<draw:frame").count(), 2);
    }

    let source = source();
    let mut setup = source.edit();
    setup
        .edit_embedded_resources(&[EmbeddedResourceChange::add(&image(
            "only",
            "Pictures/only.png",
            b"only",
        ))])
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();
    let mut edit = setup.edit();
    edit.edit_embedded_resources(&[EmbeddedResourceChange::remove_image(Position::new(0))])
        .unwrap();
    let committed = edit.commit().unwrap();
    let document = committed.snapshot().document().unwrap();
    assert!(document.images().unwrap().is_empty());
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    assert!(!content.contains("<draw:frame"));
    // Retain is the default: resource deletion never silently performs GC.
    assert_eq!(document.get_file("Pictures/only.png").unwrap(), b"only");
}

#[test]
fn replacement_preserves_frame_bytes_and_removal_retains_unknown_frame_children() {
    let source = source();
    let mut setup = source.edit();
    setup
        .edit_embedded_resources(&[EmbeddedResourceChange::add(&image(
            "owned",
            "Pictures/owned.png",
            b"before",
        ))])
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();
    let document = setup.document().unwrap();
    let mut content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    content = content.replacen("<draw:frame ", "<draw:frame draw:z-index=\"9\" ", 1);
    let close = content.find("</draw:frame>").unwrap();
    content.insert_str(close, "<draw:glue-point draw:id=\"7\"/>");
    let setup = with_content(&setup, &content);

    let mut replacement = setup.edit();
    replacement
        .edit_embedded_resources(&[EmbeddedResourceChange::replace_image(
            Position::new(0),
            &image("requested different frame", "Pictures/owned.png", b"after"),
        )])
        .unwrap();
    let replacement = replacement.commit().unwrap().into_snapshot();
    let document = replacement.document().unwrap();
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    assert_eq!(content.matches("<draw:frame").count(), 1);
    assert!(content.contains("draw:z-index=\"9\""));
    assert!(content.contains("draw:name=\"owned\""));
    assert!(!content.contains("requested different frame"));
    assert!(content.contains("<draw:glue-point draw:id=\"7\"/>"));
    assert_eq!(document.get_file("Pictures/owned.png").unwrap(), b"after");

    let mut removal = replacement.edit();
    removal
        .edit_embedded_resources(&[EmbeddedResourceChange::remove_image(Position::new(0))])
        .unwrap();
    let removed = removal.commit().unwrap().into_snapshot();
    let document = removed.document().unwrap();
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    assert!(document.images().unwrap().is_empty());
    assert_eq!(content.matches("<draw:frame").count(), 1);
    assert!(content.contains("draw:z-index=\"9\""));
    assert!(content.contains("<draw:glue-point draw:id=\"7\"/>"));
    assert_eq!(document.get_file("Pictures/owned.png").unwrap(), b"after");
}

#[test]
fn planned_same_frame_removals_collapse_only_after_the_complete_owner_set_is_empty() {
    let source = source();
    let mut setup = source.edit();
    setup
        .edit_embedded_resources(&[
            EmbeddedResourceChange::add(&image("one", "Pictures/one.png", b"one")),
            EmbeddedResourceChange::add(&image("two", "Pictures/two.png", b"two")),
        ])
        .unwrap();
    let setup = setup.commit().unwrap().into_snapshot();
    let document = setup.document().unwrap();
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    let first_close = content.find("</draw:frame>").unwrap();
    let second_start = content[first_close + "</draw:frame>".len()..]
        .find("<draw:frame")
        .map(|offset| first_close + "</draw:frame>".len() + offset)
        .unwrap();
    let second_start_end = content[second_start..]
        .find('>')
        .map(|offset| second_start + offset + 1)
        .unwrap();
    let mut merged = content;
    merged.replace_range(first_close..second_start_end, "");
    let setup = with_content(&setup, &merged);

    let mut removal = setup.edit();
    removal
        .edit_embedded_resources(&[
            EmbeddedResourceChange::remove_image(Position::new(0)),
            EmbeddedResourceChange::remove_image(Position::new(1)),
        ])
        .unwrap();
    let removed = removal.commit().unwrap().into_snapshot();
    let document = removed.document().unwrap();
    assert!(document.images().unwrap().is_empty());
    let content = String::from_utf8(document.get_file("content.xml").unwrap()).unwrap();
    assert!(!content.contains("<draw:frame"));
    assert_eq!(document.get_file("Pictures/one.png").unwrap(), b"one");
    assert_eq!(document.get_file("Pictures/two.png").unwrap(), b"two");
}

#[test]
fn batch_rejects_file_directory_prefix_collisions_in_both_orders() {
    let source = source();
    for changes in [
        vec![
            EmbeddedResourceChange::add(&package_object("file", "Collision")),
            EmbeddedResourceChange::add(&object("directory", "Collision")),
        ],
        vec![
            EmbeddedResourceChange::add(&object("directory", "Collision")),
            EmbeddedResourceChange::add(&package_object("file", "Collision")),
        ],
    ] {
        let mut edit = source.edit();
        edit.edit_embedded_resources(&changes).unwrap();
        assert!(edit.commit().is_err());
    }
    assert_eq!(
        source.document().unwrap().embedded_objects().unwrap().len(),
        0
    );
}
