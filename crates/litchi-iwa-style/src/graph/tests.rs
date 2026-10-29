use litchi_iwa_core::{Archive, ArchiveObject, Limits, RawMessage};

use super::Workspace;

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
        Limits::default(),
    )
    .expect("valid object")
}

fn components() -> Vec<(String, Archive)> {
    vec![
        (
            "Index/Document.iwa".to_owned(),
            Archive {
                objects: vec![object(1, 1, vec![0x08, 0x01]), object(7, 2, Vec::new())],
            },
        ),
        (
            "Index/Metadata.iwa".to_owned(),
            Archive {
                // last_object_identifier = 9.
                objects: vec![object(2, 11_006, vec![0x08, 0x09])],
            },
        ),
    ]
}

#[test]
fn objects_are_copy_on_write_and_identifiers_are_fresh() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    assert_eq!(workspace.next_object_identifier().expect("next"), 10);
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(9));
    workspace.remove(7).expect("removes");
    assert!(workspace.object(7).is_err());
    let changes = workspace.finish();
    assert_eq!(changes.edits.len(), 1);
    assert_eq!(changes.edits[0].0, "Index/Document.iwa");
    assert!(source[0].1.object(7).is_some());
}

#[test]
fn the_watermark_is_raised_and_lowered() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    workspace.set_last_object_identifier(12).expect("raises");
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(12));
    workspace.set_last_object_identifier(7).expect("lowers");
    assert_eq!(workspace.last_object_identifier().expect("last"), Some(7));
    // Object 7 still exists, so releasing it must fail closed.
    let mut again = Workspace::new(&source, Limits::default()).expect("workspace");
    again.set_last_object_identifier(7).expect("lowers");
    assert!(again.release_identifier_suffix(&[7]).is_err());
}

#[test]
fn a_package_without_registry_skips_registry_edits() {
    let source = vec![components().remove(0)];
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    assert!(!workspace.has_registry());
    workspace.add_object_uuids(1, &[7]).expect("no-op");
    workspace.set_last_object_identifier(40).expect("no-op");
    assert!(workspace.finish().is_empty());
}

#[test]
fn registry_inspection_work_only_charges_an_uncached_state() {
    let source = components();
    let mut workspace = Workspace::new(&source, Limits::default()).expect("workspace");
    let initial = workspace
        .registry_inspection_requirements()
        .expect("requirements")
        .expect("registry")
        .work_bytes();

    // Preparing the report is reused, but the first lookup still needs the
    // actual two-pass inspection that populates the root cache.
    assert_eq!(workspace.registry_inspection_work(), initial);
    assert_eq!(
        workspace
            .component_for_member("Index/Document.iwa")
            .expect("component lookup"),
        None
    );
    assert_eq!(workspace.registry_inspection_work(), 0);

    // A metadata rewrite invalidates both caches; the next lookup must be
    // admitted again for the new payload rather than reusing stale facts.
    workspace
        .set_last_object_identifier(10)
        .expect("watermark rewrite");
    assert!(workspace.registry_inspection_work() > 0);
}
