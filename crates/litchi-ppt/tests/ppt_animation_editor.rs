#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_core::edit::RefusalKind;
use litchi_ppt::Package;
use litchi_ppt::animation::{self, ExtendedTimeNode, Scope, TimeNodeAtom, TimeNodeKind};

/// A built package with one slide and its main master, and their persist IDs.
fn built() -> (Package, u32, u32) {
    let mut builder = litchi_ppt::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 10, 10, 240, 40, "Animated")
        .unwrap();
    let package = builder.build().unwrap();
    let edit = package.edit_animations().unwrap();
    let timelines = edit.timelines();
    let slide = timelines
        .iter()
        .find(|timeline| timeline.scope == Scope::Slide)
        .unwrap()
        .persist_id;
    let master = timelines
        .iter()
        .find(|timeline| timeline.scope == Scope::MainMaster)
        .unwrap()
        .persist_id;
    drop(edit);
    (package, slide, master)
}

fn node(kind: TimeNodeKind, duration_ms: i32) -> ExtendedTimeNode {
    ExtendedTimeNode {
        atom: TimeNodeAtom {
            node_type: Some(kind),
            duration_ms: Some(duration_ms),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn root_children(edit: &animation::Edit, persist_id: u32) -> usize {
    edit.find(persist_id)
        .unwrap()
        .extension
        .time_node
        .map_or(0, |root| root.children.len())
}

#[test]
fn built_timeline_add_update_reorder_remove_and_reopen() {
    let (source, slide, master) = built();
    let source_bytes = source.to_bytes().unwrap();
    let mut edit = source.edit_animations().unwrap();
    assert_eq!(edit.find(slide).unwrap().scope, Scope::Slide);
    assert_eq!(edit.find(master).unwrap().scope, Scope::MainMaster);
    let first = node(TimeNodeKind::Parallel, 100);
    let second = node(TimeNodeKind::Sequential, 200);
    edit.add(slide, 0, first).unwrap();
    edit.add(slide, 1, second.clone()).unwrap();
    edit.reorder(slide, &[1, 0]).unwrap();
    edit.add(master, 0, ExtendedTimeNode::default()).unwrap();
    edit.update(slide, 0, second).unwrap();
    edit.remove(slide, 1).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );

    let reopened = commit.package().edit_animations().unwrap();
    assert_eq!(root_children(&reopened, slide), 1);
    assert_eq!(root_children(&reopened, master), 1);
    drop(reopened);

    let replayed = source.apply_animations(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_animations(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_animations(commit.patch()),
        Err(animation::Error::PatchConflict)
    ));
}

#[test]
fn an_unstaged_or_dropped_edit_leaves_the_package_unchanged() {
    let (source, slide, _master) = built();
    let source_bytes = source.to_bytes().unwrap();
    let no_op = source.edit_animations().unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_animations().unwrap();
    edit.add(slide, 0, node(TimeNodeKind::Parallel, 100))
        .unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}

#[test]
fn invalid_indexes_and_limits_are_atomic() {
    let (source, slide, _master) = built();
    let mut edit = source.edit_animations().unwrap();
    let before = edit.timelines();
    assert!(matches!(
        edit.add(slide, 1, ExtendedTimeNode::default()),
        Err(animation::Error::Refused(_))
    ));
    assert!(edit.reorder(slide, &[0]).is_err());
    assert_eq!(edit.timelines(), before);
    let Err(animation::Error::Refused(refusal)) =
        edit.add(u32::MAX, 0, ExtendedTimeNode::default())
    else {
        panic!("an unknown persist ID must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    let Err(animation::Error::Refused(refusal)) =
        edit.replace_shape_animation(slide, u32::MAX, None)
    else {
        panic!("an unknown shape must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    assert_eq!(edit.timelines(), before);
}

#[test]
fn poi_animation_fixtures_are_strictly_gated() {
    let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for path in [
        root.join("test-data/poi/test-data/slideshow/sound.ppt"),
        root.join("test-data/poi/test-data/slideshow/datetime.ppt"),
    ] {
        let original = std::fs::read(&path).unwrap();
        let Ok(package) = Package::from_bytes(&original) else {
            continue;
        };
        match package.edit_animations() {
            Ok(edit) => {
                let _timelines = edit.timelines();
                let commit = edit.commit().unwrap();
                assert!(!commit.diagnostics().changed());
                assert_eq!(
                    commit.package().to_bytes().unwrap(),
                    package.to_bytes().unwrap()
                );
            },
            Err(_) => assert_eq!(std::fs::read(&path).unwrap(), original),
        }
    }
}

fn with_signature(source: &[u8]) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};
    use std::io::Cursor;

    let mut source_file = OleFile::open(Cursor::new(source.to_vec())).unwrap();
    let mut writer = OleWriter::new();
    for stream_path in source_file.list_streams() {
        let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
        let data = source_file.open_stream(&refs).unwrap();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer
        .create_stream(&["_xmlsignatures", "origin.sigs"], b"signature")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn a_signed_package_refuses_animation_edits() {
    let (source, _slide, _master) = built();
    let signed = Package::from_bytes(&with_signature(&source.to_bytes().unwrap())).unwrap();
    // Every staged animation change would invalidate the signature, and the
    // record editor cannot stage over signature storages, so the edit is
    // refused up front.
    assert!(matches!(
        signed.edit_animations(),
        Err(animation::Error::Signed)
    ));
}
