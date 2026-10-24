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
    let timelines = edit.timelines().unwrap();
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
    assert_eq!(edit.find(slide).unwrap().unwrap().scope, Scope::Slide);
    assert_eq!(edit.find(master).unwrap().unwrap().scope, Scope::MainMaster);
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
    let before = edit.timelines().unwrap();
    assert!(matches!(
        edit.add(slide, 1, ExtendedTimeNode::default()),
        Err(animation::Error::Refused(_))
    ));
    assert!(edit.reorder(slide, &[0]).is_err());
    assert_eq!(edit.timelines().unwrap(), before);
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
    assert_eq!(edit.timelines().unwrap(), before);
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
                let _timelines = match edit.timelines() {
                    Ok(timelines) => timelines,
                    Err(animation::Error::Refused(refusal)) => {
                        assert_eq!(refusal.kind(), RefusalKind::UnsupportedContent);
                        assert_eq!(std::fs::read(&path).unwrap(), original);
                        continue;
                    },
                    Err(error) => panic!("{path:?}: animation projection failed: {error:?}"),
                };
                let commit = edit
                    .commit()
                    .unwrap_or_else(|error| panic!("{path:?}: animation commit failed: {error:?}"));
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

    let mut source_file = OleFile::open(
        Cursor::new(source.to_vec()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
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
fn a_signed_package_admits_animation_edits_and_refuses_only_a_changed_commit() {
    let (source, slide, _master) = built();
    let signed_bytes = with_signature(&source.to_bytes().unwrap());
    let signed = Package::from_bytes(&signed_bytes).unwrap();

    // Staging and reading work; an exact no-op commit keeps the signature.
    let edit = signed.edit_animations().unwrap();
    assert_eq!(edit.find(slide).unwrap().unwrap().scope, Scope::Slide);
    let no_op = edit.commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), signed_bytes);

    // A change that is reverted before the commit is a no-op too.
    let original = signed
        .edit_animations()
        .unwrap()
        .find(slide)
        .unwrap()
        .unwrap();
    let mut edit = signed.edit_animations().unwrap();
    edit.add(slide, 0, node(TimeNodeKind::Parallel, 100))
        .unwrap();
    edit.replace(
        slide,
        original.extension.time_node,
        original.extension.build_list,
    )
    .unwrap();
    let reverted = edit.commit().unwrap();
    assert!(!reverted.diagnostics().changed());
    assert_eq!(reverted.package().to_bytes().unwrap(), signed_bytes);

    // A changed commit would invalidate the signature.
    let mut edit = signed.edit_animations().unwrap();
    edit.add(slide, 0, node(TimeNodeKind::Parallel, 100))
        .unwrap();
    assert!(matches!(edit.commit(), Err(animation::Error::Signed)));
    assert_eq!(signed.to_bytes().unwrap(), signed_bytes);
}

#[test]
fn a_change_reverted_before_the_commit_is_an_exact_no_op() {
    let (built, slide, master) = built();
    // A source whose slide already carries a timeline with one child.
    let mut edit = built.edit_animations().unwrap();
    edit.add(slide, 0, node(TimeNodeKind::Parallel, 100))
        .unwrap();
    let source = edit.commit().unwrap().package().clone();
    let source_bytes = source.to_bytes().unwrap();
    let staged = |package: &Package, persist_id| {
        package
            .edit_animations()
            .unwrap()
            .find(persist_id)
            .unwrap()
            .unwrap()
    };
    let original = staged(&source, slide);
    let original_master = staged(&source, master);
    let assert_no_op = |edit: animation::Edit| {
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(commit.patch().changed_bytes(), 0);
        assert_eq!(commit.package().to_bytes().unwrap(), source_bytes);
    };

    // Update a node and update it back.
    let mut edit = source.edit_animations().unwrap();
    let first = original.extension.time_node.as_ref().unwrap().children[0].clone();
    edit.update(slide, 0, node(TimeNodeKind::Sequential, 999))
        .unwrap();
    edit.update(slide, 0, first.clone()).unwrap();
    assert_eq!(edit.find(slide).unwrap().unwrap(), original);
    assert_no_op(edit);

    // Add, reorder and remove back to the source order.
    let mut edit = source.edit_animations().unwrap();
    edit.add(slide, 1, node(TimeNodeKind::Sequential, 200))
        .unwrap();
    edit.reorder(slide, &[1, 0]).unwrap();
    edit.remove(slide, 0).unwrap();
    assert_eq!(edit.find(slide).unwrap().unwrap(), original);
    assert_no_op(edit);

    // Give the master a timeline, then restore its source timeline: the
    // record is rewritten back to its source meaning and keeps its bytes.
    let mut edit = source.edit_animations().unwrap();
    edit.add(master, 0, node(TimeNodeKind::Sequential, 200))
        .unwrap();
    edit.replace(
        master,
        original_master.extension.time_node.clone(),
        original_master.extension.build_list.clone(),
    )
    .unwrap();
    assert_eq!(edit.find(master).unwrap().unwrap(), original_master);
    assert_no_op(edit);

    // A partial revert publishes only the record that still changed: the
    // master keeps its source timeline.
    let mut edit = source.edit_animations().unwrap();
    edit.add(slide, 1, node(TimeNodeKind::Sequential, 200))
        .unwrap();
    edit.add(master, 0, node(TimeNodeKind::Sequential, 200))
        .unwrap();
    edit.replace(
        master,
        original_master.extension.time_node.clone(),
        original_master.extension.build_list.clone(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        root_children(&commit.package().edit_animations().unwrap(), slide),
        2
    );
    assert_eq!(staged(commit.package(), master), original_master);
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
}
