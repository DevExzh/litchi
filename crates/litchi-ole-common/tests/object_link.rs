#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "integration tests use concise assertions and checked fixture-sized literals"
)]

use litchi_cfb::OleWriter;
use litchi_ole_common::Error;
use litchi_ole_common::object::link::{Kind, Limits as LinkLimits};
use litchi_ole_common::object::{Edit, Limits, OpenOptions, Package, Target, Targets};
use std::io::Cursor;

fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
    let mut writer = OleWriter::new();
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("test CFB should write");
    output.into_inner()
}

fn linked_wire() -> Vec<u8> {
    let mut output = Vec::new();
    for value in [0x0200_0001u32, 0x1001, 7, 0, 0] {
        output.extend_from_slice(&value.to_le_bytes());
    }
    output.extend_from_slice(&20u32.to_le_bytes());
    output.extend_from_slice(&[0x10; 16]);
    output.extend_from_slice(&20u32.to_le_bytes());
    output.extend_from_slice(&[0x20; 16]);
    output.extend_from_slice(&u32::MAX.to_le_bytes());
    output.extend_from_slice(&[0x30; 16]);
    output.extend_from_slice(&4u32.to_le_bytes());
    output.extend_from_slice(&[0; 4]);
    output.extend_from_slice(&0xDEAD_BEEFu32.to_le_bytes());
    for value in [11u64, 22, 33] {
        output.extend_from_slice(&value.to_le_bytes());
    }
    output.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
    output
}

fn target() -> Targets {
    Targets::one(Target::new("object", ["ObjectPool", "_1"]).expect("target should validate"))
}

fn open_editor(bytes: Vec<u8>, targets: Targets, limits: Limits) -> Result<Edit, Error> {
    Package::open_with_options(bytes, OpenOptions::new().targets(targets).limits(limits))?.edit()
}

fn staged(editor: &Edit) -> Package {
    editor
        .clone()
        .commit()
        .expect("staged package")
        .into_package()
}

trait RenderBytes {
    fn render_bytes(&self) -> Result<Vec<u8>, Error>;
}

impl RenderBytes for Edit {
    fn render_bytes(&self) -> Result<Vec<u8>, Error> {
        let package = self.clone().commit()?.into_package();
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

impl RenderBytes for Package {
    fn render_bytes(&self) -> Result<Vec<u8>, Error> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

fn package() -> Vec<u8> {
    let link = linked_wire();
    write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_1"])
            .expect("object storage should write");
        writer
            .create_stream(&["ObjectPool", "_1", "\u{0001}Ole"], &link)
            .expect("link stream should write");
        writer
            .create_stream(&["ObjectPool", "_1", "Payload"], b"opaque")
            .expect("payload should write");
    })
}

#[test]
fn selected_objects_expose_shared_link_metadata() {
    let editor = open_editor(package(), target(), Limits::default()).expect("editor should open");
    let link = editor
        .objects()
        .get("object")
        .expect("object should be discovered")
        .link(LinkLimits::default())
        .expect("link should parse")
        .expect("link stream should be present");
    assert_eq!(link.kind(), Kind::Linked);
    assert_eq!(link.link_update_option(), 7);
    assert_eq!(link.unknown_tail(), &[0xAA, 0xBB, 0xCC]);
    assert_eq!(
        staged(&editor)
            .link("object", LinkLimits::default())
            .unwrap()
            .unwrap(),
        link
    );
}

#[test]
fn link_edits_publish_atomically_and_preserve_opaque_streams() {
    let original = package();
    let mut editor =
        open_editor(original.clone(), target(), Limits::default()).expect("editor should open");
    editor
        .update_link("object", LinkLimits::default(), |link| {
            link.set_cache_hint(false);
            link.set_link_update_option(19);
            Ok(())
        })
        .expect("link edit should commit");
    let object = editor
        .objects()
        .get("object")
        .expect("object should remain");
    let link = object
        .link(LinkLimits::default())
        .expect("edited link should parse")
        .expect("edited link should remain");
    assert!(!link.cache_hint());
    assert_eq!(link.link_update_option(), 19);
    assert_eq!(link.unknown_tail(), &[0xAA, 0xBB, 0xCC]);
    assert_eq!(object.stream(&["Payload"]), Some(&b"opaque"[..]));

    let mut rejected =
        open_editor(original.clone(), target(), Limits::default()).expect("editor should reopen");
    assert!(
        rejected
            .update_link("object", LinkLimits::default(), |_link| {
                Err(litchi_cfb::OleError::InvalidFormat("reject edit".into()).into())
            })
            .is_err()
    );
    assert!(!rejected.clone().commit().unwrap().diagnostics().changed());
    assert_eq!(
        rejected
            .render_bytes()
            .expect("rejected editor should finish"),
        original
    );
}
