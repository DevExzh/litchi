//! Borrowed patch admission and foreign payload-owner regressions.

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use super::{Patch, relationships_match, state_matches, validate_before};
use crate::opened::Limits;
use crate::{Error, Result};

#[derive(Clone, Debug)]
struct ForeignPart {
    inner: BlobPart,
    shared: Arc<Vec<u8>>,
    calls: Arc<AtomicUsize>,
    prefix: Option<usize>,
    switch_after_shared: bool,
}

impl Part for ForeignPart {
    fn blob(&self) -> &[u8] {
        if self.switch_after_shared && self.calls.load(Ordering::Relaxed) != 0 {
            return &self.shared;
        }
        let visible = self.inner.blob();
        &visible[..self.prefix.unwrap_or(visible.len())]
    }
    fn blob_arc(&self) -> Arc<Vec<u8>> {
        self.calls.fetch_add(1, Ordering::Relaxed);
        Arc::clone(&self.shared)
    }
    fn content_type(&self) -> &str {
        self.inner.content_type()
    }
    fn partname(&self) -> &PackURI {
        self.inner.partname()
    }
    fn rels(&self) -> &Relationships {
        self.inner.rels()
    }
    fn rels_mut(&mut self) -> &mut Relationships {
        self.inner.rels_mut()
    }
    fn set_blob(&mut self, blob: Vec<u8>) {
        self.inner.set_blob(blob);
    }
    fn set_content_type(&mut self, value: String) -> litchi_opc::Result<()> {
        self.inner.set_content_type(value)
    }
}

fn package(
    payloads: &[&[u8]],
    shared: &Arc<Vec<u8>>,
    calls: &Arc<AtomicUsize>,
) -> Result<OpcPackage> {
    let mut package = OpcPackage::new();
    for (index, payload) in payloads.iter().enumerate() {
        package.try_add_part(Box::new(ForeignPart {
            inner: BlobPart::new(
                name(index)?,
                "application/octet-stream".into(),
                payload.to_vec(),
            ),
            shared: Arc::clone(shared),
            calls: Arc::clone(calls),
            prefix: None,
            switch_after_shared: false,
        }))?;
    }
    Ok(package)
}

fn name(index: usize) -> Result<PackURI> {
    PackURI::new(format!("/part{index}.bin")).map_err(Error::Invalid)
}

fn limits(parts: usize, bytes: usize) -> Result<Limits> {
    Limits::new(parts, bytes, 1024, 10, 1024 * 1024, 1024 * 1024)
        .ok_or_else(|| Error::Invalid("test limits must be nonzero".into()))
}

#[test]
fn identical_foreign_owners_cannot_hide_visible_changes() -> Result<()> {
    let shared = Arc::new(b"forged".to_vec());
    let calls = Arc::new(AtomicUsize::new(0));
    let source = package(&[b"before"], &shared, &calls)?;
    let target = package(&[b"after!"], &shared, &calls)?;
    assert!(!crate::opened::model::packages_equal(&source, &target));
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    let patch = Patch::capture(&source, &target, name(9)?, Limits::default())?;
    assert_eq!(patch.resource_count(), 1);
    let before = patch.deltas[0].before.as_ref().unwrap();
    let after = patch.deltas[0].after.as_ref().unwrap();
    assert_eq!(before.blob.as_slice(), b"before");
    assert_eq!(after.blob.as_slice(), b"after!");
    assert!(!Arc::ptr_eq(&before.blob, &shared));
    assert!(!Arc::ptr_eq(&after.blob, &shared));
    let encoded = patch.to_bytes()?;
    assert_eq!(Patch::from_bytes(&encoded, Limits::default())?, patch);
    validate_before(&source, &patch)?;
    validate_before(&target, &patch.inverse())?;
    assert!(validate_before(&target, &patch).is_err());
    Ok(())
}

#[test]
fn noops_and_validation_do_not_request_shared_payloads() -> Result<()> {
    let calls = Arc::new(AtomicUsize::new(0));
    let source = package(&[b"visible"], &Arc::new(b"fake one".to_vec()), &calls)?;
    let target = package(&[b"visible"], &Arc::new(b"fake two".to_vec()), &calls)?;
    let patch = Patch::capture(&source, &target, name(9)?, Limits::default())?;
    assert!(patch.is_noop());
    assert!(crate::opened::model::packages_equal(&source, &target));
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    let changed = package(&[b"changed"], &Arc::new(Vec::new()), &calls)?;
    let patch = Patch::capture(&source, &changed, name(9)?, Limits::default())?;
    assert!(patch.metadata.is_empty());
    calls.store(0, Ordering::Relaxed);
    validate_before(&source, &patch)?;
    assert!(validate_before(&changed, &patch).is_err());
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    Ok(())
}

#[test]
fn complete_write_set_is_admitted_before_requesting_any_payload_owner() -> Result<()> {
    let shared = Arc::new(Vec::new());
    let calls = Arc::new(AtomicUsize::new(0));
    let source = package(&[b"old one", b"old two"], &shared, &calls)?;
    let target = package(&[b"new one", b"new two"], &shared, &calls)?;
    let full = Patch::capture(&source, &target, name(9)?, Limits::default())?;
    let exact = full.to_bytes()?.len();
    calls.store(0, Ordering::Relaxed);
    for bounded in [limits(1, exact)?, limits(2, exact - 1)?] {
        assert!(matches!(
            Patch::capture(&source, &target, name(9)?, bounded),
            Err(Error::Limit { .. })
        ));
        assert_eq!(calls.load(Ordering::Relaxed), 0);
    }
    let admitted = Patch::capture(&source, &target, name(9)?, limits(2, exact)?)?;
    assert_eq!(admitted.to_bytes()?, full.to_bytes()?);
    assert_eq!(calls.load(Ordering::Relaxed), 4);
    Ok(())
}

#[test]
fn unchanged_large_parts_are_outside_the_patch_budget() -> Result<()> {
    let shared = Arc::new(Vec::new());
    let calls = Arc::new(AtomicUsize::new(0));
    let large = vec![42; 4096];
    let source = package(&[&large, b"old"], &shared, &calls)?;
    let target = package(&[&large, b"new"], &shared, &calls)?;
    let patch = Patch::capture(&source, &target, name(9)?, limits(1, 256)?)?;
    assert_eq!(patch.resource_count(), 1);
    assert_eq!(calls.load(Ordering::Relaxed), 2);
    Ok(())
}

#[test]
fn root_relationships_and_add_remove_states_share_exact_admission() -> Result<()> {
    let calls = Arc::new(AtomicUsize::new(0));
    let source = package(&[], &Arc::new(Vec::new()), &calls)?;
    let mut target = package(&[b"added"], &Arc::new(Vec::new()), &calls)?;
    target.rels_mut().try_add_relationship(
        "urn:kind".into(),
        "part0.bin".into(),
        "rId1".into(),
        TargetMode::Internal,
    )?;
    let full = Patch::capture(&source, &target, name(9)?, Limits::default())?;
    let length = full.to_bytes()?.len();
    calls.store(0, Ordering::Relaxed);
    assert!(matches!(
        Patch::capture(&source, &target, name(9)?, limits(1, length - 1)?),
        Err(Error::Limit { .. })
    ));
    assert_eq!(calls.load(Ordering::Relaxed), 0);
    let bounded = Patch::capture(&source, &target, name(9)?, limits(1, length)?)?;
    assert_eq!(bounded.to_bytes()?, full.to_bytes()?);
    let reverse = Patch::capture(&target, &source, name(9)?, limits(1, length)?)?;
    assert_eq!(reverse, bounded.inverse());
    assert!(relationships_match(
        target.rels(),
        &bounded.root.as_ref().unwrap().1
    ));
    assert!(state_matches(
        target.get_part(&name(0)?).ok(),
        bounded.deltas[0].after.as_ref()
    ));
    Ok(())
}

#[test]
fn aliasing_payloads_keep_their_original_owner() -> Result<()> {
    let source = OpcPackage::new();
    let mut target = OpcPackage::new();
    let bytes = Arc::new(b"shared bytes".to_vec());
    target.try_add_part(Box::new(BlobPart::new_shared(
        name(0)?,
        "application/octet-stream".into(),
        Arc::clone(&bytes),
    )))?;
    let patch = Patch::capture(&source, &target, name(9)?, Limits::default())?;
    assert!(Arc::ptr_eq(
        &patch.deltas[0].after.as_ref().unwrap().blob,
        &bytes
    ));
    Ok(())
}

#[test]
fn relationship_count_admission_matches_the_durable_decoder() -> Result<()> {
    let calls = Arc::new(AtomicUsize::new(0));
    let source = package(&[], &Arc::new(Vec::new()), &calls)?;
    for root in [false, true] {
        let mut target = package(&[b"added"], &Arc::new(Vec::new()), &calls)?;
        let relationships = if root {
            target.rels_mut()
        } else {
            target.get_part_mut(&name(0)?)?.rels_mut()
        };
        for index in 0..2 {
            relationships.try_add_relationship(
                "urn:external".into(),
                "https://example.test".into(),
                format!("rId{index}"),
                TargetMode::External,
            )?;
        }
        assert!(matches!(
            Patch::capture(&source, &target, name(9)?, limits(1, 4096)?),
            Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                ..
            })
        ));
        assert_eq!(calls.load(Ordering::Relaxed), 0);
        let admitted = Patch::capture(&source, &target, name(9)?, limits(2, 4096)?)?;
        assert_eq!(
            Patch::from_bytes(&admitted.to_bytes()?, limits(2, 4096)?)?,
            admitted
        );
        let narrow_noop = Patch::capture(&source, &source, name(9)?, limits(1, 4096)?)?;
        for joined in [admitted.join(&narrow_noop), narrow_noop.join(&admitted)] {
            assert!(matches!(
                joined,
                Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    ..
                })
            ));
        }
        calls.store(0, Ordering::Relaxed);
    }
    Ok(())
}

#[test]
fn shared_owner_prefix_and_later_accessor_switch_preserve_admitted_bytes() -> Result<()> {
    for switch_after_shared in [false, true] {
        let source = OpcPackage::new();
        let mut target = OpcPackage::new();
        let shared = Arc::new(b"visible plus hidden suffix".to_vec());
        target.try_add_part(Box::new(ForeignPart {
            inner: BlobPart::new_shared(
                name(0)?,
                "application/octet-stream".into(),
                Arc::clone(&shared),
            ),
            shared,
            calls: Arc::new(AtomicUsize::new(0)),
            prefix: Some(7),
            switch_after_shared,
        }))?;
        let patch = Patch::capture(&source, &target, name(9)?, Limits::default())?;
        assert_eq!(
            patch.deltas[0].after.as_ref().unwrap().blob.as_slice(),
            b"visible"
        );
        assert_eq!(
            Patch::from_bytes(&patch.to_bytes()?, Limits::default())?,
            patch
        );
    }
    Ok(())
}

#[test]
fn captured_foreign_payload_publishes_and_inverts_on_a_real_presentation() -> Result<()> {
    let mut builder = crate::builder::Builder::new()?;
    builder
        .presentation()?
        .add_slide()?
        .add_text_box("original", 0, 0, 100, 100);
    let authored = builder.build()?;
    let opened = crate::Package::from_reader(std::io::Cursor::new(authored.to_bytes()?))?;
    let source = opened.opc()?.clone();
    let mut target = source.clone();
    let calls = Arc::new(AtomicUsize::new(0));
    let donor = package(
        &[b"visible added bytes"],
        &Arc::new(b"forged".to_vec()),
        &calls,
    )?;
    target.try_add_part(donor.get_part(&name(0)?)?.clone_part())?;
    let main = opened.opened_presentation()?.presentation_name;
    let patch = Patch::capture(&source, &target, main, Limits::default())?;
    let durable = Patch::from_bytes(&patch.to_bytes()?, Limits::default())?;
    let mut published = source.clone();
    super::apply(&mut published, &durable, true)?;
    assert_eq!(
        published.get_part(&name(0)?)?.blob(),
        b"visible added bytes"
    );
    assert!(crate::opened::model::packages_equal(&published, &target));
    super::apply(&mut published, &durable.inverse(), true)?;
    assert!(crate::opened::model::packages_equal(&published, &source));
    Ok(())
}

#[test]
fn wire_width_admission_rejects_unencodable_lengths() {
    assert!(super::require_wire_u32(u32::MAX as usize).is_ok());
    if let Some(too_large) = (u32::MAX as usize).checked_add(1) {
        assert!(super::require_wire_u32(too_large).is_err());
    }
}

#[test]
fn invalid_durable_part_name_error_omits_authored_value() {
    let error = super::parse_part_name(b"/private-client-sentinel/../slide.xml")
        .expect_err("parent segments are not valid part names");
    assert!(matches!(error, Error::Invalid(_)));
    assert!(!format!("{error} {error:?}").contains("private-client-sentinel"));
}
