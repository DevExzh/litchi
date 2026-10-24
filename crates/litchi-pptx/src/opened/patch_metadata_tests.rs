//! Native durable metadata preservation regressions.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
#[test]
fn durable_forward_and_inverse_preserve_exact_metadata() {
    use super::{Patch, apply};
    use crate::opened::Limits;
    use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
    use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

    let initial = crate::Package::new().unwrap().to_bytes().unwrap();
    let archive =
        ArchiveReader::new(&initial, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    for name in archive.file_names() {
        let mut bytes = archive.read(name).unwrap();
        if name == "[Content_Types].xml" {
            let mut xml = String::from_utf8(bytes).unwrap();
            let end = xml.rfind("</").unwrap();
            xml.insert_str(end, "<!--CT-SENTINEL-->");
            bytes = xml.into_bytes();
        } else if name == "ppt/_rels/presentation.xml.rels" {
            let mut xml = String::from_utf8(bytes).unwrap();
            let end = xml.rfind("</").unwrap();
            xml.insert_str(end, "<!--RELS-SENTINEL-->");
            bytes = xml.into_bytes();
        }
        writer.write_stored(name, &bytes).unwrap();
    }
    let original = writer.finish().unwrap().into_inner();
    let source = OpcPackage::from_bytes(&original, litchi_opc::SourceLimits::default()).unwrap();
    let main = PackURI::new("/ppt/presentation.xml").unwrap();
    let added = PackURI::new("/custom/probe.bin").unwrap();
    let content_types = source
        .source_content_types(litchi_opc::SourceLimits::default())
        .unwrap();
    let relationships = source
        .source_relationships(&main, litchi_opc::SourceLimits::default())
        .unwrap();
    assert!(
        content_types
            .bytes()
            .windows(11)
            .any(|w| w == b"CT-SENTINEL")
    );
    assert!(
        relationships
            .bytes()
            .windows(13)
            .any(|w| w == b"RELS-SENTINEL")
    );
    let next_types = content_types
        .with_part_overrides(&[(&added, "application/octet-stream")], 1 << 20)
        .unwrap();
    let next_rels = relationships
        .with_relationship(
            "urn:litchi:probe",
            "../custom/probe.bin",
            "rIdProbe",
            TargetMode::Internal,
            1 << 20,
        )
        .unwrap();
    let mut direct = source.clone();
    direct
        .try_add_parts_with_source_tokens(
            content_types.bytes(),
            &next_types,
            &relationships,
            &next_rels,
            vec![Box::new(BlobPart::new(
                added,
                "application/octet-stream".into(),
                vec![1, 2, 3],
            ))],
        )
        .unwrap();
    let patch = Patch::capture(&source, &direct, main.clone(), Limits::default()).unwrap();
    let patch = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
    let mut replayed = source.clone();
    apply(&mut replayed, &patch, true).unwrap();
    assert_eq!(
        direct
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        replayed
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
    assert_eq!(
        direct
            .source_relationships(&main, litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        replayed
            .source_relationships(&main, litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
    assert_eq!(
        PackageWriter::to_bytes(&direct).unwrap(),
        PackageWriter::to_bytes(&replayed).unwrap()
    );
    apply(&mut replayed, &patch.inverse(), true).unwrap();
    assert_eq!(
        source
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        replayed
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
    assert_eq!(
        source
            .source_relationships(&main, litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        replayed
            .source_relationships(&main, litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
}

use super::{Patch, apply};
use crate::opened::Limits;
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};
use std::sync::Arc;
fn main() -> PackURI {
    PackURI::new("/ppt/presentation.xml").unwrap()
}
fn source() -> OpcPackage {
    let bytes = crate::Package::new().unwrap().to_bytes().unwrap();
    let mut source = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let token = source
        .source_content_types(litchi_opc::SourceLimits::default())
        .unwrap();
    let bytes = String::from_utf8(token.bytes().to_vec())
        .unwrap()
        .replace("</Types>", "<!--types retained--></Types>");
    source
        .try_replace_content_types_bytes(
            token.bytes(),
            Arc::new(bytes.into_bytes()),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    let token = source
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let bytes = String::from_utf8(token.bytes().to_vec())
        .unwrap()
        .replace("</Relationships>", "<!--rels retained--></Relationships>");
    source
        .try_replace_relationships_bytes(
            &token,
            Arc::new(bytes.into_bytes()),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    source
}
fn add(source: &OpcPackage, id: &str) -> OpcPackage {
    let name = PackURI::new(format!("/custom/{id}.bin")).unwrap();
    let types = source
        .source_content_types(litchi_opc::SourceLimits::default())
        .unwrap();
    let rels = source
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let after_types = types
        .with_part_overrides(&[(&name, "application/octet-stream")], 1 << 20)
        .unwrap();
    let after_rels = rels
        .with_relationship(
            "urn:probe",
            &format!("../custom/{id}.bin"),
            id,
            TargetMode::Internal,
            1 << 20,
        )
        .unwrap();
    let mut target = source.clone();
    target
        .try_add_parts_with_source_tokens(
            types.bytes(),
            &after_types,
            &rels,
            &after_rels,
            vec![Box::new(BlobPart::new(
                name,
                "application/octet-stream".into(),
                vec![1, 2, 3],
            ))],
        )
        .unwrap();
    target
}
fn durable(source: &OpcPackage, target: &OpcPackage) -> Patch {
    let patch = Patch::capture(source, target, main(), Limits::default()).unwrap();
    Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap()
}
#[test]
fn disjoint_metadata_keys_join_deterministically_and_invert_exactly() {
    let source = source();
    let left = durable(&source, &add(&source, "rIdLeft"));
    let right = durable(&source, &add(&source, "rIdRight"));
    assert!(!left.conflicts_with(&right));
    let joined = left.join(&right).unwrap();
    assert_eq!(
        joined.to_bytes().unwrap(),
        right.join(&left).unwrap().to_bytes().unwrap()
    );
    let joined = Patch::from_bytes(&joined.to_bytes().unwrap(), Limits::default()).unwrap();
    let mut result = source.clone();
    apply(&mut result, &joined, true).unwrap();
    assert!(
        result
            .get_part(&main())
            .unwrap()
            .rels()
            .get("rIdLeft")
            .is_some()
    );
    assert!(
        result
            .get_part(&main())
            .unwrap()
            .rels()
            .get("rIdRight")
            .is_some()
    );
    let published = PackageWriter::to_bytes(&result).unwrap();
    let mut result =
        OpcPackage::from_bytes(&published, litchi_opc::SourceLimits::default()).unwrap();
    let inverse =
        Patch::from_bytes(&joined.inverse().to_bytes().unwrap(), Limits::default()).unwrap();
    apply(&mut result, &inverse, true).unwrap();
    assert_eq!(
        result
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        source
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
    assert_eq!(
        result
            .source_relationships(&main(), litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        source
            .source_relationships(&main(), litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes()
    );
}
#[test]
fn replay_and_inverse_preserve_unrelated_destination_metadata_and_edges() {
    let source = source();
    let patch = durable(&source, &add(&source, "rIdPatch"));
    let overlay = add(&source, "rIdOverlay");
    let expected_types = overlay
        .source_content_types(litchi_opc::SourceLimits::default())
        .unwrap();
    let expected_rels = overlay
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let mut result = overlay.clone();
    apply(&mut result, &patch, true).unwrap();
    assert!(
        result
            .get_part(&main())
            .unwrap()
            .rels()
            .get("rIdOverlay")
            .is_some()
    );
    assert!(
        result
            .get_part(&main())
            .unwrap()
            .rels()
            .get("rIdPatch")
            .is_some()
    );
    let bytes = PackageWriter::to_bytes(&result).unwrap();
    let mut reopened = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    apply(&mut reopened, &patch.inverse(), true).unwrap();
    assert_eq!(
        reopened
            .source_content_types(litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        expected_types.bytes()
    );
    assert_eq!(
        reopened
            .source_relationships(&main(), litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        expected_rels.bytes()
    );
}
#[test]
fn metadata_only_spelling_change_is_durable_and_not_a_noop() {
    let source = source();
    let mut target = source.clone();
    let token = target
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let changed = String::from_utf8(token.bytes().to_vec())
        .unwrap()
        .replace("rels retained", "rels changed");
    target
        .try_replace_relationships_bytes(
            &token,
            Arc::new(changed.into_bytes()),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    let patch = durable(&source, &target);
    assert!(!patch.is_noop());
    assert_eq!(patch.resource_count(), 1);
    let mut result = source.clone();
    apply(&mut result, &patch, true).unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&result).unwrap(),
        PackageWriter::to_bytes(&target).unwrap()
    );
    apply(&mut result, &patch.inverse(), true).unwrap();
    assert_eq!(
        result
            .source_relationships(&main(), litchi_opc::SourceLimits::default())
            .unwrap()
            .bytes(),
        token.bytes()
    );
}
#[test]
fn source_context_and_touched_entry_conflicts_leave_destination_unchanged() {
    let source = source();
    let patch = durable(&source, &add(&source, "rIdPatch"));
    let mut stale = add(&source, "rIdPatch");
    let before = PackageWriter::to_bytes(&stale).unwrap();
    assert!(apply(&mut stale, &patch, true).is_err());
    assert_eq!(PackageWriter::to_bytes(&stale).unwrap(), before);
    let mut context = add(&source, "rIdOverlay");
    let token = context
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let bytes = String::from_utf8(token.bytes().to_vec()).unwrap().replace(
        "<Relationships ",
        "<Relationships xmlns:extra='urn:changed' ",
    );
    context
        .try_replace_relationships_bytes(
            &token,
            Arc::new(bytes.into_bytes()),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    let before = PackageWriter::to_bytes(&context).unwrap();
    assert!(apply(&mut context, &patch, true).is_err());
    assert_eq!(PackageWriter::to_bytes(&context).unwrap(), before);
}

#[test]
fn metadata_wire_is_framed_before_allocation_and_old_versions_are_refused() {
    let source = source();
    let patch = durable(&source, &add(&source, "rIdPatch"));
    let bytes = patch.to_bytes().unwrap();
    for end in 0..bytes.len() {
        assert!(
            Patch::from_bytes(&bytes[..end], Limits::default()).is_err(),
            "prefix {end}"
        );
    }
    let mut old = bytes.clone();
    old[..8].copy_from_slice(b"LPTX0002");
    assert!(Patch::from_bytes(&old, Limits::default()).is_err());
    let metadata_start = bytes.len() - patch.metadata.encoded_len().unwrap();
    let mut stripped = bytes[..metadata_start].to_vec();
    stripped.extend_from_slice(&0u32.to_le_bytes());
    assert!(Patch::from_bytes(&stripped, Limits::default()).is_err());
    let mut forged = bytes.clone();
    forged[metadata_start..metadata_start + 4].copy_from_slice(&u32::MAX.to_le_bytes());
    let generous = Limits::new(usize::MAX, bytes.len(), 1024, 1, bytes.len(), bytes.len()).unwrap();
    assert!(Patch::from_bytes(&forged, generous).is_err());
    let mut trailing = bytes.clone();
    trailing.push(0);
    assert!(Patch::from_bytes(&trailing, Limits::default()).is_err());
    let exact = Limits::new(4096, bytes.len(), 1024, 1, bytes.len(), bytes.len()).unwrap();
    assert_eq!(
        Patch::from_bytes(&bytes, exact)
            .unwrap()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

fn root_edit(source: &OpcPackage, id: &str, target: &str) -> OpcPackage {
    let owner = PackURI::new("/").unwrap();
    let token = source
        .source_relationships(&owner, litchi_opc::SourceLimits::default())
        .unwrap();
    let after = token
        .with_relationship("urn:root-probe", target, id, TargetMode::External, 1 << 20)
        .unwrap();
    let mut result = source.clone();
    result
        .try_replace_relationships_bytes(
            &token,
            after.bytes_arc(),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    result
}
fn root_comment(source: &OpcPackage, comment: &str) -> OpcPackage {
    let token = source
        .source_relationships(
            &PackURI::new("/").unwrap(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    let xml = String::from_utf8(token.bytes().to_vec()).unwrap().replace(
        "</Relationships>",
        &format!("<!--{comment}--></Relationships>"),
    );
    let mut result = source.clone();
    result
        .try_replace_relationships_bytes(
            &token,
            Arc::new(xml.into_bytes()),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    result
}
#[test]
fn root_three_way_merges_disjoint_ids_and_inverts_after_durable_decode() {
    let source = root_comment(&source(), "root sentinel");
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let left = durable(
        &source,
        &root_edit(&source, "left", "https://example.com/left"),
    );
    let right = durable(
        &source,
        &root_edit(&source, "right", "https://example.com/right"),
    );
    let plan = Patch::three_way(&base, &left, &right).unwrap();
    assert_eq!(plan.unresolved_count(), 0);
    let merged = plan.finish().unwrap();
    assert_eq!(
        merged.to_bytes().unwrap(),
        Patch::three_way(&base, &right, &left)
            .unwrap()
            .finish()
            .unwrap()
            .to_bytes()
            .unwrap()
    );
    let merged = Patch::from_bytes(&merged.to_bytes().unwrap(), Limits::default()).unwrap();
    let mut target = source.clone();
    apply(&mut target, &merged, true).unwrap();
    assert!(target.rels().get("left").is_some() && target.rels().get("right").is_some());
    let inverse =
        Patch::from_bytes(&merged.inverse().to_bytes().unwrap(), Limits::default()).unwrap();
    apply(&mut target, &inverse, true).unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&source).unwrap(),
        PackageWriter::to_bytes(&target).unwrap()
    );
}
#[test]
fn root_three_way_resolution_keeps_graph_and_exact_metadata_together() {
    use super::Resolution;
    let source = root_comment(&source(), "root sentinel");
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    for (left_target, right_target) in [
        (
            root_edit(&source, "shared", "https://example.com/left"),
            root_edit(&source, "shared", "https://example.com/right"),
        ),
        (
            root_comment(&source, "left spelling"),
            root_comment(&source, "right spelling"),
        ),
        (
            root_comment(&source, "left spelling"),
            root_edit(&source, "shared", "https://example.com/right"),
        ),
    ] {
        let left = durable(&source, &left_target);
        let right = durable(&source, &right_target);
        let plan = Patch::three_way(&base, &left, &right).unwrap();
        assert_eq!(plan.unresolved_count(), 1);
        assert_eq!(plan.conflicts().next().unwrap().resource().as_str(), "/");
        assert!(plan.finish().is_err());
        assert!(plan.resolve(1, Resolution::Left).is_err());
        for (resolution, expected) in [
            (Resolution::Base, &source),
            (Resolution::Left, &left_target),
            (Resolution::Right, &right_target),
        ] {
            let patch = plan.resolve(0, resolution).unwrap().finish().unwrap();
            let patch = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
            let mut target = source.clone();
            apply(&mut target, &patch, true).unwrap();
            assert_eq!(
                PackageWriter::to_bytes(expected).unwrap(),
                PackageWriter::to_bytes(&target).unwrap()
            );
            apply(
                &mut target,
                &Patch::from_bytes(&patch.inverse().to_bytes().unwrap(), Limits::default())
                    .unwrap(),
                true,
            )
            .unwrap();
            assert_eq!(
                PackageWriter::to_bytes(&source).unwrap(),
                PackageWriter::to_bytes(&target).unwrap()
            );
        }
        assert_eq!(plan.unresolved_count(), 1);
    }
}
#[test]
fn root_resolution_retains_independent_part_metadata() {
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let left_target = root_edit(&add(&source, "part"), "shared", "https://example.com/left");
    let right_target = root_edit(&source, "shared", "https://example.com/right");
    let plan = Patch::three_way(
        &base,
        &durable(&source, &left_target),
        &durable(&source, &right_target),
    )
    .unwrap();
    let selected = plan
        .resolve(0, super::Resolution::Right)
        .unwrap()
        .finish()
        .unwrap();
    let selected = Patch::from_bytes(&selected.to_bytes().unwrap(), Limits::default()).unwrap();
    let mut target = source.clone();
    apply(&mut target, &selected, true).unwrap();
    assert_eq!(
        target.rels().get("shared").unwrap().target_ref(),
        "https://example.com/right"
    );
    assert!(
        target
            .get_part(&PackURI::new("/custom/part.bin").unwrap())
            .is_ok()
    );
    apply(
        &mut target,
        &Patch::from_bytes(&selected.inverse().to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&source).unwrap(),
        PackageWriter::to_bytes(&target).unwrap()
    );
}

#[test]
fn root_and_payload_conflicts_resolve_independently_in_stable_order() {
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let mut left = root_edit(&source, "shared", "https://example.com/left");
    let mut right = root_edit(&source, "shared", "https://example.com/right");
    for (target, comment) in [
        (&mut left, "<!--left payload-->"),
        (&mut right, "<!--right payload-->"),
    ] {
        let part = target.get_part_mut(&main()).unwrap();
        let mut bytes = part.blob().to_vec();
        bytes.extend_from_slice(comment.as_bytes());
        part.set_blob(bytes);
    }
    let plan =
        Patch::three_way(&base, &durable(&source, &left), &durable(&source, &right)).unwrap();
    let resources: Vec<_> = plan.conflicts().map(|c| c.resource().as_str()).collect();
    assert_eq!(resources, ["/", "/ppt/presentation.xml"]);
    let selected = plan
        .resolve(0, super::Resolution::Left)
        .unwrap()
        .resolve(1, super::Resolution::Right)
        .unwrap()
        .finish()
        .unwrap();
    let mut target = source.clone();
    apply(
        &mut target,
        &Patch::from_bytes(&selected.to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    assert_eq!(
        target.rels().get("shared").unwrap().target_ref(),
        "https://example.com/left"
    );
    assert_eq!(
        target.get_part(&main()).unwrap().blob(),
        right.get_part(&main()).unwrap().blob()
    );
    apply(&mut target, &selected.inverse(), true).unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&source).unwrap(),
        PackageWriter::to_bytes(&target).unwrap()
    );
}

#[test]
fn root_three_way_admits_combined_counts_and_wire_bytes() {
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let mut left = durable(
        &source,
        &root_edit(&source, "left", "https://example.com/left"),
    );
    let mut right = durable(
        &source,
        &root_edit(&source, "right", "https://example.com/right"),
    );
    let bytes = Patch::three_way(&base, &left, &right)
        .unwrap()
        .finish()
        .unwrap()
        .to_bytes()
        .unwrap()
        .len();
    let policy = |parts, bytes| Limits::new(parts, bytes, 1024, 10, 1 << 20, 1 << 20).unwrap();
    left.limits = policy(source.rels().len() + 1, 1 << 20);
    right.limits = left.limits;
    assert!(
        Patch::three_way(&base, &left, &right).is_err(),
        "merged graph must be admitted before cloning"
    );
    left.limits = policy(1024, bytes);
    right.limits = left.limits;
    assert_eq!(
        Patch::three_way(&base, &left, &right)
            .unwrap()
            .finish()
            .unwrap()
            .to_bytes()
            .unwrap()
            .len(),
        bytes
    );
    left.limits = policy(1024, bytes - 1);
    right.limits = left.limits;
    let plan = Patch::three_way(&base, &left, &right).unwrap();
    assert!(plan.finish().is_err());
    assert_eq!(plan.unresolved_count(), 0);
}

fn part_relationship(source: &OpcPackage, target: &str) -> OpcPackage {
    let token = source
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let after = token
        .with_relationship(
            "urn:part-probe",
            target,
            "shared",
            TargetMode::External,
            1 << 20,
        )
        .unwrap();
    let mut result = source.clone();
    result
        .try_replace_relationships_bytes(
            &token,
            after.bytes_arc(),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    result
}
fn part_comment(source: &OpcPackage, comment: &str) -> OpcPackage {
    let token = source
        .source_relationships(&main(), litchi_opc::SourceLimits::default())
        .unwrap();
    let bytes = String::from_utf8(token.bytes().to_vec()).unwrap().replace(
        "</Relationships>",
        &format!("<!--{comment}--></Relationships>"),
    );
    let mut result = source.clone();
    result
        .try_replace_relationships_bytes(
            &token,
            Arc::new(bytes.into_bytes()),
            true,
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
    result
}
fn payload_comment(source: &OpcPackage, comment: &str) -> OpcPackage {
    let mut result = source.clone();
    let part = result.get_part_mut(&main()).unwrap();
    let mut bytes = part.blob().to_vec();
    bytes.extend_from_slice(format!("<!--{comment}-->").as_bytes());
    part.set_blob(bytes);
    result
}
#[test]
fn part_three_way_selects_payload_graph_and_metadata_as_one_candidate() {
    use super::Resolution;
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    for (left, right) in [
        (
            part_relationship(&source, "https://example.com/left"),
            part_relationship(&source, "https://example.com/right"),
        ),
        (
            part_comment(&source, "left spelling"),
            part_comment(&source, "right spelling"),
        ),
        (
            part_comment(&source, "left spelling"),
            part_relationship(&source, "https://example.com/right"),
        ),
        (
            payload_comment(
                &part_relationship(&source, "https://example.com/left"),
                "left payload",
            ),
            payload_comment(&source, "right payload"),
        ),
    ] {
        let plan =
            Patch::three_way(&base, &durable(&source, &left), &durable(&source, &right)).unwrap();
        assert_eq!(plan.unresolved_count(), 1);
        assert_eq!(plan.conflicts().next().unwrap().resource(), &main());
        assert!(plan.finish().is_err());
        for (resolution, expected) in [
            (Resolution::Base, &source),
            (Resolution::Left, &left),
            (Resolution::Right, &right),
        ] {
            let patch = plan.resolve(0, resolution).unwrap().finish().unwrap();
            let patch = Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap();
            let mut result = source.clone();
            apply(&mut result, &patch, true).unwrap();
            assert_eq!(
                PackageWriter::to_bytes(expected).unwrap(),
                PackageWriter::to_bytes(&result).unwrap()
            );
            let inverse =
                Patch::from_bytes(&patch.inverse().to_bytes().unwrap(), Limits::default()).unwrap();
            apply(&mut result, &inverse, true).unwrap();
            assert_eq!(
                PackageWriter::to_bytes(&source).unwrap(),
                PackageWriter::to_bytes(&result).unwrap()
            );
        }
        assert_eq!(plan.unresolved_count(), 1);
    }
}
#[test]
fn resolving_a_part_preserves_other_owners_and_content_type_edits() {
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let left = payload_comment(&add(&source, "independent"), "left");
    let right = payload_comment(
        &root_edit(&source, "root", "https://example.com/root"),
        "right",
    );
    let plan =
        Patch::three_way(&base, &durable(&source, &left), &durable(&source, &right)).unwrap();
    let selected = plan
        .resolve(0, super::Resolution::Right)
        .unwrap()
        .finish()
        .unwrap();
    let mut result = source.clone();
    apply(
        &mut result,
        &Patch::from_bytes(&selected.to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    assert!(
        result
            .get_part(&PackURI::new("/custom/independent.bin").unwrap())
            .is_ok()
    );
    assert!(result.rels().get("root").is_some());
    assert_eq!(
        result.get_part(&main()).unwrap().blob(),
        right.get_part(&main()).unwrap().blob()
    );
    // The main owner's relationship edit belonged to the discarded candidate.
    assert!(
        result
            .get_part(&main())
            .unwrap()
            .rels()
            .get("independent")
            .is_none()
    );
    apply(
        &mut result,
        &Patch::from_bytes(&selected.inverse().to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&source).unwrap(),
        PackageWriter::to_bytes(&result).unwrap()
    );
}

#[test]
fn root_and_part_metadata_choices_share_exact_combined_admission() {
    let source = source();
    let base = crate::opened::model::capture(&source, Limits::default(), true).unwrap();
    let left = part_comment(&root_comment(&source, "root left"), "part left");
    let right = part_comment(&root_comment(&source, "root right"), "part right");
    let plan =
        Patch::three_way(&base, &durable(&source, &left), &durable(&source, &right)).unwrap();
    assert_eq!(
        plan.conflicts()
            .map(|c| c.resource().as_str())
            .collect::<Vec<_>>(),
        ["/", "/ppt/presentation.xml"]
    );
    let mut selected = plan
        .resolve(0, super::Resolution::Left)
        .unwrap()
        .resolve(1, super::Resolution::Right)
        .unwrap();
    let size = selected.finish().unwrap().to_bytes().unwrap().len();
    selected.limits = Limits::new(1024, size, 1024, 10, 1 << 20, 1 << 20).unwrap();
    let patch = selected.finish().unwrap();
    assert_eq!(patch.to_bytes().unwrap().len(), size);
    let mut result = source.clone();
    apply(
        &mut result,
        &Patch::from_bytes(&patch.to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    let expected = part_comment(&root_comment(&source, "root left"), "part right");
    assert_eq!(
        PackageWriter::to_bytes(&expected).unwrap(),
        PackageWriter::to_bytes(&result).unwrap()
    );
    apply(
        &mut result,
        &Patch::from_bytes(&patch.inverse().to_bytes().unwrap(), Limits::default()).unwrap(),
        true,
    )
    .unwrap();
    assert_eq!(
        PackageWriter::to_bytes(&source).unwrap(),
        PackageWriter::to_bytes(&result).unwrap()
    );
    selected.limits = Limits::new(1024, size - 1, 1024, 10, 1 << 20, 1 << 20).unwrap();
    assert!(selected.finish().is_err());
}
