use super::*;
use crate::source::Patch;

fn wire(source: &Package, text: &str) -> Vec<u8> {
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(0, 0_usize, text).unwrap();
    edit.commit().unwrap().patch().to_bytes().unwrap()
}
fn fresh(bytes: Vec<u8>) -> Package {
    Package::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes))).unwrap()
}
fn read_u32(wire: &[u8], offset: usize) -> usize {
    u32::from_le_bytes(wire[offset..offset + 4].try_into().unwrap()) as usize
}
fn read_u64(wire: &[u8], offset: usize) -> usize {
    u64::from_le_bytes(wire[offset..offset + 8].try_into().unwrap()) as usize
}
fn skip_key(wire: &[u8], offset: usize) -> usize {
    let relationship = offset + 4 + read_u32(wire, offset) + 4;
    relationship + 4 + read_u32(wire, relationship)
}
fn entry_start(wire: &[u8]) -> usize {
    let mut offset = 68;
    for _ in 0..read_u32(wire, 40) + read_u32(wire, 44) {
        offset = skip_key(wire, offset);
    }
    offset + read_u64(wire, 52) + read_u64(wire, 60)
}
fn xml_start(wire: &[u8]) -> usize {
    skip_key(wire, entry_start(wire))
}
fn xmls(wire: &[u8]) -> (&[u8], &[u8]) {
    let start = xml_start(wire);
    let after = start + 8 + read_u64(wire, start);
    (
        &wire[start + 8..after],
        &wire[after + 8..after + 8 + read_u64(wire, after)],
    )
}
fn with_xmls(wire: &[u8], before: &[u8], after: &[u8]) -> Vec<u8> {
    let mut result = wire[..xml_start(wire)].to_vec();
    for xml in [before, after] {
        result.extend_from_slice(&(xml.len() as u64).to_le_bytes());
        result.extend_from_slice(xml);
    }
    result
}

#[test]
fn durable_round_trip_reopens_source_and_replays_inverse_without_reading_unselected_slide() {
    let bytes = fixture();
    let counted = Arc::new(CountingSource::new(bytes.clone()));
    let source = Package::from_read_at(counted.clone()).unwrap();
    let encoded = wire(&source, "portable change");
    assert_eq!(counted.second_payload_reads.load(Ordering::SeqCst), 0);
    let patch = Patch::from_bytes(&encoded).unwrap();
    assert_eq!(patch.to_bytes().unwrap(), encoded);
    drop(source);
    let reopened = fresh(bytes.clone());
    let candidate = reopened.apply(&patch).unwrap().into_package();
    assert_eq!(text(&candidate, 0), "portable change");
    assert_eq!(reopened.to_bytes().unwrap(), bytes);
    assert_eq!(
        candidate
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    let inverse = Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let reopened_changed = fresh(candidate.to_bytes().unwrap());
    assert_eq!(
        text(&reopened_changed.apply(&inverse).unwrap().into_package(), 0),
        "First slide"
    );
    assert!(reopened_changed.apply(&patch).is_err());
    assert!(!format!("{patch:?}").contains("portable change"));
}

#[test]
fn durable_replay_preserves_unrelated_destination_overlay() {
    let bytes = fixture();
    let patch = Patch::from_bytes(&wire(&fresh(bytes.clone()), "left")).unwrap();
    let destination = fresh(bytes);
    let mut edit = destination.edit().unwrap();
    edit.set_shape_text(1, 0_usize, "right").unwrap();
    let candidate = edit
        .commit()
        .unwrap()
        .into_package()
        .apply(&patch)
        .unwrap()
        .into_package();
    assert_eq!(text(&candidate, 0), "left");
    assert_eq!(text(&candidate, 1), "right");
}

#[test]
fn durable_binding_and_before_state_drift_refuse_without_changing_destination() {
    let bytes = fixture();
    let encoded = wire(&fresh(bytes.clone()), "changed");
    let patch = Patch::from_bytes(&encoded).unwrap();
    for root_drift in [false, true] {
        let mut raw = litchi_opc::OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
        if root_drift {
            let root = raw
                .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
                .unwrap();
            let xml = String::from_utf8(root.blob().to_vec())
                .unwrap()
                .replace("</p:presentation>", "<!-- drift --></p:presentation>");
            root.set_blob(xml.into_bytes());
        } else {
            raw.get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:inert".into(),
                    "https://example.invalid".into(),
                    "rIdDrift".into(),
                    true,
                );
        }
        let bytes = litchi_opc::PackageWriter::to_bytes(&raw).unwrap();
        let destination = fresh(bytes.clone());
        assert!(matches!(
            destination.apply(&patch),
            Err(Error::PatchConflict)
        ));
        assert_eq!(destination.to_bytes().unwrap(), bytes);
    }
    let (before, after) = xmls(&encoded);
    let before = String::from_utf8(before.to_vec())
        .unwrap()
        .replace("First slide", "not the source");
    let forged = Patch::from_bytes(&with_xmls(&encoded, before.as_bytes(), after)).unwrap();
    assert!(matches!(
        fresh(bytes).apply(&forged),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn forged_nontext_and_opaque_changes_are_rejected_by_native_owner_proof() {
    let source = fresh(fixture());
    let encoded = wire(&source, "changed");
    let (before, after) = xmls(&encoded);
    for replacement in [
        String::from_utf8(after.to_vec())
            .unwrap()
            .replace("private shape", "renamed shape"),
        String::from_utf8(after.to_vec())
            .unwrap()
            .replace("<p:spPr/>", "<p:spPr><a:xfrm/></p:spPr>"),
        String::from_utf8(after.to_vec())
            .unwrap()
            .replace("</p:sld>", "<!-- injected --></p:sld>"),
    ] {
        assert!(Patch::from_bytes(&with_xmls(&encoded, before, replacement.as_bytes())).is_err());
    }
    let before = String::from_utf8(before.to_vec()).unwrap().replace(
        "</p:cSld>",
        "</p:cSld><p:extLst><p:ext uri=\"urn:opaque\"><a:t>opaque secret</a:t></p:ext></p:extLst>",
    );
    let source = fresh(with_first_slide(before.as_bytes()));
    let encoded = wire(&source, "changed");
    let (before, after) = xmls(&encoded);
    let forged = String::from_utf8(after.to_vec())
        .unwrap()
        .replace("opaque secret", "opaque changed");
    assert!(Patch::from_bytes(&with_xmls(&encoded, before, forged.as_bytes())).is_err());
}

#[test]
fn native_empty_text_and_cdata_changes_transport_and_inverse_exactly() {
    let original = wire(&fresh(fixture()), "changed");
    let (before, _) = xmls(&original);
    for replacement in [
        "<a:t/>",
        "<a:t xml:space=\"preserve\" />",
        "<a:t><![CDATA[First]]> slide</a:t>",
    ] {
        let xml = String::from_utf8(before.to_vec())
            .unwrap()
            .replace("<a:t>First slide</a:t>", replacement);
        let bytes = with_first_slide(xml.as_bytes());
        let patch = Patch::from_bytes(&wire(&fresh(bytes.clone()), "transported")).unwrap();
        let source = fresh(bytes.clone());
        let candidate = source.apply(&patch).unwrap().into_package();
        assert_eq!(text(&candidate, 0), "transported");
        assert_eq!(
            candidate
                .apply(&patch.inverse())
                .unwrap()
                .into_package()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn bounded_wire_framing_refuses_truncation_unknown_versions_duplicates_and_trailing_bytes() {
    let encoded = wire(&fresh(fixture()), "changed");
    for end in 0..encoded.len() {
        assert!(Patch::from_bytes(&encoded[..end]).is_err());
    }
    let mut unknown = encoded.clone();
    unknown[7] = b'9';
    assert!(Patch::from_bytes(&unknown).is_err());
    let mut trailing = encoded.clone();
    trailing.push(0);
    assert!(Patch::from_bytes(&trailing).is_err());
    let mut too_many = encoded.clone();
    too_many[48..52].copy_from_slice(&33_u32.to_le_bytes());
    assert!(matches!(
        Patch::from_bytes(&too_many),
        Err(Error::Limit { .. })
    ));
    let mut length = encoded.clone();
    let start = xml_start(&encoded);
    length[start..start + 8].copy_from_slice(&u64::MAX.to_le_bytes());
    assert!(Patch::from_bytes(&length).is_err());
    let mut duplicates = encoded.clone();
    duplicates[48..52].copy_from_slice(&2_u32.to_le_bytes());
    duplicates.extend_from_slice(&encoded[entry_start(&encoded)..]);
    assert!(Patch::from_bytes(&duplicates).is_err());
}

#[test]
fn durable_noop_preserves_signed_source_and_changed_signed_replay_refuses() {
    let signed =
        fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx").unwrap();
    let source = fresh(signed.clone());
    let noop = source
        .edit()
        .unwrap()
        .commit()
        .unwrap()
        .patch()
        .to_bytes()
        .unwrap();
    let noop = Patch::from_bytes(&noop).unwrap();
    assert!(noop.is_noop());
    assert_eq!(
        fresh(signed.clone())
            .apply(&noop)
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        signed
    );
    // Rebind an unsigned fixture's real delta to the signed root. Its selected
    // slide closure still matches, so the refusal must be signature policy.
    let mut unsigned = crate::Package::from_bytes(&signed).unwrap();
    unsigned = unsigned.unsign().unwrap();
    let mut encoded = wire(&fresh(unsigned.to_bytes().unwrap()), "changed");
    encoded[8..40].copy_from_slice(&noop.to_bytes().unwrap()[8..40]);
    let decoded = Patch::from_bytes(&encoded).unwrap();
    assert!(matches!(
        source.apply(&decoded),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert_eq!(source.to_bytes().unwrap(), signed);
}

#[test]
fn managed_destination_policy_and_retained_payloads_are_independent_of_encoding_source() {
    let bytes = fixture();
    let patch = Patch::from_bytes(&wire(&fresh(bytes.clone()), "portable payload")).unwrap();
    let (budget, cancellation, context) = source_backed_context(16 * 1024 * 1024);
    let destination = Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(bytes)),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let candidate = destination.apply(&patch).unwrap().into_package();
    let scene = candidate
        .presentation()
        .unwrap()
        .slide(0)
        .unwrap()
        .scene()
        .unwrap();
    let handle = scene.shape(0_usize).unwrap().unwrap();
    cancellation.cancel();
    assert!(destination.apply(&patch).is_err());
    assert!(candidate.to_bytes().is_err());
    drop((candidate, destination, patch, scene));
    assert!(budget.used(Resource::Memory) > 0);
    assert_eq!(handle.view().name(), Some("private shape"));
    drop(handle);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn destination_xml_event_limit_refuses_expansion_before_publication() {
    let encoded = wire(&fresh(fixture()), "changed");
    let (before, _) = xmls(&encoded);
    let empty = String::from_utf8(before.to_vec())
        .unwrap()
        .replace("<a:t>First slide</a:t>", "<a:t/>");
    let bytes = with_first_slide(empty.as_bytes());
    let patch = Patch::from_bytes(&wire(&fresh(bytes.clone()), "expanded")).unwrap();
    let limits = SourceLimits::builder()
        .max_events(24)
        .unwrap()
        .build()
        .unwrap();
    let projection = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(bytes.clone())),
        crate::OpenOptions::default().source_limits(limits),
    )
    .unwrap();
    let destination = Package::from_presentation(&projection).unwrap();
    assert!(matches!(
        destination.apply(&patch),
        Err(Error::Opc(litchi_opc::OpcError::Limit {
            kind: SourceLimitKind::Events,
            ..
        }))
    ));
    assert_eq!(destination.to_bytes().unwrap(), bytes);
}

#[test]
fn managed_replay_budget_failure_is_atomic_and_recovers() {
    const MEMORY: u64 = 16 * 1024 * 1024;
    let bytes = fixture();
    let patch = Patch::from_bytes(&wire(&fresh(bytes.clone()), "portable")).unwrap();
    let (budget, _cancel, context) = source_backed_context(MEMORY);
    let destination = Package::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(bytes.clone())),
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    drop(destination.presentation().unwrap().slide(0).unwrap());
    let used = budget.used(Resource::Memory);
    let pressure = budget.reserve(Resource::Memory, MEMORY - used).unwrap();
    assert!(matches!(
        destination.apply(&patch),
        Err(Error::Opc(litchi_opc::OpcError::Execution(
            ExecutionError::ResourceLimit(_)
        )))
    ));
    drop(pressure);
    assert_eq!(budget.used(Resource::Memory), used);
    assert_eq!(destination.to_bytes().unwrap(), bytes);
    let candidate = destination.apply(&patch).unwrap().into_package();
    assert_eq!(text(&candidate, 0), "portable");
    drop((candidate, destination));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn second_slide_conflict_rejects_complete_durable_batch() {
    let bytes = fixture();
    let source = fresh(bytes.clone());
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "left").unwrap();
    edit.set_shape_text(1, 0_usize, "right").unwrap();
    let mut encoded = edit.commit().unwrap().patch().to_bytes().unwrap();
    let offset = encoded
        .windows(b"Second slide".len())
        .position(|value| value == b"Second slide")
        .unwrap();
    encoded[offset..offset + 12].copy_from_slice(b"Oldbad slide");
    let patch = Patch::from_bytes(&encoded).unwrap();
    let destination = fresh(bytes.clone());
    assert!(matches!(
        destination.apply(&patch),
        Err(Error::PatchConflict)
    ));
    assert_eq!(destination.to_bytes().unwrap(), bytes);
}

#[test]
fn decoding_does_not_replace_explicit_source_limits_with_defaults() {
    let encoded = wire(&fresh(fixture()), "changed");
    let (before, _) = xmls(&encoded);
    let value = "x".repeat(SourceLimits::default().max_xml_attribute_bytes() + 1);
    let xml = String::from_utf8(before.to_vec()).unwrap().replace(
        "</p:sld>",
        &format!("<p:extLst><p:ext uri=\"urn:opaque\"><x:meta xmlns:x=\"urn:opaque\" value=\"{value}\"/></p:ext></p:extLst></p:sld>"),
    );
    let bytes = with_first_slide(xml.as_bytes());
    let limits = SourceLimits::builder()
        .max_xml_attribute_bytes(value.len() + 1024)
        .unwrap()
        .build()
        .unwrap();
    let projection = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(bytes.clone())),
        crate::OpenOptions::default().source_limits(limits),
    )
    .unwrap();
    let source = Package::from_presentation(&projection).unwrap();
    let patch = Patch::from_bytes(&wire(&source, "portable")).unwrap();
    assert_eq!(
        text(&source.apply(&patch).unwrap().into_package(), 0),
        "portable"
    );
    assert!(fresh(bytes).apply(&patch).is_err());
}

#[test]
fn durable_root_proof_rejects_changes_outside_slide_order() {
    let encoded = wire(&fresh(fixture()), "changed");
    let after_start = entry_start(&encoded) - read_u64(&encoded, 60);
    let after_end = entry_start(&encoded);
    let after = String::from_utf8(encoded[after_start..after_end].to_vec()).unwrap();
    let forged = after.replace("</p:presentation>", "<!-- injected --></p:presentation>");
    assert_ne!(forged, after);
    let mut bytes = encoded[..after_start].to_vec();
    bytes[60..68].copy_from_slice(&(forged.len() as u64).to_le_bytes());
    bytes.extend_from_slice(forged.as_bytes());
    bytes.extend_from_slice(&encoded[after_end..]);
    assert!(Patch::from_bytes(&bytes).is_err());
}

#[test]
fn durable_text_and_order_patches_allow_large_unchanged_catalogs() {
    let mut builder = crate::builder::Builder::new().unwrap();
    for _ in 0..40 {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box("original", 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let source = fresh(bytes.clone());
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(39, 0_usize, "last").unwrap();
    edit.move_slide(Position::new(39), Position::new(0))
        .unwrap();
    let patch = Patch::from_bytes(&edit.commit().unwrap().patch().to_bytes().unwrap()).unwrap();
    let replayed = fresh(bytes).apply(&patch).unwrap().into_package();
    assert_eq!(text(&replayed, 0), "last");
    assert_eq!(replayed.presentation().unwrap().slide_count().unwrap(), 40);
}

#[test]
fn durable_text_only_replay_keeps_positional_owner_contexts_supported() {
    let mut raw = litchi_opc::OpcPackage::from_bytes(&fixture(), SourceLimits::default()).unwrap();
    let root = raw
        .get_part_mut(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap();
    let xml = String::from_utf8(root.blob().to_vec())
        .unwrap()
        .replace("</p:presentation>", "<p:custShowLst/></p:presentation>");
    root.set_blob(xml.into_bytes());
    let bytes = litchi_opc::PackageWriter::to_bytes(&raw).unwrap();
    let source = fresh(bytes.clone());
    let patch = Patch::from_bytes(&wire(&source, "text only")).unwrap();
    let changed = fresh(bytes.clone()).apply(&patch).unwrap().into_package();
    assert_eq!(text(&changed, 0), "text only");
    assert_eq!(
        changed
            .apply(&patch.inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

#[test]
fn durable_decode_rejects_redundant_slide_deltas() {
    let encoded = wire(&fresh(fixture()), "changed");
    let (before, _) = xmls(&encoded);
    assert!(Patch::from_bytes(&with_xmls(&encoded, before, before)).is_err());
}

#[test]
fn durable_catalog_counts_are_bounded_before_reserving() {
    let encoded = wire(&fresh(fixture()), "changed");
    let mut oversized = encoded.clone();
    oversized[40..44].copy_from_slice(&100_001_u32.to_le_bytes());
    assert!(matches!(
        Patch::from_bytes(&oversized),
        Err(Error::Limit { .. })
    ));
    let mut root_consumes_frame = encoded;
    let remaining = root_consumes_frame.len() - 68;
    root_consumes_frame[52..60].copy_from_slice(&(remaining as u64).to_le_bytes());
    assert!(Patch::from_bytes(&root_consumes_frame).is_err());
}
