use std::io::Cursor;

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_core::patch::{BlobLimits, Patch as CorePatch, PatchLimits, Reversible};
use litchi_docx::{Package, web_extensions as web};
use litchi_opc::constants::relationship_type as rt;
use litchi_opc::{BlobPart, PackURI};
use serde_json::Value;

use super::*;

fn durable_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(16, 64 * 1024 * 1024, 128 * 1024 * 1024),
        32 * 1024 * 1024,
        64,
        32,
        1024 * 1024,
        64 * 1024 * 1024,
    )
}

fn zero_blob_noop_limits() -> PatchLimits {
    PatchLimits::new(
        BlobLimits::new(0, 0, 0),
        1024 * 1024,
        4,
        32,
        1024 * 1024,
        1024 * 1024,
    )
}

fn durable_patch(patch: &web::Patch) -> CorePatch<Reversible> {
    let durable = patch.to_durable(durable_limits()).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded =
        CorePatch::<Reversible>::from_deterministic_json(&wire, durable_limits()).unwrap();
    assert_eq!(wire, decoded.to_deterministic_json().unwrap());
    decoded
}

fn apply_durable(
    package: &mut Package,
    patch: &CorePatch<Reversible>,
) -> litchi_docx::Result<bool> {
    package.apply_durable_task_panes_patch(patch)
}

fn apply_durable_with_limits(
    package: &mut Package,
    patch: &CorePatch<Reversible>,
    limits: &web::Limits,
) -> litchi_docx::Result<bool> {
    package.apply_durable_task_panes_patch_with_limits(patch, limits)
}

fn apply_and_inverse(original: &[u8], durable: &CorePatch<Reversible>) {
    let mut package = Package::from_reader(Cursor::new(original)).unwrap();
    assert!(apply_durable(&mut package, durable).unwrap());
    let published = package_bytes(&package);

    let mut reopened = Package::from_reader(Cursor::new(published)).unwrap();
    assert!(apply_durable(&mut reopened, &durable.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&reopened)), members(original));
}

fn assert_rejected_unchanged(
    source: &[u8],
    durable: &CorePatch<Reversible>,
    label: &str,
) -> String {
    let mut package = Package::from_reader(Cursor::new(source)).unwrap();
    let before = members(&package_bytes(&package));
    let error = apply_durable(&mut package, durable)
        .err()
        .unwrap_or_else(|| panic!("{label}: durable patch was accepted"));
    assert_eq!(members(&package_bytes(&package)), before, "{label}");
    error.to_string()
}

fn wire_blob(wire: &[u8], collection: &str) -> (String, String, Vec<u8>) {
    let root: Value = serde_json::from_slice(wire).unwrap();
    let blob = root
        .get(collection)
        .and_then(Value::as_array)
        .and_then(|blobs| blobs.first())
        .and_then(Value::as_object)
        .unwrap();
    let encoded = blob
        .get("bytes")
        .and_then(Value::as_str)
        .unwrap()
        .to_owned();
    let digest = blob
        .get("sha256")
        .and_then(Value::as_str)
        .unwrap()
        .to_owned();
    let bytes = BASE64.decode(&encoded).unwrap();
    (encoded, digest, bytes)
}

fn wire_target_hash(wire: &[u8]) -> String {
    let root: Value = serde_json::from_slice(wire).unwrap();
    root.get("operations")
        .and_then(Value::as_array)
        .and_then(|operations| operations.first())
        .and_then(Value::as_object)
        .and_then(|operation| operation.get("forward"))
        .and_then(Value::as_object)
        .and_then(|operation| operation.get("preconditions"))
        .and_then(Value::as_object)
        .and_then(|preconditions| preconditions.get("target_sha256"))
        .and_then(Value::as_str)
        .unwrap()
        .to_owned()
}

fn sha256_hex(bytes: &[u8]) -> String {
    use sha2::{Digest as _, Sha256};
    let mut result = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        use std::fmt::Write as _;
        write!(result, "{byte:02x}").unwrap();
    }
    result
}

fn durable_replace_once(bytes: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    let index = bytes
        .windows(from.len())
        .position(|window| window == from)
        .unwrap_or_else(|| panic!("missing durable fixture marker {from:?}"));
    let mut result = Vec::with_capacity(bytes.len() + to.len().saturating_sub(from.len()));
    result.extend_from_slice(&bytes[..index]);
    result.extend_from_slice(to);
    result.extend_from_slice(&bytes[index + from.len()..]);
    result
}

fn replace_wire_blob(wire: &[u8], collection: &str, precondition: &str, bytes: &[u8]) -> Vec<u8> {
    let (old_encoded, old_digest, _) = wire_blob(wire, collection);
    let new_encoded = BASE64.encode(bytes);
    let new_digest = sha256_hex(bytes);
    let old_entry = format!(r#"{{"bytes":"{old_encoded}","sha256":"{old_digest}"}}"#);
    let new_entry = format!(r#"{{"bytes":"{new_encoded}","sha256":"{new_digest}"}}"#);
    let mut result = durable_replace_once(wire, old_entry.as_bytes(), new_entry.as_bytes());
    let old_precondition = format!(r#""{precondition}":"{old_digest}""#);
    let new_precondition = format!(r#""{precondition}":"{new_digest}""#);
    result = durable_replace_once(
        &result,
        old_precondition.as_bytes(),
        new_precondition.as_bytes(),
    );
    result
}

fn append_binary_closure_record(
    restore: &[u8],
    name: &str,
    content_type: &str,
    before_payload: &[u8],
    relationship_present: bool,
    relationship_bytes: &[u8],
    after_payload: &[u8],
) -> Vec<u8> {
    fn read_u64(bytes: &[u8], offset: &mut usize) -> usize {
        let end = (*offset).checked_add(8).unwrap();
        let value = u64::from_le_bytes(bytes[*offset..end].try_into().unwrap());
        *offset = end;
        usize::try_from(value).unwrap()
    }

    fn write_u64(bytes: &mut Vec<u8>, value: usize) {
        bytes.extend_from_slice(&u64::try_from(value).unwrap().to_le_bytes());
    }

    fn write_bytes(bytes: &mut Vec<u8>, value: &[u8]) {
        write_u64(bytes, value.len());
        bytes.extend_from_slice(value);
    }

    assert!(restore.starts_with(b"WCR1"));
    let mut offset = 4;
    let intent_len = read_u64(restore, &mut offset);
    let intent_end = offset.checked_add(intent_len).unwrap();
    let intent = &restore[offset..intent_end];
    offset = intent_end;
    let closure_len = read_u64(restore, &mut offset);
    let closure_end = offset.checked_add(closure_len).unwrap();
    assert_eq!(closure_end, restore.len());
    let closure = &restore[offset..closure_end];
    assert!(closure.starts_with(b"WCC1"));

    let count = u64::from_le_bytes(closure[4..12].try_into().unwrap());
    let mut forged_closure = closure.to_vec();
    forged_closure[4..12].copy_from_slice(&count.checked_add(1).unwrap().to_le_bytes());
    forged_closure.push(2);
    write_bytes(&mut forged_closure, name.as_bytes());
    forged_closure.push(1);
    write_bytes(&mut forged_closure, content_type.as_bytes());
    write_bytes(&mut forged_closure, before_payload);
    forged_closure.push(u8::from(relationship_present));
    write_bytes(&mut forged_closure, relationship_bytes);
    forged_closure.push(1);
    write_bytes(&mut forged_closure, content_type.as_bytes());
    write_bytes(&mut forged_closure, after_payload);
    forged_closure.push(u8::from(relationship_present));
    write_bytes(&mut forged_closure, relationship_bytes);

    let mut result = Vec::with_capacity(restore.len() + forged_closure.len() - closure.len());
    result.extend_from_slice(b"WCR1");
    write_u64(&mut result, intent.len());
    result.extend_from_slice(intent);
    write_u64(&mut result, forged_closure.len());
    result.extend_from_slice(&forged_closure);
    result
}

fn replace_wire_target(wire: &[u8], target: &str) -> Vec<u8> {
    let old_target = wire_target_hash(wire);
    let old_marker = format!(r#""target_sha256":"{old_target}""#);
    let new_marker = format!(r#""target_sha256":"{target}""#);
    durable_replace_once(wire, old_marker.as_bytes(), new_marker.as_bytes())
}

fn decode_wire(wire: &[u8]) -> CorePatch<Reversible> {
    CorePatch::<Reversible>::from_deterministic_json(wire, durable_limits()).unwrap()
}

fn source_with_unrelated_add_in_markup() -> Vec<u8> {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    rewrite_member(&source, "webextensions/webextension1.xml", |content| {
        insert_before(
            content,
            b"</we:extLst>",
            b"<!-- unrelated durable source -->",
        )
    })
}

#[test]
fn durable_creation_roundtrips_and_inverse_restores_exact_members_after_reopen() {
    let (panes, _) = authored_panes();
    let source = Package::new().unwrap();
    let original = package_bytes(&source);
    let patch = source
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&patch);
    apply_and_inverse(&original, &durable);
}

#[test]
fn durable_custom_only_two_pane_edit_roundtrips_and_inverse_is_exact() {
    let mut source = Package::new().unwrap();
    source
        .put_task_panes(two_authored_panes(), web::Conformance::Transitional)
        .unwrap();
    let original = package_bytes(&source);
    let source = Package::from_reader(Cursor::new(&original)).unwrap();
    let desired = panes_with_metadata(&source, changed_custom_functions());
    let patch = source
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&patch);
    apply_and_inverse(&original, &durable);
}

#[test]
fn durable_transitional_custom_edit_ignores_unrelated_strict_namespace_text() {
    let (panes, _) = authored_panes();
    let mut authored = Package::new().unwrap();
    authored
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let source = package_bytes(&authored);
    let original = rewrite_member(&source, "word/document.xml", |content| {
        insert_before(
            content,
            b"</w:body>",
            b"<!-- unrelated strict namespace text: http://purl.oclc.org/ooxml/officeDocument/relationships -->",
        )
    });
    let mut package = Package::from_reader(Cursor::new(&original)).unwrap();
    let desired = panes_with_metadata(&package, changed_custom_functions());
    let ordinary = package
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    assert!(!ordinary.is_empty());
    let durable = durable_patch(&ordinary);
    assert!(apply_durable(&mut package, &durable).unwrap());

    let published = members(&package_bytes(&package));
    assert_eq!(
        published.get("word/document.xml"),
        members(&original).get("word/document.xml")
    );
    let mut reopened = Package::from_reader(Cursor::new(package_bytes(&package))).unwrap();
    assert!(apply_durable(&mut reopened, &durable.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&reopened)), members(&original));
}

#[test]
fn durable_removal_roundtrips_and_inverse_restores_exact_members_after_reopen() {
    let (panes, _) = authored_panes();
    let mut source = Package::new().unwrap();
    source
        .put_task_panes(panes, web::Conformance::Strict)
        .unwrap();
    let original = package_bytes(&source);
    let source = Package::from_reader(Cursor::new(&original)).unwrap();
    let patch = source.plan_remove_task_panes().unwrap();
    let durable = durable_patch(&patch);
    apply_and_inverse(&original, &durable);
}

#[test]
fn durable_noop_is_deterministic_and_preserves_signed_source() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    package
        .edit_opc(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .unwrap();
    let before = members(&package_bytes(&package));
    let loaded = package.task_panes().unwrap().unwrap();
    let patch = package
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(patch.is_empty());
    let durable = durable_patch(&patch);
    let wire = durable.to_deterministic_json().unwrap();
    assert_eq!(wire, decode_wire(&wire).to_deterministic_json().unwrap());
    assert!(!apply_durable(&mut package, &durable).unwrap());
    assert!(package.is_signed());
    assert_eq!(members(&package_bytes(&package)), before);
}

#[test]
fn durable_exact_noop_exports_with_zero_blob_allowance() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let loaded = package.task_panes().unwrap().unwrap();
    let ordinary = package
        .plan_task_panes(loaded, web::Conformance::Transitional)
        .unwrap();
    assert!(ordinary.is_empty());
    let limits = zero_blob_noop_limits();
    let durable = ordinary.to_durable(limits).unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    let decoded = CorePatch::<Reversible>::from_deterministic_json(&wire, limits).unwrap();
    let before = members(&package_bytes(&package));
    assert!(!apply_durable(&mut package, &decoded).unwrap());
    assert_eq!(members(&package_bytes(&package)), before);
}

#[test]
fn durable_exported_ordinary_inverse_applies_after_forward_publication() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    let original = package_bytes(&package);
    let ordinary = package
        .plan_task_panes(panes, web::Conformance::Strict)
        .unwrap();
    let forward = durable_patch(&ordinary);
    let inverse = durable_patch(&ordinary.inverse());
    assert!(apply_durable(&mut package, &forward).unwrap());
    assert!(apply_durable(&mut package, &inverse).unwrap());
    assert_eq!(members(&package_bytes(&package)), members(&original));
}

#[test]
fn durable_creation_and_removal_roundtrip_in_both_conformances() {
    for conformance in [web::Conformance::Transitional, web::Conformance::Strict] {
        let (panes, _) = authored_panes();
        let source = Package::new().unwrap();
        let original = package_bytes(&source);
        let creation = source.plan_task_panes(panes, conformance).unwrap();
        apply_and_inverse(&original, &durable_patch(&creation));

        let mut populated = Package::new().unwrap();
        populated
            .put_task_panes(authored_panes().0, conformance)
            .unwrap();
        let original = package_bytes(&populated);
        let source = Package::from_reader(Cursor::new(&original)).unwrap();
        let removal = source.plan_remove_task_panes().unwrap();
        apply_and_inverse(&original, &durable_patch(&removal));
    }
}

#[test]
fn durable_changed_signed_patch_requires_unsign_and_replan() {
    let (panes, _) = authored_panes();
    let mut package = Package::new().unwrap();
    package
        .put_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    package
        .edit_opc(|opc| {
            opc.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
            Ok(())
        })
        .unwrap();
    let changed = changed_panes(&package);
    let ordinary = package
        .plan_task_panes(changed.clone(), web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&ordinary);
    let before = members(&package_bytes(&package));
    let error = apply_durable(&mut package, &durable).unwrap_err();
    assert!(error.to_string().contains("explicit Package::unsign"));
    assert_eq!(members(&package_bytes(&package)), before);

    package.unsign();
    let after_unsign = members(&package_bytes(&package));
    let error = apply_durable(&mut package, &durable).unwrap_err();
    assert!(error.to_string().to_ascii_lowercase().contains("source"));
    assert_eq!(members(&package_bytes(&package)), after_unsign);

    let replanned = durable_patch(
        &package
            .plan_task_panes(changed, web::Conformance::Transitional)
            .unwrap(),
    );
    assert!(apply_durable(&mut package, &replanned).unwrap());
    assert!(!package.is_signed());
}

#[test]
fn durable_stale_and_wire_tampering_are_atomic() {
    let (panes, _) = authored_panes();
    let source = Package::new().unwrap();
    let original = package_bytes(&source);
    let ordinary = source
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&ordinary);

    let stale = rewrite_member(&original, "[Content_Types].xml", |content| {
        insert_before(content, b"</Types>", b"<!-- stale durable source -->")
    });
    let stale_error = assert_rejected_unchanged(&stale, &durable, "stale content types");
    assert!(
        stale_error.contains("Web Extensions durable typed replay target mismatch"),
        "unexpected stale-source error: {stale_error}"
    );

    let wire = durable.to_deterministic_json().unwrap();
    let op_tampered = decode_wire(&durable_replace_once(&wire, b"web.edit", b"web.edIt"));
    let op_error = assert_rejected_unchanged(&original, &op_tampered, "operation tamper");
    assert!(op_error.to_ascii_lowercase().contains("unsupported"));

    let target_tampered = {
        let target = wire_target_hash(&wire);
        let replacement = if let Some(suffix) = target.strip_prefix('0') {
            format!("1{suffix}")
        } else {
            format!("0{}", &target[1..])
        };
        replace_wire_target(&wire, &replacement)
    };
    let target_error =
        assert_rejected_unchanged(&original, &decode_wire(&target_tampered), "target tamper");
    assert!(target_error.to_ascii_lowercase().contains("target"));

    let (_, _, intent) = wire_blob(&wire, "forward_blobs");
    let altered_intent = durable_replace_once(&intent, b"runtime-public", b"runtime-stale");
    let blob_tampered = decode_wire(&replace_wire_blob(
        &wire,
        "forward_blobs",
        "intent_sha256",
        &altered_intent,
    ));
    let blob_error = assert_rejected_unchanged(&original, &blob_tampered, "intent tamper");
    assert!(
        blob_error.to_ascii_lowercase().contains("target"),
        "unexpected intent-tamper error: {blob_error}"
    );
}

#[test]
fn durable_altered_unrelated_restore_reaches_forward_replay_mismatch() {
    let original = source_with_unrelated_add_in_markup();
    let source = Package::from_reader(Cursor::new(&original)).unwrap();
    let desired = panes_with_metadata(&source, changed_custom_functions());
    let ordinary = source
        .plan_task_panes(desired, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&ordinary);

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    assert!(apply_durable(&mut published, &durable).unwrap());
    let published_bytes = package_bytes(&published);

    let mut exact_inverse = Package::from_reader(Cursor::new(&published_bytes)).unwrap();
    assert!(apply_durable(&mut exact_inverse, &durable.inverse()).unwrap());
    assert_eq!(members(&package_bytes(&exact_inverse)), members(&original));

    let inverse = durable.inverse();
    let inverse_wire = inverse.to_deterministic_json().unwrap();
    let (_, _, restore) = wire_blob(&inverse_wire, "forward_blobs");
    let altered_restore = durable_replace_once(
        &restore,
        b"<!-- unrelated durable source -->",
        b"<!-- corrupted durable source -->",
    );
    let altered_source = rewrite_member(&original, "webextensions/webextension1.xml", |content| {
        durable_replace_once(
            content,
            b"<!-- unrelated durable source -->",
            b"<!-- corrupted durable source -->",
        )
    });
    let altered_package = Package::from_reader(Cursor::new(&altered_source)).unwrap();
    let altered_noop = altered_package
        .plan_task_panes(
            altered_package.task_panes().unwrap().unwrap(),
            web::Conformance::Transitional,
        )
        .unwrap();
    let altered_target = wire_target_hash(
        &durable_patch(&altered_noop)
            .to_deterministic_json()
            .unwrap(),
    );
    let tampered_wire = replace_wire_target(
        &replace_wire_blob(
            &inverse_wire,
            "forward_blobs",
            "restore_sha256",
            &altered_restore,
        ),
        &altered_target,
    );
    let tampered = decode_wire(&tampered_wire);
    let mut reopened = Package::from_reader(Cursor::new(published_bytes)).unwrap();
    let before = members(&package_bytes(&reopened));
    let error = apply_durable(&mut reopened, &tampered).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("Web Extensions durable restore forward replay mismatch"),
        "unexpected altered-restore error: {error}"
    );
    assert_eq!(members(&package_bytes(&reopened)), before);
}

#[test]
fn durable_inverse_rejects_unrelated_binary_closure_record() {
    let (panes, _) = authored_panes();
    let mut source = Package::new().unwrap();
    let forged_name = PackURI::new("/custom/forged.bin").unwrap();
    let forged_content_type = "application/octet-stream";
    let original_payload = b"original durable binary";
    let forged_payload = b"forged durable binary";
    source
        .edit_opc(|opc| {
            opc.add_part(Box::new(BlobPart::new(
                forged_name.clone(),
                forged_content_type.to_owned(),
                original_payload.to_vec(),
            )));
            Ok(())
        })
        .unwrap();
    let source_part = source.opc_package().get_part(&forged_name).unwrap();
    assert_eq!(source_part.content_type(), forged_content_type);
    assert_eq!(source_part.blob(), original_payload);
    let source_relationships = source
        .opc_package()
        .source_relationships(&forged_name)
        .unwrap();
    let original = package_bytes(&source);
    let ordinary = source
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&ordinary);

    let mut published = Package::from_reader(Cursor::new(&original)).unwrap();
    assert!(apply_durable(&mut published, &durable).unwrap());
    let published_bytes = package_bytes(&published);
    let inverse = durable.inverse();
    let inverse_wire = inverse.to_deterministic_json().unwrap();
    let (_, _, restore) = wire_blob(&inverse_wire, "forward_blobs");
    let forged_restore = append_binary_closure_record(
        &restore,
        forged_name.as_str(),
        forged_content_type,
        original_payload,
        source_relationships.member_present(),
        source_relationships.bytes(),
        forged_payload,
    );
    let forged_wire = replace_wire_blob(
        &inverse_wire,
        "forward_blobs",
        "restore_sha256",
        &forged_restore,
    );
    let forged = decode_wire(&forged_wire);

    let mut reopened = Package::from_reader(Cursor::new(published_bytes)).unwrap();
    let before = members(&package_bytes(&reopened));
    let error = apply_durable(&mut reopened, &forged).unwrap_err();
    let message = error.to_string();
    assert_eq!(
        message,
        "shared OOXML error: invalid OOXML relationship: Web Extensions durable restore forward replay mismatch"
    );
    assert_eq!(members(&package_bytes(&reopened)), before);
}

#[test]
fn durable_caller_limits_and_dirty_facade_leave_packages_unchanged() {
    let (panes, _) = authored_panes();
    let source = Package::new().unwrap();
    let original = package_bytes(&source);
    let ordinary = source
        .plan_task_panes(panes, web::Conformance::Transitional)
        .unwrap();
    let durable = durable_patch(&ordinary);
    let wire = durable.to_deterministic_json().unwrap();

    let tiny = PatchLimits::new(BlobLimits::new(1, 1, 1), 1024, 1, 8, 128, 1);
    assert!(ordinary.to_durable(tiny).is_err());
    assert!(CorePatch::<Reversible>::from_deterministic_json(&wire, tiny).is_err());

    let mut item_limits = web::Limits::standard();
    item_limits.items = 0;
    let mut limited = Package::from_reader(Cursor::new(&original)).unwrap();
    let before = members(&package_bytes(&limited));
    let error = apply_durable_with_limits(&mut limited, &durable, &item_limits).unwrap_err();
    assert!(
        error.to_string().contains("task pane"),
        "unexpected item-limit error: {error}"
    );
    assert_eq!(members(&package_bytes(&limited)), before);

    let mut xml_limits = web::Limits::standard();
    xml_limits.xml_bytes = 1;
    let mut xml_limited = Package::from_reader(Cursor::new(&original)).unwrap();
    let before = members(&package_bytes(&xml_limited));
    let error = apply_durable_with_limits(&mut xml_limited, &durable, &xml_limits).unwrap_err();
    assert!(
        error.to_string().contains("Web Extensions task-pane XML"),
        "unexpected XML-limit error: {error}"
    );
    assert_eq!(members(&package_bytes(&xml_limited)), before);

    let mut dirty = Package::from_reader(Cursor::new(&original)).unwrap();
    dirty
        .document_mut()
        .unwrap()
        .add_paragraph_with_text("unmaterialized durable edit");
    let before = members(&package_bytes(&dirty));
    let error = apply_durable(&mut dirty, &durable).unwrap_err();
    assert!(error.to_string().contains("unmaterialized changes"));
    assert_eq!(members(&package_bytes(&dirty)), before);
}
