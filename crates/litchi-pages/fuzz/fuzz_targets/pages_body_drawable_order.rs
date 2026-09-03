#![no_main]

//! Bounded owner-level fuzzing for Pages body drawable stacking order.
//!
//! Inputs are small fixture descriptors rather than native package bytes.  A
//! descriptor still controls the physical object order, semantic drawable
//! permutation, optional body-storage slot, unknown fields, malformed native
//! variants, and the public selector operation.  This makes the full package
//! owner reachable during ordinary libFuzzer mutation while keeping every
//! generated package finite and reviewable.

use std::{hint::black_box, io, sync::OnceLock};

use libfuzzer_sys::fuzz_target;
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{pages_drawable_order_codec, tp, tsa, tsp};
use litchi_pages::{BodyDrawableSelector, DrawableLayerMove, Limits, Package, Position};
use prost::Message as _;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_OUTPUT_BYTES: usize = 128 * 1024;
const MAX_FIELDS: usize = 16 * 1024;
const MAX_WORK_BYTES: usize = 512 * 1024;
const MAX_REFERENCES: usize = 4 * 1024;
const MAX_RECURSION: u32 = 64;
const MAX_DRAWABLES: usize = 16;
const MAX_DESCRIPTOR_PAYLOAD_BYTES: usize = 64;

const ROOT_IDENTIFIER: u64 = 1;
const ZORDER_IDENTIFIER: u64 = 46;
const BODY_STORAGE_IDENTIFIER: u64 = 47;
const UNRELATED_IDENTIFIER: u64 = 700;
const DRAWABLE_BASE: u64 = 100;

const ROOT_MESSAGE_TYPE: u32 = 10_000;
const ZORDER_MESSAGE_TYPE: u32 = 10_015;
const WRONG_ZORDER_MESSAGE_TYPE: u32 = 10_016;
const BODY_STORAGE_MESSAGE_TYPE: u32 = 2_001;
const DRAWABLE_MESSAGE_TYPE: u32 = 9_000;

const MODE_VALID: u8 = 0;
const MODE_DUPLICATE_ORDER_REFERENCE: u8 = 1;
const MODE_DUPLICATE_ROOT_ZORDER: u8 = 2;
const MODE_WRONG_ZORDER_MESSAGE_TYPE: u8 = 3;
const MODE_MISSING_ZORDER_OBJECT: u8 = 4;
const MODE_ZERO_IDENTIFIER: u8 = 5;
const MODE_MISSING_IDENTIFIER: u8 = 6;
const MODE_EXTERNAL_REFERENCE: u8 = 7;
const MODE_WRONG_ORDER_WIRE: u8 = 8;
const MODE_NONCANONICAL_IDENTIFIER: u8 = 9;
const MODE_DUPLICATE_BODY_REFERENCE: u8 = 10;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SENTINEL_MEMBER: &str = "Data/drawable-order-sentinel.bin";
const PREVIEW_MEMBERS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type BuildResult<T> = Result<T, Box<dyn std::error::Error>>;

fuzz_target!(|data: &[u8]| {
    let Some(descriptor) = normalize_input(data) else {
        return;
    };

    // Keep the ZIP ingress itself in the fuzz loop when the mutator happens
    // to produce a package-shaped input.  Most inputs are descriptors below,
    // so this path is deliberately best-effort and never retains them.
    if descriptor.starts_with(b"PK") {
        let _ = black_box(Package::from_bytes_with_limits(
            &descriptor,
            Limits::default(),
        ));
    }

    exercise_descriptor(&descriptor);

    // Descriptor mutation is intentionally biased toward malformed modes.
    // Run a small deterministic valid set once per worker so semantic moves,
    // inverse patches, and the native body-storage slot remain hot even when
    // a campaign has not yet discovered those mode bytes.
    static FIXTURES: OnceLock<()> = OnceLock::new();
    FIXTURES.get_or_init(exercise_known_descriptors);
});

fn normalize_input(data: &[u8]) -> Option<Vec<u8>> {
    if let Some(encoded) = data.strip_prefix(b"hex:") {
        return decode_hex(encoded);
    }
    (data.len() <= MAX_INPUT_BYTES).then(|| data.to_vec())
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    if encoded.len() > MAX_INPUT_BYTES.saturating_mul(2).saturating_add(16) {
        return None;
    }
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = hex_nibble(byte)?;
        if let Some(high_nibble) = high.take() {
            output.push((high_nibble << 4) | nibble);
            if output.len() > MAX_INPUT_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn exercise_descriptor(descriptor: &[u8]) {
    let mode = descriptor.get(3).copied().unwrap_or_default() % 11;
    let Ok(package_bytes) = fixture_bytes(descriptor, mode) else {
        return;
    };
    exercise_package(&package_bytes, descriptor, mode);
}

fn exercise_known_descriptors() {
    for descriptor in [
        &[3, 0, 1, MODE_VALID, 0, 0, 0, 0, 0, 0][..],
        &[4, 1, 2, MODE_VALID, 1, 1, 0, 0, 0, 0, 0, 0][..],
        &[6, 1, 4, MODE_VALID, 0, 1, 0, 0, 0, 0, 0, 0][..],
        &[3, 1, 0, MODE_DUPLICATE_ORDER_REFERENCE, 0, 0][..],
        &[3, 0, 0, MODE_DUPLICATE_ROOT_ZORDER, 0, 0][..],
        &[3, 1, 0, MODE_WRONG_ZORDER_MESSAGE_TYPE, 0, 0][..],
        &[3, 0, 0, MODE_MISSING_ZORDER_OBJECT, 0, 0][..],
        &[3, 1, 0, MODE_ZERO_IDENTIFIER, 0, 0][..],
        &[3, 1, 0, MODE_MISSING_IDENTIFIER, 0, 0][..],
        &[3, 1, 0, MODE_EXTERNAL_REFERENCE, 0, 0][..],
        &[3, 1, 0, MODE_WRONG_ORDER_WIRE, 0, 0][..],
        &[3, 1, 0, MODE_NONCANONICAL_IDENTIFIER, 0, 0][..],
        &[3, 1, 0, MODE_DUPLICATE_BODY_REFERENCE, 0, 0][..],
    ] {
        let mode = descriptor[3];
        if let Ok(package_bytes) = fixture_bytes(descriptor, mode) {
            exercise_package(&package_bytes, descriptor, mode);
        }
    }
}

fn exercise_package(package_bytes: &[u8], descriptor: &[u8], mode: u8) {
    let parsed = Package::from_bytes_with_limits(package_bytes, Limits::default());
    let Ok(package) = parsed else {
        return;
    };
    let Ok(source_bytes) = exact_bytes(&package) else {
        return;
    };
    let Ok(handles) = package.body_drawable_order() else {
        // All malformed modes are expected to fail closed at this boundary;
        // a valid descriptor is handled below with stronger invariants.
        return;
    };
    assert_positions(&handles);

    let with_body = descriptor.get(1).copied().unwrap_or_default() & 1 != 0;
    let before_native = match native_order(&package) {
        Ok(order) => order,
        Err(error) if mode == MODE_VALID => {
            panic!("valid fixture native order must decode: {error}")
        },
        Err(_) => return,
    };
    if mode != MODE_VALID {
        // A malformed mode may still be accepted by package ingress when its
        // root shape is independently valid.  It must not make a successful
        // semantic read lose the structural body slot.
        if with_body {
            assert_body_slot(&before_native, &before_native);
        }
        black_box(handles);
        return;
    }

    if with_body {
        assert_eq!(
            before_native
                .iter()
                .filter(|identifier| **identifier == BODY_STORAGE_IDENTIFIER)
                .count(),
            1,
            "valid body fixture must contain one native body slot"
        );
    }
    assert_eq!(
        handles.len(),
        before_native
            .iter()
            .filter(|identifier| { !with_body || **identifier != BODY_STORAGE_IDENTIFIER })
            .count()
    );

    if handles.is_empty() {
        let Ok(mut edit) = package.edit_body_drawable_order() else {
            return;
        };
        if edit.set_order(&[]).is_err() {
            return;
        }
        let Ok(commit) = edit.commit() else {
            return;
        };
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            exact_bytes(commit.package()).ok(),
            Some(source_bytes.to_vec())
        );
        return;
    }

    let requested_kind = descriptor.get(4).copied().unwrap_or_default() & 1;
    let index = usize::from(descriptor.get(5).copied().unwrap_or_default()) % handles.len();
    let mut expected_drawables: Vec<u64> = before_native
        .iter()
        .copied()
        .filter(|identifier| !with_body || *identifier != BODY_STORAGE_IDENTIFIER)
        .collect();
    let changed = if requested_kind == 0 {
        let movement = movement(descriptor.get(6).copied().unwrap_or_default());
        let expected_changed = move_drawable(&mut expected_drawables, index, movement);
        let Ok(mut edit) = package.edit_body_drawable_order() else {
            return;
        };
        let actual_changed = match edit.move_drawable(
            BodyDrawableSelector::position(Position::new(index)),
            movement,
        ) {
            Ok(changed) => changed,
            Err(_) => return,
        };
        assert_eq!(actual_changed, expected_changed);
        match edit.commit() {
            Ok(commit) => verify_commit(
                &package,
                &source_bytes,
                &before_native,
                &expected_drawables,
                with_body,
                expected_changed,
                commit,
            ),
            Err(_) => return,
        }
        expected_changed
    } else {
        let mut requested = handles.clone();
        if descriptor.get(6).copied().unwrap_or_default() & 1 == 0 {
            requested.reverse();
            expected_drawables.reverse();
        } else {
            requested.rotate_left(1);
            expected_drawables.rotate_left(1);
        }
        let expected_changed = expected_drawables
            != before_native
                .iter()
                .copied()
                .filter(|identifier| !with_body || *identifier != BODY_STORAGE_IDENTIFIER)
                .collect::<Vec<_>>();
        let Ok(mut edit) = package.edit_body_drawable_order() else {
            return;
        };
        if edit.set_order(&requested).is_err() {
            return;
        }
        let Ok(commit) = edit.commit() else {
            return;
        };
        verify_commit(
            &package,
            &source_bytes,
            &before_native,
            &expected_drawables,
            with_body,
            expected_changed,
            commit,
        );
        expected_changed
    };

    black_box(changed);
}

fn verify_commit(
    source: &Package,
    source_bytes: &[u8],
    before_native: &[u64],
    expected_drawables: &[u64],
    with_body: bool,
    expected_changed: bool,
    commit: litchi_pages::BodyDrawableOrderCommit,
) {
    let after_native = native_order(commit.package())
        .expect("successful body-order commit must preserve a decodable native order");
    let expected_native = with_body
        .then(|| insert_body_slot(before_native, expected_drawables))
        .unwrap_or_else(|| expected_drawables.to_vec());
    assert_eq!(after_native, expected_native);
    if with_body {
        assert_body_slot(before_native, &after_native);
    }
    if expected_changed {
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(
            commit.diagnostics().deleted_previews(),
            PREVIEW_MEMBERS.len()
        );
    } else {
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert!(commit.patch().is_noop());
        assert_eq!(
            exact_bytes(commit.package()).ok(),
            Some(source_bytes.to_vec())
        );
        assert_eq!(exact_bytes(source).ok(), Some(source_bytes.to_vec()));
        return;
    }

    let inverse = commit
        .patch()
        .try_inverse()
        .expect("changed body-order commit must provide an inverse patch");
    let restored = commit
        .package()
        .apply_body_drawable_order(&inverse)
        .expect("inverse body-order patch must apply");
    assert_eq!(
        exact_bytes(restored.package()).ok(),
        Some(source_bytes.to_vec())
    );
    assert_eq!(restored.diagnostics().deleted_previews(), 0);
    assert_eq!(exact_bytes(source).ok(), Some(source_bytes.to_vec()));
}

fn assert_positions(handles: &[litchi_pages::BodyDrawableHandle]) {
    for (index, handle) in handles.iter().enumerate() {
        assert_eq!(handle.position(), Position::new(index));
    }
}

fn assert_body_slot(before: &[u64], after: &[u64]) {
    assert_eq!(
        before
            .iter()
            .position(|identifier| *identifier == BODY_STORAGE_IDENTIFIER),
        after
            .iter()
            .position(|identifier| *identifier == BODY_STORAGE_IDENTIFIER),
        "body storage moved out of its structural z-order slot"
    );
}

fn insert_body_slot(before: &[u64], drawables: &[u64]) -> Vec<u64> {
    let Some(slot) = before
        .iter()
        .position(|identifier| *identifier == BODY_STORAGE_IDENTIFIER)
    else {
        return drawables.to_vec();
    };
    let mut output = drawables.to_vec();
    output.insert(slot.min(output.len()), BODY_STORAGE_IDENTIFIER);
    output
}

fn movement(value: u8) -> DrawableLayerMove {
    match value % 4 {
        0 => DrawableLayerMove::ToBack,
        1 => DrawableLayerMove::Backward,
        2 => DrawableLayerMove::Forward,
        _ => DrawableLayerMove::ToFront,
    }
}

fn move_drawable(order: &mut Vec<u64>, index: usize, movement: DrawableLayerMove) -> bool {
    let Some(final_index) = order.len().checked_sub(1) else {
        return false;
    };
    let target = match movement {
        DrawableLayerMove::ToBack => 0,
        DrawableLayerMove::Backward => index.saturating_sub(1),
        DrawableLayerMove::Forward => index.saturating_add(1).min(final_index),
        DrawableLayerMove::ToFront => final_index,
        _ => return false,
    };
    if target == index {
        return false;
    }
    let value = order.remove(index);
    order.insert(target, value);
    true
}

fn fixture_bytes(descriptor: &[u8], mode: u8) -> BuildResult<Vec<u8>> {
    let count = usize::from(descriptor.first().copied().unwrap_or_default()) % (MAX_DRAWABLES + 1);
    let with_body = descriptor.get(1).copied().unwrap_or_default() & 1 != 0;
    let body_slot = with_body
        .then(|| usize::from(descriptor.get(2).copied().unwrap_or_default()) % (count + 1));
    let mut drawables: Vec<u64> = (0..count)
        .map(|index| DRAWABLE_BASE.saturating_add(index as u64))
        .collect();
    if descriptor.get(7).copied().unwrap_or_default() & 1 == 1 {
        drawables.reverse();
    }
    if !drawables.is_empty() {
        let rotation =
            usize::from(descriptor.get(8).copied().unwrap_or_default()) % drawables.len();
        drawables.rotate_left(rotation);
    }
    let mut native_order = drawables.clone();
    if let Some(position) = body_slot {
        native_order.insert(position, BODY_STORAGE_IDENTIFIER);
    }

    let mut root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: with_body.then(|| reference(BODY_STORAGE_IDENTIFIER)),
        drawables_zorder: Some(reference(ZORDER_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    }
    .encode_to_vec();
    append_varint_field(&mut root, 77, descriptor_value(descriptor, 9))?;
    if mode == MODE_DUPLICATE_ROOT_ZORDER {
        append_length_delimited_field(
            &mut root,
            20,
            &reference(ZORDER_IDENTIFIER).encode_to_vec(),
        )?;
    }
    if mode == MODE_DUPLICATE_BODY_REFERENCE && with_body {
        append_length_delimited_field(
            &mut root,
            4,
            &reference(BODY_STORAGE_IDENTIFIER).encode_to_vec(),
        )?;
    }
    if mode == MODE_WRONG_ORDER_WIRE {
        append_varint_field(&mut root, 20, ZORDER_IDENTIFIER)?;
    }

    let order_payload = order_payload(&native_order, descriptor, mode)?;
    let mut objects = Vec::new();
    objects.push(object(ROOT_IDENTIFIER, ROOT_MESSAGE_TYPE, root)?);
    // Keep physical object order independent from semantic z-order so the
    // owner must resolve the complete bounded component catalog.
    objects.push(object(
        UNRELATED_IDENTIFIER,
        DRAWABLE_MESSAGE_TYPE,
        descriptor_payload(descriptor, 10),
    )?);
    if mode != MODE_MISSING_ZORDER_OBJECT {
        objects.push(object(
            ZORDER_IDENTIFIER,
            if mode == MODE_WRONG_ZORDER_MESSAGE_TYPE {
                WRONG_ZORDER_MESSAGE_TYPE
            } else {
                ZORDER_MESSAGE_TYPE
            },
            order_payload,
        )?);
    }
    if with_body {
        objects.push(object(
            BODY_STORAGE_IDENTIFIER,
            BODY_STORAGE_MESSAGE_TYPE,
            Vec::new(),
        )?);
    }
    for identifier in drawables {
        objects.push(object(
            identifier,
            DRAWABLE_MESSAGE_TYPE,
            descriptor_payload(descriptor, identifier as usize),
        )?);
    }
    let archive_bytes = Archive { objects }.to_bytes()?;
    let compressed = SnappyStream::compress(&archive_bytes)?;

    let sentinel = descriptor_payload(descriptor, 12);
    let preview_full = descriptor_payload(descriptor, 13);
    let preview_micro = descriptor_payload(descriptor, 14);
    let preview_web = descriptor_payload(descriptor, 15);
    let entries = vec![
        (SENTINEL_MEMBER, sentinel),
        (PREVIEW_MEMBERS[0], preview_full),
        (PREVIEW_MEMBERS[1], preview_micro),
        (PREVIEW_MEMBERS[2], preview_web),
        (DOCUMENT_MEMBER, compressed.as_slice().to_vec()),
    ];
    let entries = entries
        .iter()
        .map(|(name, data)| (*name, data.as_slice()))
        .collect::<Vec<_>>();
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn order_payload(native_order: &[u64], descriptor: &[u8], mode: u8) -> BuildResult<Vec<u8>> {
    if mode == MODE_WRONG_ORDER_WIRE {
        return Ok(vec![0x08, 0x01]);
    }
    let mut payload = Vec::new();
    for (position, identifier) in native_order.iter().copied().enumerate() {
        let mut nested = if position == 0 && mode == MODE_ZERO_IDENTIFIER {
            reference_bytes(0)
        } else if position == 0 && mode == MODE_MISSING_IDENTIFIER {
            Vec::new()
        } else if position == 0 && mode == MODE_NONCANONICAL_IDENTIFIER {
            vec![0x08, 0x81, 0x00]
        } else {
            reference_bytes(identifier)
        };
        append_varint_field(&mut nested, 99, descriptor_value(descriptor, 16 + position))?;
        if position == 0 && mode == MODE_EXTERNAL_REFERENCE {
            append_varint_field(&mut nested, 3, 1)?;
        }
        append_length_delimited_field(&mut payload, 1, &nested)?;
    }
    if mode == MODE_DUPLICATE_ORDER_REFERENCE {
        let duplicate = native_order
            .first()
            .copied()
            .map(reference_bytes)
            .unwrap_or_default();
        append_length_delimited_field(&mut payload, 1, &duplicate)?;
    }
    append_varint_field(&mut payload, 77, descriptor_value(descriptor, 24))?;
    Ok(payload)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn reference_bytes(identifier: u64) -> Vec<u8> {
    reference(identifier).encode_to_vec()
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> BuildResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn exact_bytes(package: &Package) -> BuildResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn native_order(package: &Package) -> BuildResult<Vec<u64>> {
    let bytes = exact_bytes(package)?;
    let catalog = Catalog::from_bytes(&bytes)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("Pages document member is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(&stream.into_bytes())?;
    let object = archive
        .object(ZORDER_IDENTIFIER)
        .ok_or_else(|| io::Error::other("Pages z-order object is missing"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == ZORDER_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("Pages z-order message is missing"))?;
    let snapshot = pages_drawable_order_codec::decode_drawable_order(
        &message.data,
        pages_drawable_order_codec::DecodeOptions::new(
            MAX_INPUT_BYTES,
            MAX_OUTPUT_BYTES,
            MAX_FIELDS,
            MAX_WORK_BYTES,
            MAX_RECURSION,
            MAX_REFERENCES,
        ),
    )?;
    Ok(snapshot.identifiers().collect())
}

fn descriptor_value(descriptor: &[u8], offset: usize) -> u64 {
    let mut value = 0_u64;
    for shift in 0..8 {
        value |=
            u64::from(descriptor.get(offset + shift).copied().unwrap_or_default()) << (shift * 8);
    }
    value
}

fn descriptor_payload(descriptor: &[u8], offset: usize) -> Vec<u8> {
    descriptor
        .get(offset..)
        .unwrap_or_default()
        .iter()
        .copied()
        .take(MAX_DESCRIPTOR_PAYLOAD_BYTES)
        .collect()
}
