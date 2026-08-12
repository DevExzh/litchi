#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_rtf::{
    Document,
    edit::{
        Composition, CompositionLimits, Error, Limits, MAX_PICTURE_PAYLOAD_OPERATIONS,
        PicturePayloadReplacement,
    },
};
use serde_json::Value;

const PNG_A: &[u8] = &[
    0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 0x01, 0x02, 0xab, 0xcd,
];
const PNG_B: &[u8] = &[
    0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 0xde, 0xad, 0xbe, 0xef,
];
const PNG_C: &[u8] = &[
    0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a, 0x11, 0x22, 0xcd, 0xef,
];

fn two_picture_source() -> &'static str {
    concat!(
        "{\\rtf1\\ansi Before{\\pict\\pngblip\\picw1\\pich1 89504E47 0d0A1a0A\r\n 0102aBcd}",
        r"Middle{\pict\pngblip\picw2\pich2 89504e470D0A1A0A DEADBEEF}After}",
    )
}

fn durable_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        1024 * 1024,
        max_operations,
        8,
        256 * 1024,
        512 * 1024,
    )
}

fn encode_hex(bytes: &[u8]) -> String {
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        use std::fmt::Write as _;
        write!(&mut output, "{byte:02x}").unwrap();
    }
    output
}

fn decode_hex(input: &str) -> Vec<u8> {
    input
        .as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}

#[test]
fn selected_payload_splice_preserves_hex_layout_metadata_and_other_bytes() {
    let source = Document::parse(two_picture_source()).unwrap();
    assert_eq!(source.pictures()[0].data(), PNG_A);
    assert_eq!(source.pictures()[1].data(), PNG_B);
    let before_second = source.pictures()[1].clone();

    let mut edit = source.edit();
    edit.replace_picture_payload(0, PNG_C).unwrap();
    let commit = edit.commit().unwrap();
    let output = commit.snapshot().to_bytes().unwrap();
    let expected = two_picture_source().replace("0102aBcd", "1122cDef");

    assert_eq!(output, expected.as_bytes());
    assert_eq!(commit.snapshot().pictures()[0].data(), PNG_C);
    assert_eq!(commit.snapshot().pictures()[1], before_second);
    assert_eq!(commit.snapshot().text(), source.text());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().operation_count(), 1);
}

#[test]
fn no_op_shares_snapshot_and_patch_inverse_is_exact_and_source_checked() {
    let source = Document::parse(two_picture_source()).unwrap();
    let mut no_op = source.edit();
    no_op.replace_picture_payload(0, PNG_A).unwrap();
    let no_op = no_op.commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.snapshot().same_snapshot(&source));

    let mut edit = source.edit();
    edit.replace_picture_payload(0, PNG_C).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    let stale = Document::parse(r"{\rtf1 unrelated}").unwrap();
    assert!(matches!(
        commit.patch().apply(&stale),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn batch_is_atomic_bounded_and_updates_only_selected_pictures() {
    let source = Document::parse(two_picture_source()).unwrap();
    let replacements = [
        PicturePayloadReplacement::new(0, PNG_C),
        PicturePayloadReplacement::new(1, PNG_A),
    ];
    let mut edit = source.edit();
    edit.replace_picture_payloads(&replacements).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().pictures()[0].data(), PNG_C);
    assert_eq!(commit.snapshot().pictures()[1].data(), PNG_A);
    let durable = commit.patch().to_durable(durable_limits(2)).unwrap();
    let applied = source.apply_durable(&durable).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );
    assert_eq!(
        applied
            .apply_durable(&durable.inverse())
            .unwrap()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );

    let mut mixed = source.edit();
    mixed.replace_picture_payload(0, PNG_C).unwrap();
    assert!(matches!(
        mixed.replace_body_text("not composable"),
        Err(Error::BodyDestinationConflict)
    ));

    let invalid = [
        PicturePayloadReplacement::new(0, PNG_C),
        PicturePayloadReplacement::new(2, PNG_A),
    ];
    let mut atomic = source.edit();
    assert!(matches!(
        atomic.replace_picture_payloads(&invalid),
        Err(Error::PictureOutOfRange {
            position: 2,
            count: 2
        })
    ));
    assert_eq!(atomic.operation_count(), 0);

    let duplicate = [
        PicturePayloadReplacement::new(1, PNG_A),
        PicturePayloadReplacement::new(1, PNG_C),
    ];
    assert!(matches!(
        source.edit().replace_picture_payloads(&duplicate),
        Err(Error::PicturePayloadBatchOutOfOrder {
            previous: 1,
            incoming: 1
        })
    ));

    let mut bounded = source.edit_with_limits(Limits::new(1));
    assert!(matches!(
        bounded.replace_picture_payloads(&replacements),
        Err(Error::OperationLimit {
            observed: 2,
            limit: 1
        })
    ));
    assert_eq!(bounded.operation_count(), 0);
}

#[test]
fn refuses_size_format_nested_opaque_and_protected_sources() {
    let source = Document::parse(two_picture_source()).unwrap();
    let mut wrong_size = source.edit();
    assert!(matches!(
        wrong_size.replace_picture_payload(0, &PNG_C[..PNG_C.len() - 1]),
        Err(Error::PicturePayloadSizeMismatch { position: 0, .. })
    ));
    assert_eq!(wrong_size.operation_count(), 0);

    let mut wrong_format = PNG_C.to_vec();
    wrong_format[0] = 0;
    assert!(matches!(
        source.edit().replace_picture_payload(0, wrong_format),
        Err(Error::UnsupportedSource(_))
    ));

    let nested = Document::parse(concat!(
        r"{\rtf1{\*\shppict{\pict\pngblip ",
        "89504e470d0a1a0a0102abcd",
        r"}}}",
    ))
    .unwrap();
    assert!(matches!(
        nested.edit().replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));

    let opaque = Document::parse(concat!(
        r"{\rtf1\future42{\pict\pngblip ",
        "89504e470d0a1a0a0102abcd",
        r"}}",
    ))
    .unwrap();
    assert!(matches!(
        opaque.edit().replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));

    let protected = Document::parse(concat!(
        r"{\rtf1\allprot\enforceprot1{\pict\pngblip ",
        "89504e470d0a1a0a0102abcd",
        r"}}",
    ))
    .unwrap();
    let mut edit = protected.edit();
    edit.replace_picture_payload(0, PNG_C).unwrap();
    assert!(matches!(
        edit.commit(),
        Err(Error::ProtectedDocument { .. })
    ));

    assert!(Document::parse(r"{\rtf1{\pict\pngblip 89504e470d0a1a0a0}}").is_err());

    let late_control =
        Document::parse(r"{\rtf1{\pict\pngblip 89504e470d0a1a0a0102abcd\picw1}}").unwrap();
    assert!(matches!(
        late_control.edit().replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));

    let unknown_picture_control =
        Document::parse(r"{\rtf1{\pict\pngblip\vendorpicture1 89504e470d0a1a0a0102abcd}}").unwrap();
    assert!(matches!(
        unknown_picture_control
            .edit()
            .replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));

    for ambiguous in [
        r"{\rtf1{\pict\pngblip\pngblip 89504e470d0a1a0a0102abcd}}",
        r"{\rtf1{\pict\pngblip\jpegblip 89504e470d0a1a0a0102abcd}}",
        r"{\rtf1{\pict\pngblip\picw1\picw2 89504e470d0a1a0a0102abcd}}",
    ] {
        let ambiguous = Document::parse(ambiguous).unwrap();
        assert!(matches!(
            ambiguous.edit().replace_picture_payload(0, PNG_C),
            Err(Error::UnsupportedSource(_))
        ));
    }

    let external_picture =
        Document::parse(r"{\rtf1{\pict\pngblip\blipfile1 89504e470d0a1a0a0102abcd}}").unwrap();
    assert!(matches!(
        external_picture.edit().replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));
}

#[test]
fn refuses_binary_compressed_field_object_and_shape_ownership() {
    let mut binary = br"{\rtf1{\pict\jpegblip\bin6 ".to_vec();
    binary.extend_from_slice(&[0xff, 0xd8, 0x01, 0x02, 0xff, 0xd9]);
    binary.extend_from_slice(br"}}");
    let binary = Document::from_bytes(&binary).unwrap();
    assert!(matches!(
        binary
            .edit()
            .replace_picture_payload(0, [0xff, 0xd8, 3, 4, 0xff, 0xd9]),
        Err(Error::UnsupportedSource(_))
    ));

    let compressed_bytes =
        litchi_rtf::transport::compress(two_picture_source().as_bytes(), true).unwrap();
    let compressed = Document::from_bytes(&compressed_bytes).unwrap();
    assert!(matches!(
        compressed.edit().replace_picture_payload(0, PNG_C),
        Err(Error::UnsupportedSource(_))
    ));

    let sources = [
        concat!(
            r"{\rtf1{\pict\pngblip 89504e470d0a1a0a0102abcd}",
            r#"{\field{\*\fldinst HYPERLINK "https://example.test"}{\fldrslt Link}}}"#,
        ),
        concat!(
            r"{\rtf1{\pict\pngblip 89504e470d0a1a0a0102abcd}",
            r"{\object\objemb{\*\objclass Package}{\*\objdata 0102}{\result fallback}}}",
        ),
        concat!(
            r"{\rtf1{\pict\pngblip 89504e470d0a1a0a0102abcd}",
            r"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}}",
        ),
    ];
    for source in sources {
        let source = Document::parse(source).unwrap();
        assert!(matches!(
            source.edit().replace_picture_payload(0, PNG_C),
            Err(Error::UnsupportedSource(_))
        ));
    }
}

#[test]
fn jpeg_payload_retains_markers_and_exact_layout() {
    let source =
        Document::parse(r"{\rtf1 A{\pict\jpegblip\picw9\pich8 Ff D8 01aB fF d9}B}").unwrap();
    let replacement = [0xff, 0xd8, 0xaa, 0x55, 0xff, 0xd9];
    let mut edit = source.edit();
    edit.replace_picture_payload(0, replacement).unwrap();
    let output = edit.commit().unwrap().snapshot().to_bytes().unwrap();
    assert_eq!(
        output,
        br"{\rtf1 A{\pict\jpegblip\picw9\pich8 Ff D8 aa55 fF d9}B}"
    );
    let reopened = Document::from_bytes(&output).unwrap();
    assert_eq!(reopened.pictures()[0].data(), replacement);
}

#[test]
fn picture_specific_batch_ceiling_accepts_64_and_refuses_65_atomically() {
    let mut source = String::from(r"{\rtf1 ");
    for _ in 0..=MAX_PICTURE_PAYLOAD_OPERATIONS {
        source.push_str(r"{\pict\pngblip 89504e470d0a1a0a0102abcd}");
    }
    source.push('}');
    let source = Document::parse(&source).unwrap();
    let accepted = (0..MAX_PICTURE_PAYLOAD_OPERATIONS)
        .map(|position| PicturePayloadReplacement::new(position, PNG_C))
        .collect::<Vec<_>>();
    let mut edit = source.edit();
    edit.replace_picture_payloads(&accepted).unwrap();
    assert_eq!(edit.operation_count(), MAX_PICTURE_PAYLOAD_OPERATIONS);
    assert!(edit.commit().unwrap().diagnostics().changed());

    let refused = (0..=MAX_PICTURE_PAYLOAD_OPERATIONS)
        .map(|position| PicturePayloadReplacement::new(position, PNG_C))
        .collect::<Vec<_>>();
    let mut edit = source.edit();
    assert!(matches!(
        edit.replace_picture_payloads(&refused),
        Err(Error::OperationLimit {
            observed: 65,
            limit: 64
        })
    ));
    assert_eq!(edit.operation_count(), 0);
}

#[test]
fn composition_cannot_bypass_the_picture_specific_batch_ceiling() {
    let mut source = String::from(r"{\rtf1 ");
    for _ in 0..=MAX_PICTURE_PAYLOAD_OPERATIONS {
        source.push_str(r"{\pict\pngblip 89504e470d0a1a0a0102abcd}");
    }
    source.push('}');
    let source = Document::parse(&source).unwrap();
    let limits = CompositionLimits::new(65, 2, 130, 4);
    let mut composition = Composition::new(&source, limits);
    for position in 0..MAX_PICTURE_PAYLOAD_OPERATIONS {
        let mut edit = source.edit();
        edit.replace_picture_payload(position, PNG_C).unwrap();
        composition
            .join(
                edit.into_sub_edit(format!("picture-{position}"), limits)
                    .unwrap(),
            )
            .unwrap();
    }
    assert_eq!(
        composition
            .commit()
            .unwrap()
            .diagnostics()
            .operation_count(),
        MAX_PICTURE_PAYLOAD_OPERATIONS
    );

    let mut composition = Composition::new(&source, limits);
    for position in 0..=MAX_PICTURE_PAYLOAD_OPERATIONS {
        let mut edit = source.edit();
        edit.replace_picture_payload(position, PNG_C).unwrap();
        composition
            .join(
                edit.into_sub_edit(format!("picture-{position}"), limits)
                    .unwrap(),
            )
            .unwrap();
    }
    assert!(matches!(
        composition.commit(),
        Err(Error::OperationLimit {
            observed: 65,
            limit: 64
        })
    ));
}

#[test]
fn durable_patch_is_deterministic_exactly_reversible_and_stale_checked() {
    let source = Document::parse(two_picture_source()).unwrap();
    let mut edit = source.edit();
    edit.replace_picture_payload(0, PNG_C).unwrap();
    let commit = edit.commit().unwrap();
    let limits = durable_limits(1);
    let durable = commit.patch().to_durable(limits).unwrap();
    let first = durable.to_deterministic_json().unwrap();
    let second = durable.to_deterministic_json().unwrap();
    assert_eq!(first, second);

    let decoded =
        litchi_core::patch::Patch::<litchi_core::patch::Reversible>::from_deterministic_json(
            &first, limits,
        )
        .unwrap();
    let applied = source.apply_durable(&decoded).unwrap();
    assert_eq!(
        applied.to_bytes().unwrap(),
        commit.snapshot().to_bytes().unwrap()
    );
    let restored = applied.apply_durable(&decoded.inverse()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), source.to_bytes().unwrap());

    let stale = Document::parse(
        two_picture_source()
            .replace("0102aBcd", "0102abce")
            .as_str(),
    )
    .unwrap();
    assert!(matches!(
        stale.apply_durable(&decoded),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn durable_patch_rejects_transport_layout_changes_and_unknown_fields() {
    let source = Document::parse(two_picture_source()).unwrap();
    let mut edit = source.edit();
    edit.replace_picture_payload(0, PNG_C).unwrap();
    let commit = edit.commit().unwrap();
    let limits = durable_limits(1);
    let durable = commit.patch().to_durable(limits).unwrap();
    let inverse_operation = durable.inverse().operations()[0].clone();

    let mut changed_layout = durable.operations()[0].clone();
    let value = changed_layout.value.as_object_mut().unwrap();
    let transport = value.get("transport").and_then(Value::as_str).unwrap();
    let mut transport = decode_hex(transport);
    let whitespace = transport
        .iter()
        .position(|byte| byte.is_ascii_whitespace())
        .unwrap();
    let digit = transport
        .iter()
        .position(|byte| byte.is_ascii_hexdigit())
        .unwrap();
    transport.swap(whitespace, digit);
    value.insert(
        "transport".to_string(),
        Value::String(encode_hex(&transport)),
    );
    let malicious = litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [litchi_core::patch::ReversibleOperation::new(
            changed_layout,
            inverse_operation.clone(),
        )],
        litchi_core::patch::BlobBundle::new(limits.blobs()),
        litchi_core::patch::BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&malicious),
        Err(Error::DurablePatch(_))
    ));

    let mut unknown_field = durable.operations()[0].clone();
    unknown_field
        .value
        .as_object_mut()
        .unwrap()
        .insert("unknown".to_string(), Value::Bool(true));
    let malicious = litchi_core::patch::Patch::<litchi_core::patch::Reversible>::new(
        limits,
        "litchi-rtf",
        [litchi_core::patch::ReversibleOperation::new(
            unknown_field,
            inverse_operation,
        )],
        litchi_core::patch::BlobBundle::new(limits.blobs()),
        litchi_core::patch::BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_durable(&malicious),
        Err(Error::DurablePatch(_))
    ));
}
