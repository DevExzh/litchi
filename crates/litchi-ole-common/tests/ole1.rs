use litchi_cfb::OleError;
use litchi_ole_common::ole1::{
    CF_DIB, Compatibility, EncodedString, FORMAT_ID_EMBEDDED, Limits, ObjectHeader, ObjectKind,
    ObjectRef, ObjectSnapshot, PRESENTATION_FORMAT_ID_CLASS, Presentation, PresentationKind,
    PresentationRef, RegisteredForm, TextEncoding,
};
use std::sync::Arc;

fn ansi(value: &[u8]) -> EncodedString {
    let payload = if value.is_empty() {
        Vec::new()
    } else {
        let mut payload = value.to_vec();
        payload.push(0);
        payload
    };
    EncodedString::ansi(payload).expect("valid ANSI test string")
}

fn unicode_registered(value: u8) -> EncodedString {
    let mut payload = Vec::new();
    for byte in b"OleExternal".iter().copied().chain([value]) {
        payload.extend_from_slice(&[byte, 0]);
    }
    payload.extend_from_slice(&[0, 0]);
    EncodedString::unicode(payload).expect("valid UTF-16LE registered name")
}

fn embedded_header() -> ObjectHeader {
    ObjectHeader::embedded(ansi(b"Test.Class"), ansi(b""), ansi(b""))
        .expect("valid embedded header")
}

fn linked_header() -> ObjectHeader {
    ObjectHeader::linked(ansi(b"Test.Class"), ansi(b"C:\\source.bin"), ansi(b"item"))
        .expect("valid linked header")
}

fn embedded(presentation: Presentation) -> ObjectSnapshot {
    ObjectSnapshot::new_embedded(embedded_header(), b"native bytes".to_vec(), presentation)
        .expect("valid embedded object")
}

fn push_u32(bytes: &mut Vec<u8>, value: u32) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn generic_registered_prefix(bytes: &mut Vec<u8>) {
    push_u32(bytes, 0x501);
    push_u32(bytes, PRESENTATION_FORMAT_ID_CLASS);
    push_u32(bytes, 7);
    bytes.extend_from_slice(b"Custom\0");
    push_u32(bytes, 0);
}

fn object_prefix(bytes: &mut Vec<u8>) {
    push_u32(bytes, 0x501);
    push_u32(bytes, FORMAT_ID_EMBEDDED);
    push_u32(bytes, 0);
    push_u32(bytes, 0);
    push_u32(bytes, 0);
    push_u32(bytes, 0);
}

#[test]
fn canonical_authoring_covers_all_five_presentation_forms() {
    let values = [
        Presentation::metafile(-12, 34, b"wmf".to_vec()).unwrap(),
        Presentation::bitmap(-56, 78, b"bmp".to_vec()).unwrap(),
        Presentation::dib(90, -12, b"dib".to_vec()).unwrap(),
        Presentation::standard_clipboard(ansi(b"Custom.Clipboard"), CF_DIB, b"standard".to_vec())
            .unwrap(),
        Presentation::registered_clipboard(
            ansi(b"Custom.Registered"),
            unicode_registered(b'R'),
            b"registered".to_vec(),
        )
        .unwrap(),
    ];

    let expected = [
        PresentationKind::MetaFile,
        PresentationKind::Bitmap,
        PresentationKind::Dib,
        PresentationKind::StandardClipboard,
        PresentationKind::RegisteredClipboard,
    ];

    for (presentation, expected_kind) in values.into_iter().zip(expected) {
        let snapshot = embedded(presentation);
        let borrowed = ObjectRef::parse(snapshot.bytes()).unwrap();
        assert_eq!(borrowed.bytes(), snapshot.bytes());
        let parsed = ObjectSnapshot::parse(snapshot.bytes()).unwrap();
        assert_eq!(parsed.presentation().kind(), expected_kind);
        assert_eq!(
            PresentationRef::parse(parsed.presentation().bytes())
                .unwrap()
                .kind(),
            expected_kind
        );
        assert_eq!(parsed.bytes(), snapshot.bytes());
        assert!(parsed.rewrite_safe());
    }
}

#[test]
fn source_noop_shares_arc_and_patch_is_exactly_reversible() {
    let snapshot = embedded(Presentation::metafile(1, 2, b"data".to_vec()).unwrap());
    let source_arc = snapshot.bytes_shared();
    let mut edit = snapshot.edit();
    assert!(!edit.set_width(1).unwrap());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert!(Arc::ptr_eq(&source_arc, &commit.snapshot().bytes_shared()));

    let mut changed = snapshot.edit();
    assert!(changed.set_width(-77).unwrap());
    let commit = changed.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.snapshot().presentation().width(), Some(-77));
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.bytes(), snapshot.bytes());
    assert_eq!(
        commit.patch().apply(&snapshot).unwrap().bytes(),
        commit.snapshot().bytes()
    );

    let source_arc = snapshot.bytes_shared();
    let mut reverted = snapshot.edit();
    assert!(reverted.set_width(-13).unwrap());
    assert!(reverted.set_width(1).unwrap());
    let reverted_commit = reverted.commit().unwrap();
    assert!(!reverted_commit.changed());
    assert!(Arc::ptr_eq(
        &source_arc,
        &reverted_commit.snapshot().bytes_shared()
    ));
    let applied_noop = reverted_commit.patch().apply(&snapshot).unwrap();
    assert!(Arc::ptr_eq(&source_arc, &applied_noop.bytes_shared()));
}

#[test]
fn edits_preserve_signed_dimensions_and_reserved_metafile_words() {
    let mut bytes = embedded(Presentation::metafile(-4, 5, b"wmf".to_vec()).unwrap())
        .bytes()
        .to_vec();
    let presentation_start = 8 + (4 + b"Test.Class\0".len()) + 4 + 4 + 4 + b"native bytes".len();
    let reserved_start = presentation_start + 8 + (4 + b"METAFILEPICT\0".len()) + 4 + 4 + 4;
    bytes[reserved_start..reserved_start + 8].copy_from_slice(&[1, 0, 2, 0, 3, 0, 4, 0]);
    let snapshot = ObjectSnapshot::parse(&bytes).unwrap();
    assert_eq!(
        snapshot.presentation().metafile_reserved(),
        Some([1, 2, 3, 4])
    );

    let mut edit = snapshot.edit();
    assert!(edit.set_height(-99).unwrap());
    let committed = edit.commit().unwrap();
    assert_eq!(committed.snapshot().presentation().height(), Some(-99));
    assert_eq!(
        committed.snapshot().presentation().metafile_reserved(),
        Some([1, 2, 3, 4])
    );
    assert_eq!(
        &committed.snapshot().bytes()[reserved_start..reserved_start + 8],
        &[1, 0, 2, 0, 3, 0, 4, 0]
    );
}

#[test]
fn linked_source_keeps_network_encoding_and_updates_only_requested_field() {
    let snapshot = ObjectSnapshot::new_linked(
        linked_header(),
        ansi(b"\\\\server\\share\\source.bin"),
        7,
        Presentation::bitmap(10, 20, b"bitmap".to_vec()).unwrap(),
    )
    .unwrap();
    assert_eq!(snapshot.kind(), ObjectKind::Linked);
    assert_eq!(
        snapshot.network_name().unwrap().encoding(),
        TextEncoding::Ansi
    );
    assert_eq!(snapshot.link_update_option(), Some(7));

    let mut edit = snapshot.edit();
    assert!(edit.set_link_update_option(9).unwrap());
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().link_update_option(), Some(9));
    assert_eq!(
        commit.snapshot().network_name().unwrap().payload(),
        snapshot.network_name().unwrap().payload()
    );
}

#[test]
fn registered_optional_name_is_rejected_when_two_complete_grammars_match() {
    let mut bytes = Vec::new();
    object_prefix(&mut bytes);
    generic_registered_prefix(&mut bytes);
    // An empty StringFormatData is valid under both the ANSI and Unicode
    // length-prefixed grammars, so the exact enclosing bytes are ambiguous.
    push_u32(&mut bytes, 4);
    bytes.extend_from_slice(&[0, 0, 0, 0]);
    push_u32(&mut bytes, 1);
    bytes.push(b'Z');

    let snapshot = ObjectSnapshot::parse(&bytes).unwrap();
    let presentation = snapshot.presentation();
    assert_eq!(presentation.kind(), PresentationKind::RegisteredClipboard);
    assert_eq!(
        presentation.registered_form(),
        Some(RegisteredForm::Ambiguous)
    );
    assert!(!snapshot.rewrite_safe());
    assert_eq!(presentation.data(), None);
    let mut edit = snapshot.edit();
    assert!(edit.set_presentation_data(vec![b'Y']).is_err());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), bytes.as_slice());
}

#[test]
fn unicode_registered_name_retains_exact_utf16le_payload() {
    let snapshot = embedded(
        Presentation::registered_clipboard(
            ansi(b"Custom.Registered"),
            unicode_registered(b'Q'),
            b"payload".to_vec(),
        )
        .unwrap(),
    );
    let name = snapshot.presentation().registered_format_name().unwrap();
    assert_eq!(name.encoding(), TextEncoding::Unicode);
    assert!(
        name.payload()
            .starts_with(b"O\0l\0e\0E\0x\0t\0e\0r\0n\0a\0l\0")
    );
    assert_eq!(
        snapshot.presentation().registered_form(),
        Some(RegisteredForm::WithName)
    );

    let encoded_name = name.encoded_bytes().to_vec();
    let mut edit = snapshot.edit();
    assert!(edit.set_presentation_data(b"changed".to_vec()).unwrap());
    let committed = edit.commit().unwrap();
    assert_eq!(
        committed
            .snapshot()
            .presentation()
            .registered_format_name()
            .unwrap()
            .encoded_bytes(),
        encoded_name.as_slice()
    );
}

#[test]
fn registered_without_name_is_distinguished_and_source_only() {
    let mut bytes = Vec::new();
    object_prefix(&mut bytes);
    generic_registered_prefix(&mut bytes);
    push_u32(&mut bytes, 0);

    let snapshot = ObjectSnapshot::parse(&bytes).unwrap();
    let presentation = snapshot.presentation();
    assert_eq!(
        presentation.registered_form(),
        Some(RegisteredForm::WithoutName)
    );
    assert_eq!(presentation.data(), Some(&[][..]));
    assert!(!snapshot.rewrite_safe());
    assert!(snapshot.edit().set_presentation_data(vec![b'X']).is_err());
    assert_eq!(snapshot.bytes(), bytes.as_slice());
}

#[test]
fn malformed_lp_values_and_invalid_limits_are_rejected_before_editing() {
    assert!(EncodedString::ansi(vec![0, 1]).is_err());
    assert!(EncodedString::unicode(vec![b'A', 0]).is_err());
    assert!(
        ObjectSnapshot::parse_with_limits(
            &[0u8; 1],
            Limits {
                max_bytes: 0,
                ..Limits::default()
            }
        )
        .is_err()
    );

    let mut bytes = embedded(Presentation::dib(1, 2, b"data".to_vec()).unwrap())
        .bytes()
        .to_vec();
    bytes[8..12].copy_from_slice(&u32::MAX.to_le_bytes());
    assert!(ObjectSnapshot::parse(&bytes).is_err());
}

#[test]
fn fresh_authoring_enforces_registered_clipboard_and_link_path_rules() {
    assert!(Presentation::standard_clipboard(ansi(b"Custom"), 0x1234, b"data".to_vec()).is_err());
    assert!(
        ObjectSnapshot::new_embedded(
            embedded_header(),
            b"native".to_vec(),
            Presentation::StandardClipboard {
                class_name: ansi(b"Custom"),
                clipboard_format: 0x1234,
                data: b"data".to_vec(),
            },
        )
        .is_err()
    );
    assert!(
        Presentation::registered_clipboard(
            ansi(b"Custom"),
            EncodedString::ansi(vec![0]).unwrap(),
            b"data".to_vec()
        )
        .is_err()
    );
    assert!(
        Presentation::registered_clipboard(
            ansi(b"Custom"),
            EncodedString::unicode(vec![0, 0]).unwrap(),
            b"data".to_vec()
        )
        .is_err()
    );
    assert!(
        Presentation::registered_clipboard(
            ansi(b"Custom"),
            ansi(b"Vendor.Format"),
            b"data".to_vec()
        )
        .is_err()
    );
    assert!(ObjectHeader::embedded(ansi(b""), ansi(b""), ansi(b"")).is_err());
    assert!(
        ObjectHeader::embedded(EncodedString::ansi(vec![0]).unwrap(), ansi(b""), ansi(b""))
            .is_err()
    );
    assert!(ObjectHeader::linked(ansi(b"Class"), ansi(b"relative.bin"), ansi(b"item")).is_err());
    assert!(ObjectHeader::linked(ansi(b"Class"), ansi(b"C:relative.bin"), ansi(b"item")).is_err());
    assert!(ObjectHeader::linked(ansi(b"Class"), ansi(b"\\\\server"), ansi(b"item")).is_err());
    assert!(ObjectHeader::linked(ansi(b"Class"), ansi(b"\\\\server\\\\"), ansi(b"item")).is_err());
    assert!(
        ObjectHeader::linked(
            ansi(b"Class"),
            ansi(b"\\\\server\\share\\item"),
            ansi(b"item")
        )
        .is_ok()
    );

    let source = ObjectSnapshot::new_linked(
        linked_header(),
        ansi(b"\\\\server\\share\\source.bin"),
        0,
        Presentation::dib(1, 2, b"data".to_vec()).unwrap(),
    )
    .unwrap();
    let mut edit = source.edit();
    assert!(edit.set_topic_name(ansi(b"\\\\server\\\\")).is_err());
    assert_eq!(edit.snapshot().bytes(), source.bytes());
}

#[test]
fn fresh_authoring_preflights_all_strings_at_their_caller_bound() {
    let header = ObjectHeader::embedded(ansi(b"Class"), ansi(b"Topic"), ansi(b"Item")).unwrap();
    let presentation = Presentation::dib(1, 2, b"data".to_vec()).unwrap();
    let exact = Limits {
        max_bytes: 128,
        max_native_bytes: 6,
        max_presentation_bytes: 64,
        max_string_bytes: 6,
        max_registered_format_bytes: 6,
    };
    assert!(
        ObjectSnapshot::new_embedded_with_limits(
            header.clone(),
            b"native".to_vec(),
            presentation.clone(),
            exact,
        )
        .is_ok()
    );

    // The authored object is larger than max_bytes, but the string ceiling is
    // reported first because every string is admitted before output sizing or
    // reservation.
    let string_error = ObjectSnapshot::new_embedded_with_limits(
        header,
        b"native".to_vec(),
        presentation,
        Limits {
            max_bytes: 64,
            max_string_bytes: 5,
            max_registered_format_bytes: 5,
            ..exact
        },
    )
    .expect_err("class/topic strings exceed the caller bound");
    assert!(matches!(
        string_error,
        OleError::LimitExceeded {
            resource: "OLE1 string bytes",
            ..
        }
    ));

    let generic =
        Presentation::standard_clipboard(ansi(b"LongClass"), CF_DIB, b"data".to_vec()).unwrap();
    let generic_error = ObjectSnapshot::new_embedded_with_limits(
        ObjectHeader::embedded(ansi(b"C"), ansi(b""), ansi(b"")).unwrap(),
        b"native".to_vec(),
        generic,
        Limits {
            max_string_bytes: 3,
            max_registered_format_bytes: 3,
            ..exact
        },
    )
    .expect_err("canonical and generic presentation class names are bounded");
    assert!(matches!(
        generic_error,
        OleError::LimitExceeded {
            resource: "OLE1 string bytes",
            ..
        }
    ));

    let network = ansi(b"\\\\server\\share\\source.bin");
    let linked_error = ObjectSnapshot::new_linked_with_limits(
        linked_header(),
        network.clone(),
        0,
        Presentation::dib(1, 2, b"data".to_vec()).unwrap(),
        Limits {
            max_string_bytes: network.payload().len() - 1,
            max_registered_format_bytes: network.payload().len() - 1,
            ..exact
        },
    )
    .expect_err("network name is bounded before linked output reservation");
    assert!(matches!(
        linked_error,
        OleError::LimitExceeded {
            resource: "OLE1 string bytes",
            ..
        }
    ));
}

#[test]
fn source_reader_preserves_noncanonical_clipboard_and_topic_forms_without_typed_rewrite() {
    let standard_source = embedded(
        Presentation::standard_clipboard(ansi(b"Custom"), CF_DIB, b"data".to_vec()).unwrap(),
    );
    let mut bytes = standard_source.bytes().to_vec();
    let presentation = standard_source.presentation().bytes();
    let presentation_offset = bytes.len() - presentation.len();
    let clipboard_offset = presentation_offset + 8 + (4 + b"Custom\0".len());
    bytes[clipboard_offset..clipboard_offset + 4].copy_from_slice(&0x1234u32.to_le_bytes());
    let snapshot = ObjectSnapshot::parse(&bytes).unwrap();
    assert_eq!(snapshot.presentation().clipboard_format(), Some(0x1234));
    assert!(!snapshot.rewrite_safe());
    assert!(
        snapshot
            .edit()
            .set_presentation_data(b"new".to_vec())
            .is_err()
    );
    assert_eq!(snapshot.bytes(), bytes.as_slice());

    let mut linked = Vec::new();
    push_u32(&mut linked, 0x501);
    push_u32(&mut linked, 1);
    for field in [
        b"Class\0".as_slice(),
        b"relative.bin\0".as_slice(),
        b"item\0".as_slice(),
    ] {
        push_u32(&mut linked, field.len() as u32);
        linked.extend_from_slice(field);
    }
    push_u32(&mut linked, 0);
    push_u32(&mut linked, 0);
    push_u32(&mut linked, 0);
    linked.extend_from_slice(&snapshot.bytes()[presentation_offset..]);
    let linked_snapshot = ObjectSnapshot::parse(&linked).unwrap();
    assert!(!linked_snapshot.rewrite_safe());
    assert_eq!(linked_snapshot.bytes(), linked.as_slice());
}

#[test]
fn nonzero_link_reserved_is_preserved_for_noop_but_refuses_typed_rewrites() {
    let source = ObjectSnapshot::new_linked(
        ObjectHeader::linked(ansi(b"Test.Class"), ansi(b"C:\\source.bin"), ansi(b"")).unwrap(),
        EncodedString::ansi(Vec::new()).unwrap(),
        3,
        Presentation::dib(1, 2, b"data".to_vec()).unwrap(),
    )
    .unwrap();
    let mut bytes = source.bytes().to_vec();
    // Reserved follows the variable-length ObjectHeader and NetworkName
    // fields.  Locate it from the exact authoring spellings used above.
    let reserved_offset = 8 + (4 + b"Test.Class\0".len()) + (4 + b"C:\\source.bin\0".len()) + 4 + 4;
    bytes[reserved_offset..reserved_offset + 4].copy_from_slice(&1u32.to_le_bytes());
    let snapshot = ObjectSnapshot::parse(&bytes).unwrap();
    assert!(!snapshot.rewrite_safe());
    let mut edit = snapshot.edit();
    assert!(edit.set_link_update_option(4).is_err());
    let commit = edit.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), bytes.as_slice());
}

#[test]
fn nested_format_zero_requires_explicit_empty_header_compatibility_and_is_source_only() {
    let mut bytes = Vec::new();
    push_u32(&mut bytes, 0x501);
    push_u32(&mut bytes, FORMAT_ID_EMBEDDED);
    bytes.extend_from_slice(&[0, 0, 0, 0]);
    bytes.extend_from_slice(&[0, 0, 0, 0]);
    bytes.extend_from_slice(&[0, 0, 0, 0]);
    push_u32(&mut bytes, 10);
    bytes.extend_from_slice(b"opaque CFB bytes");
    let native_size = 16u32;
    bytes[20..24].copy_from_slice(&native_size.to_le_bytes());
    push_u32(&mut bytes, 0x501);
    push_u32(&mut bytes, 0);

    assert!(ObjectSnapshot::parse(&bytes).is_err());
    let snapshot = ObjectSnapshot::parse_with_limits_and_compatibility(
        &bytes,
        Limits::default(),
        Compatibility::EmptyPresentationHeader,
    )
    .unwrap();
    assert_eq!(
        snapshot.presentation().kind(),
        PresentationKind::RawFormatZero
    );
    assert!(!snapshot.rewrite_safe());
    assert!(snapshot.edit().snapshot().bytes() == bytes.as_slice());

    let mut trailing = bytes.clone();
    trailing.push(0);
    assert!(
        ObjectSnapshot::parse_with_limits_and_compatibility(
            &trailing,
            Limits::default(),
            Compatibility::EmptyPresentationHeader,
        )
        .is_err()
    );
}

#[test]
fn missing_embedded_presentation_requires_explicit_compatibility() {
    let mut bytes = Vec::new();
    object_prefix(&mut bytes);
    let native = b"native Draw document";
    bytes[20..24].copy_from_slice(&(native.len() as u32).to_le_bytes());
    bytes.extend_from_slice(native);

    assert!(ObjectSnapshot::parse(&bytes).is_err());
    let snapshot = ObjectSnapshot::parse_with_limits_and_compatibility(
        &bytes,
        Limits::default(),
        Compatibility::MissingPresentation,
    )
    .unwrap();
    assert_eq!(snapshot.native_data(), Some(native.as_slice()));
    assert_eq!(
        snapshot.presentation().kind(),
        PresentationKind::MissingPresentation
    );
    assert!(!snapshot.rewrite_safe());
    assert!(snapshot.edit().set_presentation_data(Vec::new()).is_err());
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), bytes.as_slice());
}

#[test]
fn limits_are_checked_before_admission_and_stale_patch_is_rejected() {
    let snapshot = embedded(Presentation::dib(1, 2, b"payload".to_vec()).unwrap());
    let exact = Limits {
        max_bytes: snapshot.bytes().len(),
        max_native_bytes: b"native bytes".len(),
        max_presentation_bytes: snapshot.presentation().bytes().len(),
        max_string_bytes: 64,
        max_registered_format_bytes: 64,
    };
    assert!(ObjectSnapshot::parse_with_limits(snapshot.bytes(), exact).is_ok());
    let too_small = Limits {
        max_bytes: snapshot.bytes().len() - 1,
        ..exact
    };
    assert!(matches!(
        ObjectSnapshot::parse_with_limits(snapshot.bytes(), too_small),
        Err(OleError::LimitExceeded { .. })
    ));
    assert!(matches!(
        ObjectSnapshot::parse_with_limits(
            snapshot.bytes(),
            Limits {
                max_native_bytes: b"native bytes".len() - 1,
                ..exact
            }
        ),
        Err(OleError::LimitExceeded {
            resource: "OLE1 native bytes",
            ..
        })
    ));
    assert!(matches!(
        ObjectSnapshot::parse_with_limits(
            snapshot.bytes(),
            Limits {
                max_presentation_bytes: snapshot.presentation().bytes().len() - 1,
                ..exact
            }
        ),
        Err(OleError::LimitExceeded {
            resource: "OLE1 presentation bytes",
            ..
        })
    ));
    assert!(matches!(
        ObjectSnapshot::parse_with_limits(
            snapshot.bytes(),
            Limits {
                max_string_bytes: 1,
                max_registered_format_bytes: 1,
                ..exact
            }
        ),
        Err(OleError::LimitExceeded {
            resource: "OLE1 string bytes",
            ..
        })
    ));

    let bounded = ObjectSnapshot::parse_with_limits(snapshot.bytes(), exact).unwrap();
    let before = bounded.bytes().to_vec();
    let mut bounded_edit = bounded.edit();
    assert!(
        bounded_edit
            .set_presentation_data(vec![1, 2, 3, 4, 5, 6, 7, 8])
            .is_err()
    );
    assert_eq!(bounded_edit.snapshot().bytes(), before.as_slice());

    let mut edit = snapshot.edit();
    edit.set_width(3).unwrap();
    let commit = edit.commit().unwrap();
    let mut other_edit = snapshot.edit();
    other_edit.set_width(4).unwrap();
    let other = other_edit.commit().unwrap().into_snapshot();
    assert!(commit.patch().apply(&other).is_err());
}

#[test]
#[ignore = "requires the externally supplied decoded LibreOffice native corpus"]
fn decoded_ole_inline_fixture_uses_opaque_native_empty_presentation_profile() {
    let Some(path) = std::env::var_os("LITCHI_OLE1_INLINE_RTF_OBJECT") else {
        return;
    };
    let bytes = std::fs::read(path).expect("decoded OLE1 fixture should be readable");
    let snapshot = ObjectSnapshot::parse_with_limits_and_compatibility(
        &bytes,
        Limits::default(),
        Compatibility::EmptyPresentationHeader,
    )
    .expect("native payload plus exact empty presentation header should parse");
    assert_eq!(
        snapshot.presentation().kind(),
        PresentationKind::RawFormatZero
    );
    assert_eq!(snapshot.presentation().bytes().len(), 8);
    assert!(snapshot.native_data().is_some_and(|data| data.starts_with(&[
        0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1
    ])));
    assert!(!snapshot.rewrite_safe());
    assert_eq!(snapshot.edit().commit().unwrap().snapshot().bytes(), bytes);
}
