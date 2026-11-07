//! Explicit caption/envelope derivatives of native Word sources; fields and mail stay inert.
use litchi_doc::{Package, captions as cap, parts::envelope as env};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let b = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, b)
        })
        .collect()
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    p.save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(q.to_bytes().unwrap() == p.to_bytes().unwrap());
    q
}
fn preserved(before: &Package, after: &Package, indices: &[usize]) {
    assert_eq!(
        cfb_metadata::directory(&before.to_bytes().unwrap()),
        cfb_metadata::directory(&after.to_bytes().unwrap())
    );
    let a = streams(&before.to_bytes().unwrap());
    let b = streams(&after.to_bytes().unwrap());
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    let table = format!(
        "{}Table",
        u8::from(before.document().fib().which_table_stream())
    );
    for (path, data) in &a {
        if path.as_slice() == std::slice::from_ref(&table) {
            assert!(
                b[path].starts_with(data),
                "selected table lost original bytes"
            );
        } else if path == &["WordDocument".to_owned()] {
            let mut x = data.clone();
            let mut y = b[path].clone();
            assert_eq!(x.len(), y.len());
            for index in indices {
                let at = 154 + index * 8;
                x[at..at + 8].fill(0);
                y[at..at + 8].fill(0);
            }
            assert!(x == y, "unowned WordDocument bytes changed");
        } else {
            assert!(data == &b[path], "changed {}", path.join("/"));
        }
    }
    assert_eq!(
        before.document().text().unwrap(),
        after.document().text().unwrap()
    );
}
fn captions() -> cap::Tables {
    let labels = cap::LabelTable::try_new(vec![
        cap::Definition::try_new(
            "図表 🍎".into(),
            cap::Info::new(
                cap::Location::Above,
                Some(cap::Numbering::new(
                    cap::Heading::Level3,
                    cap::Separator::EmDash,
                )),
                false,
                cap::Format::Arabic,
            ),
        )
        .unwrap(),
        cap::Definition::try_new(
            "Table".into(),
            cap::Info::new(cap::Location::Below, None, true, cap::Format::Arabic),
        )
        .unwrap(),
    ])
    .unwrap();
    let auto = cap::AutoTable::try_new(vec![
        cap::AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
        cap::AutoEntry::try_new("Excel.Sheet.8".into(), 1).unwrap(),
    ])
    .unwrap();
    cap::Tables::try_new(Some(labels), Some(auto)).unwrap()
}
fn caption_publish(source: &Package, c: &cap::Commit) -> Package {
    let q = save(c.package());
    assert!(
        source
            .apply_captions(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == q.to_bytes().unwrap()
    );
    assert!(
        q.apply_captions(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == source.to_bytes().unwrap()
    );
    q
}
#[test]
fn native_dot_caption_lifecycle_preserves_fields_glossary_and_macro_streams() {
    for file in [
        "word-layout-extensions.dot",
        "word-glossary.dot",
        "word-recorded-macro.dot",
    ] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.document().fib().is_template());
        let document = p.document();
        let empty = document.caption_tables().unwrap();
        assert!(empty.labels().is_none() && empty.auto().is_none());
        let mut e = p.edit_captions().unwrap();
        e.set(captions()).unwrap();
        let c = e.commit().unwrap();
        let added = caption_publish(&p, &c);
        preserved(&p, &added, &[52, 53]);
        assert_eq!(added.document().caption_tables().unwrap(), &captions());
        let mut e = added.edit_captions().unwrap();
        e.set(captions()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        assert!(matches!(
            added.apply_captions(c.patch()),
            Err(cap::Error::PatchConflict)
        ));
        let mut e = added.edit_captions().unwrap();
        let labels = cap::LabelTable::try_new(vec![
            cap::Definition::try_new(
                "🍎".repeat(20),
                cap::Info::new(cap::Location::Below, None, false, cap::Format::Arabic),
            )
            .unwrap(),
            captions().labels().unwrap().definitions()[1].clone(),
        ])
        .unwrap();
        e.replace_labels(labels.clone()).unwrap();
        let c = e.commit().unwrap();
        let updated = caption_publish(&added, &c);
        assert_eq!(
            updated.document().caption_tables().unwrap().labels(),
            Some(&labels)
        );
        assert_eq!(
            updated.document().caption_tables().unwrap().auto(),
            captions().auto()
        );
        preserved(&added, &updated, &[52, 53]);
        let mut e = updated.edit_captions().unwrap();
        e.clear_auto().unwrap();
        e.clear_labels().unwrap();
        let c = e.commit().unwrap();
        let cleared = caption_publish(&updated, &c);
        preserved(&updated, &cleared, &[52, 53]);
        for index in [52, 53] {
            assert_eq!(
                cleared.document().fib().get_table_pointer(index),
                Some((0, 0))
            );
        }
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_dot_caption_invalid_reference_and_label_limit_leave_staging_intact() {
    let bytes = native("word-recorded-macro.dot");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_captions().unwrap();
    e.set(captions()).unwrap();
    let error = e.clear_labels().unwrap_err();
    assert!(
        matches!(error,cap::Error::Package(litchi_doc::Error::InvalidFormat(m)) if m=="caption transaction failed: Corrupted file: SttbfAutoCaption entry 0 references a missing SttbfCaption label")
    );
    let error = cap::Definition::try_new(
        "🍎".repeat(21),
        cap::Info::new(cap::Location::Above, None, false, cap::Format::Arabic),
    )
    .unwrap_err();
    assert!(
        matches!(error,litchi_doc::Error::Corrupted(m) if m=="caption label exceeds 40 UTF-16 code units")
    );
    assert_eq!(e.tables().unwrap(), &captions());
    let c = e.commit().unwrap();
    let q = caption_publish(&p, &c);
    assert_eq!(q.document().caption_tables().unwrap(), &captions());
    preserved(&p, &q, &[52, 53]);
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_doc_caption_edit_refuses_the_non_template_owner() {
    for file in ["word-layout-extensions.doc", "word-recorded-macro.doc"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(!p.document().fib().is_template());
        for index in [52, 53] {
            assert_eq!(p.document().fib().get_table_pointer(index).unwrap().1, 0);
        }
        let mut e = p.edit_captions().unwrap();
        assert!(
            matches!(e.set(captions()),Err(cap::Error::Package(litchi_doc::Error::Corrupted(m))) if m=="caption tables are only writable in a Normal template")
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
fn unicode(s: &str) -> env::Text {
    env::Text::unicode(s.encode_utf16().collect())
}
fn message() -> env::Message {
    env::Message {
        last_sent_time: 123456,
        flag_status: env::FollowUpStatus::Flagged,
        request: unicode("Please review 🍎"),
        subject: unicode("NATIVE_ENVELOPE_067 日本語"),
        sent_representing_name: unicode("Ada"),
        internet_account_name: unicode("ada@example.invalid"),
        categories: unicode("inert;review"),
        security: env::SecurityFlags {
            signed: true,
            encrypted: true,
        },
        importance: env::Importance::High,
        sensitivity: env::Sensitivity::Private,
        recipients: env::RecipientCollection {
            recipients: vec![env::RecipientProperties {
                properties: vec![
                    env::RecipientProperty {
                        property_id: 0x3001,
                        value: env::PropertyValue::Unicode("Recipient 🍎".encode_utf16().collect()),
                    },
                    env::RecipientProperty {
                        property_id: 0x3003,
                        value: env::PropertyValue::String8(
                            b"recipient@example.invalid".to_vec().into_boxed_slice(),
                        ),
                    },
                    env::RecipientProperty {
                        property_id: 0x0001,
                        value: env::PropertyValue::Binary(vec![0, 0xff, 6, 7].into_boxed_slice()),
                    },
                ],
            }],
        },
        attachments: vec![env::Attachment {
            method: 1,
            name: "inert.bin".encode_utf16().collect(),
            data: vec![0, 0xff, 0x67].into_boxed_slice(),
        }],
        intro_text: Some("Intro\n日本語".encode_utf16().collect()),
        tail: vec![0xfa, 0xce, 0x67].into_boxed_slice(),
        ..Default::default()
    }
}
fn envelope_publish(source: &Package, c: &env::Commit) -> Package {
    let q = save(c.package());
    assert!(
        source
            .apply_envelope(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == q.to_bytes().unwrap()
    );
    assert!(
        q.apply_envelope(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == source.to_bytes().unwrap()
    );
    q
}
#[test]
fn native_doc_and_macro_dot_envelope_lifecycle_retains_typed_and_opaque_payloads() {
    for file in ["word-text-structure.doc", "word-recorded-macro.dot"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.document().envelope().is_none());
        let value = env::Envelope::from_message(message()).unwrap();
        let mut e = p.edit_envelope().unwrap();
        e.set(value.clone()).unwrap();
        let c = e.commit().unwrap();
        let added = envelope_publish(&p, &c);
        assert_eq!(added.document().envelope(), Some(&value));
        preserved(&p, &added, &[env::FIB_INDEX]);
        assert!(matches!(
            added.apply_envelope(c.patch()),
            Err(env::Error::PatchConflict)
        ));
        let mut e = added.edit_envelope().unwrap();
        e.set(value.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_envelope().unwrap();
        let error = e
            .update_message(|m| m.contact_link_recipients = None)
            .unwrap_err();
        assert!(
            matches!(error,env::Error::Package(litchi_doc::Error::InvalidFormat(m)) if m=="MsoEnvelope transaction failed: Corrupted file: Office 8 envelopes must contain contact-link recipients")
        );
        assert_eq!(e.value().unwrap(), Some(&value));
        e.update_message(|m| m.subject = unicode("Updated subject 🍎"))
            .unwrap();
        let mut expected = message();
        expected.subject = unicode("Updated subject 🍎");
        let c = e.commit().unwrap();
        let updated = envelope_publish(&added, &c);
        assert_eq!(
            updated.document().envelope().unwrap().message(),
            Some(&expected)
        );
        preserved(&added, &updated, &[env::FIB_INDEX]);
        let mut office6 = expected;
        office6.version = env::Version::Office6;
        office6.request = env::Text::ansi(b"reply\x80".to_vec());
        office6.sent_representing_name = env::Text::ansi(b"sender".to_vec());
        office6.internet_account_stamp = env::Text::ansi(Vec::new());
        office6.internet_account_name = env::Text::ansi(b"account@example.invalid".to_vec());
        office6.categories = env::Text::ansi(b"inert".to_vec());
        office6.subject = env::Text::ansi(b"ANSI subject\x80".to_vec());
        office6.contact_link_recipients = None;
        office6.intro_text = None;
        let mut e = updated.edit_envelope().unwrap();
        e.set_message(office6.clone()).unwrap();
        let c = e.commit().unwrap();
        let ansi = envelope_publish(&updated, &c);
        preserved(&updated, &ansi, &[env::FIB_INDEX]);
        assert_eq!(
            ansi.document().envelope().unwrap().message(),
            Some(&office6)
        );
        let opaque = env::Envelope::opaque([0x67; 16], vec![0, 0xff, 1, 2, 3]).unwrap();
        let mut e = ansi.edit_envelope().unwrap();
        e.replace(opaque.clone()).unwrap();
        let c = e.commit().unwrap();
        let replaced = envelope_publish(&ansi, &c);
        assert_eq!(replaced.document().envelope(), Some(&opaque));
        assert!(replaced.document().envelope().unwrap().message().is_none());
        preserved(&ansi, &replaced, &[env::FIB_INDEX]);
        let mut rejected = replaced.edit_envelope().unwrap();
        assert!(
            matches!(rejected.update_message(|m|m.subject=unicode("must refuse")),Err(env::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m=="MsoEnvelope transaction failed: the envelope CLSID does not select a supported message")
        );
        assert!(rejected.commit().unwrap().patch().is_noop());
        let mut e = replaced.edit_envelope().unwrap();
        e.clear().unwrap();
        let c = e.commit().unwrap();
        let cleared = envelope_publish(&replaced, &c);
        assert!(cleared.document().envelope().is_none());
        assert_eq!(
            cleared.document().fib().get_table_pointer(env::FIB_INDEX),
            Some((0, 0))
        );
        preserved(&replaced, &cleared, &[env::FIB_INDEX]);
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_doc_envelope_protection_refusal_preserves_exact_noop() {
    let bytes = native("word-document-protection.doc");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut e = p.edit_envelope().unwrap();
    assert!(e.value().unwrap().is_none());
    assert!(matches!(
        e.set_message(message()),
        Err(env::Error::Package(litchi_doc::Error::ProtectionDenied(
            litchi_doc::parts::protection::EditProtection::Document
        )))
    ));
    assert!(e.commit().unwrap().patch().is_noop());
    assert!(p.to_bytes().unwrap() == bytes);
}
