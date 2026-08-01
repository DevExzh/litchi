use litchi_cfb::{OleFile, OleWriter};
use litchi_sign::cfb::{self, Editor, Format};
use litchi_sign::{Error, Limits, Policy, Signer, Status, Trust};
use p256::ecdsa::SigningKey as EcdsaSigningKey;
use rsa::{RsaPrivateKey, rand_core::OsRng};
use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

fn ole(streams: &[(&[&str], &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        if path.len() > 1 {
            writer.create_storage(&path[..path.len() - 1]).unwrap();
        }
        writer.create_stream(path, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn rsa_signer() -> Signer {
    let key = RsaPrivateKey::new(&mut OsRng, 2048).unwrap();
    Signer::rsa(key)
        .unwrap()
        .time("2026-07-19T12:34:56Z")
        .unwrap()
}

fn verify(bytes: &[u8], format: Format) -> litchi_sign::Result<Vec<cfb::Report>> {
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    cfb::verify(&mut file, format, &Policy::strict())
}

fn signature_xml(bytes: &[u8]) -> (String, Vec<u8>) {
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
    let name = file.list_directory_entries(&["_xmlsignatures"]).unwrap()[0]
        .name
        .clone();
    let xml = file.open_stream(&["_xmlsignatures", &name]).unwrap();
    (name, xml)
}

#[test]
fn rsa_and_ecdsa_signatures_round_trip_with_chain_time_and_multiple_streams() {
    let original = ole(&[(&["Payload"], b"signed bytes")]);
    let mut editor = Editor::open(original, Format::Doc).unwrap();
    editor.add(&rsa_signer()).unwrap();
    let ecdsa = Signer::p256(EcdsaSigningKey::random(&mut OsRng))
        .time("2026-07-19T12:35:56Z")
        .unwrap();
    editor.add(&ecdsa).unwrap();
    let signed = editor.finish().unwrap();

    let reports = verify(&signed, Format::Doc).unwrap();
    assert_eq!(reports.len(), 2);
    assert!(reports.iter().all(|report| {
        report.integrity() == Status::Valid && report.signature() == Status::Valid
    }));
    let first = reports
        .iter()
        .find(|report| report.stream() == "1")
        .unwrap();
    assert_eq!(first.time(), Some("2026-07-19T12:34:56Z"));
}

#[test]
fn stale_payload_is_reported_without_conflating_certificate_trust() {
    let original = ole(&[(&["Payload"], b"original")]);
    let mut editor = Editor::open(original, Format::Doc).unwrap();
    editor.add(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    let (name, xml) = signature_xml(&signed);
    let tampered = ole(&[
        (&["Payload"], b"tampered"),
        (&["_xmlsignatures", name.as_str()], xml.as_slice()),
    ]);

    let report = verify(&tampered, Format::Doc).unwrap().remove(0);
    assert_eq!(report.integrity(), Status::Invalid);
    assert_eq!(report.signature(), Status::Valid);
    assert_eq!(report.trust(), Trust::NotChecked);
}

#[test]
fn no_op_clear_and_resign_are_atomic_and_preserve_payload_streams() {
    let original = ole(&[(&["Payload"], b"preserve exactly")]);
    let editor = Editor::open(original.clone(), Format::Doc).unwrap();
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Editor::open(original, Format::Doc).unwrap();
    editor.add(&rsa_signer()).unwrap();
    editor.resign(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    assert_eq!(verify(&signed, Format::Doc).unwrap().len(), 1);

    let mut editor = Editor::open(signed, Format::Doc).unwrap();
    editor.clear().unwrap();
    let cleared = editor.finish().unwrap();
    assert!(verify(&cleared, Format::Doc).unwrap().is_empty());
    let mut file = OleFile::open(Cursor::new(cleared)).unwrap();
    assert_eq!(file.open_stream(&["Payload"]).unwrap(), b"preserve exactly");
}

#[test]
fn malformed_legacy_encrypted_and_resource_hostile_containers_are_rejected() {
    let duplicate_numeric = ole(&[
        (&["Payload"], b"x"),
        (&["_xmlsignatures", "1"], b"<Signature/>"),
        (&["_xmlsignatures", "01"], b"<Signature/>"),
    ]);
    assert!(matches!(
        verify(&duplicate_numeric, Format::Doc),
        Err(Error::Container(_))
    ));

    let legacy = ole(&[(&["Payload"], b"x"), (&["_signatures"], &[0, 0, 0, 0])]);
    assert!(matches!(verify(&legacy, Format::Doc), Err(Error::Legacy)));
    let encrypted = ole(&[(&["Payload"], b"x"), (&["EncryptionInfo"], b"encrypted")]);
    assert!(matches!(
        Editor::open(encrypted, Format::Doc),
        Err(Error::Encrypted)
    ));

    let original = ole(&[(&["Payload"], b"x")]);
    let mut editor = Editor::open(original, Format::Doc).unwrap();
    editor.add(&rsa_signer()).unwrap();
    let signed = editor.finish().unwrap();
    let limits = Limits::standard().signature_bytes(32).unwrap();
    let policy = Policy::strict().with_limits(limits);
    let mut file = OleFile::open(Cursor::new(signed)).unwrap();
    assert!(matches!(
        cfb::verify(&mut file, Format::Doc, &policy),
        Err(Error::Limit(_))
    ));
}

#[test]
fn producer_fixtures_are_noop_exact_and_facades_discover_unsigned_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for (relative, format) in [
        ("test-data/ole/doc/documentProperties.doc", Format::Doc),
        (
            "test-data/poi/test-data/spreadsheet/Simple.xls",
            Format::Xls,
        ),
        (
            "test-data/libreoffice-core/sc/qa/unit/data/xls/pivottable_number_grouping.xls",
            Format::Xls,
        ),
        ("test-data/ole/ppt/text-margins.ppt", Format::Ppt),
    ] {
        let bytes = std::fs::read(root.join(relative)).unwrap();
        let editor = Editor::open(bytes.clone(), format).unwrap();
        assert_eq!(editor.finish().unwrap(), bytes);
    }

    let mut doc =
        litchi_ole::doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc"))
            .unwrap();
    assert!(doc.signatures().unwrap().is_empty());
    let mut xls = litchi_ole::xls::XlsWorkbook::new(
        File::open(root.join("test-data/poi/test-data/spreadsheet/Simple.xls")).unwrap(),
    )
    .unwrap();
    assert!(xls.signatures().unwrap().is_empty());
    let mut ppt =
        litchi_ole::ppt::Package::open(root.join("test-data/ole/ppt/text-margins.ppt")).unwrap();
    assert!(ppt.signatures().unwrap().is_empty());
}
