//! XLS creation and publication through the unified facade.
#![cfg(feature = "xls")]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, Conformance, CreateOptions, File, Flavor, create, create_with};

#[test]
fn minimal_creation_publishes_and_reopens() {
    let file = create(FileFormat::Xls).expect("create XLS");
    assert!(matches!(file, File::Xls(_)));
    assert_eq!(
        file.workbook().unwrap().worksheet_names().unwrap(),
        ["Sheet1"]
    );
    assert!(file.text().unwrap().trim().is_empty());
    let bytes = file.to_bytes().unwrap();
    let reopened = litchi_ops::open_bytes(bytes.clone()).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), bytes);
    assert_eq!(file.save().unwrap_err().code(), Code::FileRequired);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("created.xls");
    file.save_as(&path).unwrap();
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
    for mode in [litchi_ops::SourceMode::Auto, litchi_ops::SourceMode::Eager] {
        let opened =
            litchi_ops::open_with(&path, &litchi_ops::OpenOptions::new().source_mode(mode))
                .unwrap();
        opened.save().unwrap();
        assert_eq!(std::fs::read(&path).unwrap(), bytes);
    }
    assert_eq!(
        file.save_as(directory.path().join("wrong.xlsx"))
            .unwrap_err()
            .code(),
        Code::InvalidArgument
    );
}

#[test]
fn creation_options_are_explicit_refusals_and_capabilities_agree() {
    for options in [
        CreateOptions::new().flavor(Flavor::Template),
        CreateOptions::new().conformance(Conformance::Strict),
        CreateOptions::new().metadata(Metadata {
            identifier: Some("urn:x".to_owned()),
            ..Metadata::default()
        }),
    ] {
        let error = create_with(FileFormat::Xls, &options).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Xls));
    }
    let capabilities = Capabilities::of(FileFormat::Xls);
    assert!(capabilities.verb(Verb::Create).is_supported());
    assert!(capabilities.verb(Verb::Save).is_supported());
    let gaps = format!("{:?}", capabilities.gaps().collect::<Vec<_>>());
    for context in [
        "template",
        "strict-conformance",
        "unrepresentable-initial-metadata",
    ] {
        assert!(gaps.contains(context), "missing {context}: {gaps}");
    }
}

#[test]
fn publication_preserves_signatures_and_requires_explicit_removal() {
    use std::io::Cursor;
    let original = create(FileFormat::Xls).unwrap().to_bytes().unwrap();
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(original), litchi_cfb::OleFileLimits::default())
            .unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer
        .create_stream(&["_signatures"], b"opaque-signature")
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let bytes = bytes.into_inner();
    let file = litchi_ops::open_bytes(bytes.clone()).unwrap();
    assert!(file.is_signed());
    assert_eq!(file.to_bytes().unwrap(), bytes);
    let File::Signed(signed) = file else {
        panic!("signed owner")
    };
    let unsigned = signed.unsign();
    let emitted = unsigned.to_bytes().unwrap();
    assert!(!litchi_ops::open_bytes(emitted.clone()).unwrap().is_signed());
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(emitted), litchi_cfb::OleFileLimits::default())
            .unwrap();
    assert_eq!(ole.open_stream(&["Workbook"]).unwrap(), workbook);
}

#[test]
fn atomic_save_refuses_overwrite_and_changed_sources() {
    use litchi_ops::{Overwrite, SaveOptions, SaveTo};
    use std::io::Write;
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.xls");
    let destination = directory.path().join("destination.xls");
    let created = create(FileFormat::Xls).unwrap();
    created.save_as(&source).unwrap();
    std::fs::write(&destination, b"keep").unwrap();
    let options = SaveOptions::new().overwrite(Overwrite::Refuse);
    assert!(
        created
            .save_with(SaveTo::Path(&destination), &options)
            .is_err()
    );
    assert_eq!(std::fs::read(&destination).unwrap(), b"keep");
    let opened = litchi_ops::open(&source).unwrap();
    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(&source)
        .unwrap();
    changed.write_all(&[0]).unwrap();
    changed.flush().unwrap();
    assert!(opened.to_bytes().is_err());
    assert!(opened.save_as(&destination).is_err());
    assert_eq!(std::fs::read(&destination).unwrap(), b"keep");
}

#[cfg(feature = "encryption")]
#[test]
fn unlocked_encrypted_owner_refuses_plain_publication() {
    let mut builder = litchi_xls::Builder::new();
    builder.add_worksheet("Sheet1").unwrap();
    builder
        .set_password("secret", litchi_xls::EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    let output = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("protected.xls");
    std::fs::write(&source, output).unwrap();
    let file = litchi_ops::open(&source).unwrap();
    let File::Locked(locked) = file else {
        panic!("locked owner")
    };
    let unlocked = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap();
    assert!(!unlocked.capabilities().verb(Verb::Save).is_supported());
    assert_eq!(unlocked.save().unwrap_err().code(), Code::DocumentProtected);
    assert_eq!(
        unlocked.to_bytes().unwrap_err().code(),
        Code::DocumentProtected
    );
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("encrypted.xls");
    std::fs::write(&destination, b"keep").unwrap();
    assert_eq!(
        unlocked.save_as(&destination).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(std::fs::read(destination).unwrap(), b"keep");
}
