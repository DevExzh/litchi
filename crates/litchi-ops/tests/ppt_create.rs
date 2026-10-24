//! PPT creation and publication through the unified facade.
#![cfg(feature = "ppt")]

use litchi_core::{FileFormat, Metadata};
use litchi_ops::capabilities::{Capabilities, Verb};
use litchi_ops::{Code, Conformance, CreateOptions, File, Flavor, create, create_with};

#[test]
fn minimal_creation_publishes_and_reopens() {
    let file = create(FileFormat::Ppt).expect("create PPT");
    assert!(matches!(file, File::Ppt(_)));
    assert_eq!(file.presentation().unwrap().slide_count().unwrap(), 0);
    assert!(file.presentation().unwrap().slides().unwrap().is_empty());
    let bytes = file.to_bytes().unwrap();
    let reopened = litchi_ops::open_bytes(bytes.clone()).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), bytes);
    assert_eq!(file.save().unwrap_err().code(), Code::FileRequired);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("created.ppt");
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
        file.save_as(directory.path().join("wrong.pptx"))
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
        let error = create_with(FileFormat::Ppt, &options).unwrap_err();
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert_eq!(error.format(), Some(FileFormat::Ppt));
    }
    let capabilities = Capabilities::of(FileFormat::Ppt);
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
    let original = create(FileFormat::Ppt).unwrap().to_bytes().unwrap();
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(original), litchi_cfb::OleFileLimits::default())
            .unwrap();
    let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let current_user = ole.open_stream(&["Current User"]).unwrap();
    let mut writer = litchi_cfb::OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &document)
        .unwrap();
    writer
        .create_stream(&["Current User"], &current_user)
        .unwrap();
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
    assert_eq!(ole.open_stream(&["PowerPoint Document"]).unwrap(), document);
    assert_eq!(ole.open_stream(&["Current User"]).unwrap(), current_user);
}

#[test]
fn atomic_save_refuses_overwrite_and_changed_sources() {
    use litchi_ops::{Overwrite, SaveOptions, SaveTo};
    use std::io::Write;
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.ppt");
    let destination = directory.path().join("destination.ppt");
    let created = create(FileFormat::Ppt).unwrap();
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

#[test]
fn released_edit_source_refuses_publication_but_keeps_reads() {
    let bytes = create(FileFormat::Ppt).unwrap().to_bytes().unwrap();
    let file =
        litchi_ops::open_bytes_with(bytes, &litchi_ops::OpenOptions::new().retain_source(false))
            .unwrap();
    assert_eq!(file.presentation().unwrap().slide_count().unwrap(), 0);
    assert!(!file.capabilities().verb(Verb::Save).is_supported());
    assert!(file.to_bytes().is_err());
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_publication_is_an_explicit_refusal() {
    let mut builder = litchi_ppt::Builder::new();
    builder.add_slide().unwrap();
    builder
        .set_password(
            "secret",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    builder.set_encryption_entropy(litchi_ppt::EncryptionEntropy::from_bytes([3; 16], [5; 16]));
    let bytes = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let File::Locked(locked) = litchi_ops::open_bytes(bytes).unwrap() else {
        panic!("locked presentation")
    };
    let file = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap();
    assert!(!file.capabilities().verb(Verb::Save).is_supported());
    assert_eq!(file.to_bytes().unwrap_err().code(), Code::DocumentProtected);
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("protected.ppt");
    std::fs::write(&path, b"keep").unwrap();
    assert_eq!(
        file.save_as(&path).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert_eq!(std::fs::read(path).unwrap(), b"keep");
}

#[cfg(feature = "encryption")]
#[test]
fn unsigning_cannot_bypass_encrypted_publication_policy() {
    use std::io::Cursor;
    let mut builder = litchi_ppt::Builder::new();
    builder.add_slide().unwrap();
    builder
        .set_password(
            "secret",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    builder.set_encryption_entropy(litchi_ppt::EncryptionEntropy::from_bytes(
        [19; 16], [23; 16],
    ));
    let bytes = builder.build().unwrap().to_encrypted_bytes().unwrap();
    let mut original =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    let mut signed = litchi_cfb::OleWriter::new();
    for name in ["PowerPoint Document", "Current User"] {
        signed
            .create_stream(&[name], &original.open_stream(&[name]).unwrap())
            .unwrap();
    }
    signed
        .create_stream(&["_signatures"], b"signature")
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    signed.write_to(&mut bytes).unwrap();
    let file = litchi_ops::open_bytes(bytes.into_inner()).unwrap();
    // Signature and password wrappers can appear in either order.
    let file = match file {
        File::Signed(signed) => signed.unsign(),
        file => file,
    };
    let File::Locked(locked) = file else {
        panic!("locked presentation")
    };
    let file = locked
        .unlock(litchi_ops::Password::new("secret".to_owned()))
        .unwrap();
    let file = match file {
        File::Signed(signed) => signed.unsign(),
        file => file,
    };
    let mut sink = Vec::new();
    assert_eq!(
        file.write_to(&mut sink).unwrap_err().code(),
        Code::DocumentProtected
    );
    assert!(sink.is_empty());
}
