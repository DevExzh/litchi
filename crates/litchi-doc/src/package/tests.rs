use super::{Error, OpenOptions};
use crate::raw::{OpenOutcome, Reader, open_package};
use litchi_cfb::{OleFile, OleWriter};
use litchi_crypto::spaces::{self, Definition, Header, Map, MapEntry, Reference, ReferenceKind};
use std::io::Cursor;
use std::path::Path;

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_open_package() {
    let result = Reader::open("test.doc");
    assert!(result.is_ok());
}

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_invalid_file() {
    // Create a non-DOC file
    std::fs::write("test_invalid.tmp", b"Not a DOC file").unwrap();
    let result = Reader::open("test_invalid.tmp");
    assert!(result.is_err());
    std::fs::remove_file("test_invalid.tmp").ok();
}

fn poi_fixture(name: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document")
        .join(name)
}

#[test]
fn opens_apache_poi_binary_rc4_document() {
    let path = poi_fixture("password_tika_binaryrc4.doc");

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(
        package.document_with_options(OpenOptions::default().password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Reader::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().password("tika".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn opens_apache_poi_cryptoapi_document() {
    let path = poi_fixture("password_password_cryptoapi.doc");

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(
        package.document_with_options(OpenOptions::default().password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Reader::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().password("password".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn raw_open_outcome_retains_encrypted_cfb_for_later_password_ingress() {
    let bytes = std::fs::read(poi_fixture("password_tika_binaryrc4.doc")).unwrap();
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let ole = match open_package(ole, OpenOptions::default()).unwrap() {
        OpenOutcome::PasswordRequired(ole) => *ole,
        OpenOutcome::Opened(_) => panic!("encrypted fixture unexpectedly opened without password"),
    };
    let package = match open_package(
        ole,
        OpenOptions::default().password("tika".to_owned().into()),
    )
    .unwrap()
    {
        OpenOutcome::Opened(package) => package,
        OpenOutcome::PasswordRequired(_) => panic!("correct password remained locked"),
    };
    assert!(!package.document().text().unwrap().trim().is_empty());
}

#[test]
fn ordinary_doc_exposes_an_empty_dataspaces_profile() {
    let mut package = Reader::open(poi_fixture("test.doc")).unwrap();
    assert!(package.data_spaces().unwrap().is_none());
    assert!(package.data_spaces_snapshot().unwrap().is_none());
}

fn data_spaces_doc_bytes() -> Vec<u8> {
    let limits = spaces::Limits::default();
    let map = spaces::write_map(
        &Map {
            entries: vec![MapEntry {
                references: vec![Reference {
                    kind: ReferenceKind::Stream,
                    component: "Payload".to_string(),
                }],
                data_space_name: "DataSpace".to_string(),
            }],
        },
        &limits,
    )
    .unwrap();
    let definition = spaces::write_definition(
        &Definition {
            transforms: vec!["FirstTransform".to_string()],
        },
        &limits,
    )
    .unwrap();
    let header = spaces::write_transform_header(
        &Header {
            transform_id: "{11111111-1111-1111-1111-111111111111}".to_string(),
            transform_name: "FirstTransform".to_string(),
            reader: spaces::Version::V1_0,
            updater: spaces::Version::V1_0,
            writer: spaces::Version::V1_0,
        },
        &limits,
    )
    .unwrap();

    let mut writer = OleWriter::new();
    writer.create_stream(&["WordDocument"], b"doc").unwrap();
    writer
        .create_stream(&["Payload"], b"inert payload")
        .unwrap();
    writer.create_storage(&[spaces::STORAGE]).unwrap();
    writer
        .create_storage(&[spaces::STORAGE, "DataSpaceInfo"])
        .unwrap();
    writer
        .create_storage(&[spaces::STORAGE, "TransformInfo"])
        .unwrap();
    writer
        .create_storage(&[spaces::STORAGE, "TransformInfo", "FirstTransform"])
        .unwrap();
    writer
        .create_stream(
            &[spaces::STORAGE, "Version"],
            &spaces::write_version_info(&spaces::VersionInfo::default(), &limits).unwrap(),
        )
        .unwrap();
    writer
        .create_stream(&[spaces::STORAGE, "DataSpaceMap"], &map)
        .unwrap();
    writer
        .create_stream(
            &[spaces::STORAGE, "DataSpaceInfo", "DataSpace"],
            &definition,
        )
        .unwrap();
    writer
        .create_stream(
            &[
                spaces::STORAGE,
                "TransformInfo",
                "FirstTransform",
                spaces::PRIMARY,
            ],
            &header,
        )
        .unwrap();
    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    bytes.into_inner()
}

#[test]
fn raw_reader_retains_dataspaces_limits_for_graph_projections() {
    let bytes = data_spaces_doc_bytes();

    let mut default_reader = Reader::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert!(default_reader.data_spaces().unwrap().is_some());
    assert!(default_reader.data_spaces_snapshot().unwrap().is_some());

    let selected = spaces::Limits {
        max_stream_bytes: 1,
        ..spaces::Limits::default()
    };
    let mut limited_reader = Reader::from_reader_with_options(
        Cursor::new(bytes),
        OpenOptions::default().data_spaces_limits(selected),
    )
    .unwrap();
    assert_eq!(limited_reader.data_spaces_limits(), selected);
    assert!(matches!(
        limited_reader.data_spaces(),
        Err(Error::Crypto(spaces::Error::LimitExceeded {
            kind: spaces::LimitKind::StreamBytes,
            ..
        }))
    ));
    assert!(matches!(
        limited_reader.data_spaces_snapshot(),
        Err(Error::Crypto(spaces::Error::LimitExceeded {
            kind: spaces::LimitKind::StreamBytes,
            ..
        }))
    ));
}
