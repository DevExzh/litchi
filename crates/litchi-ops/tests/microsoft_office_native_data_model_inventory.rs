//! Corpus-only Data Model evidence. Explicit shared decryption does not add XLSB ingress support.
#![cfg(all(feature = "xlsx", feature = "xlsb", feature = "encryption"))]
use litchi_opc::OpcPackage;
use litchi_xlsx::workbook::data_model as model;
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::path::Path;
fn root() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office")
}
fn files(exts: &[&str]) -> Vec<String> {
    let m: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root().join("manifest.json")).unwrap()).unwrap();
    m["files"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|f| {
            let n = f["path"].as_str().unwrap();
            (n.starts_with("excel/")
                && Path::new(n)
                    .extension()
                    .and_then(|e| e.to_str())
                    .is_some_and(|e| exts.contains(&e)))
            .then(|| n.to_owned())
        })
        .collect()
}
fn inspect(p: &OpcPackage) {
    for part in p.iter_parts() {
        assert_ne!(part.content_type(), model::DATA_MODEL_CONTENT_TYPE);
        assert!(
            !part
                .partname()
                .as_str()
                .eq_ignore_ascii_case(model::DATA_MODEL_PART_NAME)
        );
        if !part.partname().as_str().ends_with(".xml")
            && !part.content_type().ends_with("+xml")
            && !matches!(part.content_type(), "text/xml" | "application/xml")
        {
            continue;
        }
        let data = p.get_part(part.partname()).unwrap().blob();
        let xml = if data.starts_with(&[0xff, 0xfe]) || data.starts_with(&[0xfe, 0xff]) {
            assert_eq!((data.len() - 2) % 2, 0);
            let little = data[0] == 0xff;
            std::borrow::Cow::Owned(
                String::from_utf16(
                    &data[2..]
                        .chunks_exact(2)
                        .map(|b| {
                            if little {
                                u16::from_le_bytes([b[0], b[1]])
                            } else {
                                u16::from_be_bytes([b[0], b[1]])
                            }
                        })
                        .collect::<Vec<_>>(),
                )
                .unwrap(),
            )
        } else {
            std::borrow::Cow::Borrowed(std::str::from_utf8(data).unwrap())
        };
        let mut r = NsReader::from_str(&xml);
        loop {
            let (ns, e) = r.read_resolved_event().unwrap();
            match e {
                Event::Start(e) | Event::Empty(e) => {
                    let namespace = match &ns {
                        ResolveResult::Bound(n) => n.as_ref(),
                        ResolveResult::Unbound => b"",
                        ResolveResult::Unknown(_) => {
                            panic!("unbound prefix in {}", part.partname())
                        },
                    };
                    assert!(
                        !(namespace
                            == b"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
                            && e.local_name().as_ref() == b"dataModel")
                    );
                    assert!(
                        !(namespace == model::MODEL_TIME_GROUPINGS_NAMESPACE.as_bytes()
                            && e.local_name().as_ref() == b"modelTimeGroupings")
                    );
                    for a in e.attributes() {
                        let a = a.unwrap();
                        if a.key.local_name().as_ref() == b"uri" {
                            let value = a
                                .decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Explicit1_0,
                                    r.decoder(),
                                )
                                .unwrap();
                            assert!(
                                !value
                                    .trim()
                                    .eq_ignore_ascii_case(model::DATA_MODEL_EXTENSION_URI)
                            );
                            assert!(
                                !value.trim().eq_ignore_ascii_case(
                                    model::MODEL_TIME_GROUPINGS_EXTENSION_URI
                                )
                            );
                        }
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
    }
}
#[test]
fn native_xlsx_corpus_has_connections_and_pivots_without_data_model_owners() {
    let names = files(&["xlsx", "xlsm", "xltx", "xltm"]);
    assert_eq!(names.len(), 19);
    let mut encrypted = 0;
    let mut connections = false;
    let mut pivot = false;
    for name in names {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        let p = if name == "excel/excel-encrypted.xlsx" {
            encrypted += 1;
            litchi_xlsx::Package::from_bytes_with_password(&bytes, "OfficeFixture42!").unwrap()
        } else {
            litchi_xlsx::Package::from_bytes(&bytes).unwrap()
        };
        assert!(p.data_model().unwrap().is_empty());
        assert!(p.data_model().unwrap().model().is_none());
        let opc = p.clone().into_plain_opc();
        inspect(&opc);
        if name == "excel/excel-connections.xlsx" {
            assert!(
                opc.iter_parts()
                    .any(|p| p.partname().as_str() == "/xl/connections.xml")
            );
            connections = true;
        }
        if name == "excel/excel-visual-analytics.xlsx" {
            assert!(
                opc.iter_parts()
                    .any(|p| p.partname().as_str().contains("/pivotTables/"))
            );
            pivot = true;
        }
        if name != "excel/excel-encrypted.xlsx" {
            assert!(p.to_bytes().unwrap() == bytes);
        }
        assert!(std::fs::read(root().join(name)).unwrap() == bytes);
    }
    assert_eq!((encrypted, connections, pivot), (1, true, true));
}
#[test]
fn native_xlsb_corpus_including_explicitly_decrypted_fixture_has_no_model_records() {
    use litchi_xlsb::raw::{Limits, Records, kind};
    let names = files(&["xlsb"]);
    assert_eq!(names.len(), 16);
    let mut encrypted = 0;
    let mut connections = false;
    let mut pivot = false;
    for name in names {
        let bytes = std::fs::read(root().join(&name)).unwrap();
        let clear = if name == "excel/excel-encrypted.xlsb" {
            encrypted += 1;
            assert!(matches!(
                litchi_xlsb::Package::from_bytes(&bytes),
                Err(litchi_xlsb::package::PackageError::UnsupportedFeature(
                    "CFB containers at the XLSB boundary (including encrypted packages)"
                ))
            ));
            assert!(matches!(
                litchi_crypto::ooxml::open(
                    bytes.clone(),
                    "wrong",
                    &litchi_crypto::ooxml::Limits::default()
                ),
                Err(litchi_crypto::ooxml::Error::Password)
            ));
            let opened = litchi_crypto::ooxml::open(
                bytes.clone(),
                "OfficeFixture42!",
                &litchi_crypto::ooxml::Limits::default(),
            )
            .unwrap();
            assert_eq!(
                opened.integrity(),
                Some(litchi_crypto::ooxml::IntegrityStatus::Authenticated)
            );
            opened.into_bytes()
        } else {
            bytes.clone()
        };
        let p = litchi_xlsb::Package::from_bytes(&clear).unwrap();
        let snapshot = p.data_model().unwrap();
        assert!(!snapshot.is_present());
        assert!(snapshot.definition().is_none() && snapshot.part().is_none());
        let opc = p.opc_package();
        inspect(opc);
        let workbook = opc.main_document_part().unwrap();
        let mut records = 0;
        for record in Records::new(workbook.blob(), Limits::default()) {
            let record = record.unwrap();
            assert!(
                ![
                    kind::BEGIN_DATA_MODEL,
                    kind::END_DATA_MODEL,
                    kind::BEGIN_MODEL_TABLES,
                    kind::END_MODEL_TABLES,
                    kind::MODEL_TABLE,
                    kind::BEGIN_MODEL_RELATIONSHIPS,
                    kind::END_MODEL_RELATIONSHIPS,
                    kind::MODEL_RELATIONSHIP,
                    kind::BEGIN_MODEL_TIME_GROUPINGS,
                    kind::END_MODEL_TIME_GROUPINGS,
                    kind::BEGIN_MODEL_TIME_GROUPING,
                    kind::END_MODEL_TIME_GROUPING,
                    kind::MODEL_TIME_GROUPING_CALC_COL
                ]
                .contains(&record.kind()),
                "{name}: model record"
            );
            records += 1;
        }
        assert!(records > 2);
        if name == "excel/excel-connections.xlsb" {
            assert!(snapshot.connection_names().is_some_and(|n| !n.is_empty()));
            connections = true;
        }
        if name == "excel/excel-visual-analytics.xlsb" {
            assert!(
                opc.iter_parts()
                    .any(|p| p.partname().as_str().contains("/pivotTables/"))
            );
            pivot = true;
        }
        assert!(p.to_bytes().unwrap() == clear);
        assert!(std::fs::read(root().join(name)).unwrap() == bytes);
    }
    assert_eq!((encrypted, connections, pivot), (1, true, true));
}
