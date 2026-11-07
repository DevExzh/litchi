//! Library-authored signatures over native inputs; no native-app acceptance or PKI trust is claimed.
#![cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "sign"))]
use litchi_opc::{OpcPackage, PackageWriter, SourceLimits};
use litchi_ops::{Code, File, OpenOptions, SourceMode, props};
use litchi_sign::{Coverage, Limits, Policy, Signer, Status, Trust, cfb};
use std::{
    collections::BTreeMap,
    io::{Cursor, Write},
    path::Path,
};
const TIME: &str = "2026-11-07T00:00:00Z";
fn native_file(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn native(extension: &str) -> Vec<u8> {
    native_file(&format!("excel/excel-metadata.{extension}"))
}
fn cfb_format(extension: &str) -> Option<cfb::Format> {
    match extension {
        "doc" => Some(cfb::Format::Doc),
        "ppt" => Some(cfb::Format::Ppt),
        "xls" => Some(cfb::Format::Xls),
        _ => None,
    }
}

fn signer() -> Signer {
    // Public, deterministic fixture key. It is not an identity or trust anchor.
    Signer::p256(p256::ecdsa::SigningKey::from_bytes((&[7u8; 32]).into()).unwrap())
        .time(TIME)
        .unwrap()
}
fn opc(bytes: &[u8]) -> OpcPackage {
    OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut file =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    file.list_streams()
        .into_iter()
        .map(|path| {
            let data = file
                .open_stream(&path.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (path, data)
        })
        .collect()
}
fn cfb_bytes(streams: &BTreeMap<Vec<String>, Vec<u8>>) -> Vec<u8> {
    let mut writer = litchi_cfb::OleWriter::new();
    for (name, data) in streams {
        let names = name.iter().map(String::as_str).collect::<Vec<_>>();
        if names.len() > 1 {
            writer.create_storage(&names[..names.len() - 1]).unwrap();
        }
        writer.create_stream(&names, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let p = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        SourceLimits::default(),
    )
    .unwrap();
    p.member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = p.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}
fn zip_bytes(parts: &BTreeMap<String, Vec<u8>>) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in parts {
        writer
            .start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(data).unwrap();
    }
    writer.finish().unwrap().into_inner()
}
fn sign(bytes: &[u8], extension: &str) -> Vec<u8> {
    if let Some(format) = cfb_format(extension) {
        let mut edit = cfb::Editor::open(bytes.to_vec(), format, Limits::standard()).unwrap();
        assert!(edit.verify(&Policy::strict()).unwrap().is_empty());
        edit.add(&signer()).unwrap();
        edit.finish().unwrap()
    } else {
        let mut p = opc(bytes);
        assert!(p.signatures().unwrap().is_empty());
        p.sign(&signer()).unwrap();
        PackageWriter::to_bytes(&p).unwrap()
    }
}
fn reports(bytes: &[u8], extension: &str) -> Vec<litchi_sign::Report> {
    if let Some(format) = cfb_format(extension) {
        let mut file =
            litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
                .unwrap();
        cfb::verify(&mut file, format, &Policy::strict())
            .unwrap()
            .into_iter()
            .map(|r| r.details().clone())
            .collect()
    } else {
        opc(bytes)
            .signatures()
            .unwrap()
            .into_iter()
            .map(|r| r.into_parts().1)
            .collect()
    }
}
type Relation = (String, String, String, String);
type Graph = BTreeMap<String, (String, Vec<Relation>)>;
fn graph(p: &OpcPackage) -> Graph {
    fn rels(r: &litchi_opc::Relationships) -> Vec<Relation> {
        let mut result = r
            .iter()
            .map(|r| {
                (
                    r.r_id().to_string(),
                    r.reltype().to_string(),
                    r.target_ref().to_string(),
                    format!("{:?}", r.target_mode()),
                )
            })
            .collect::<Vec<_>>();
        result.sort();
        result
    }
    let mut result = BTreeMap::from([("/".to_string(), (String::new(), rels(p.rels())))]);
    for part in p.iter_parts() {
        result.insert(
            part.partname().as_str().to_string(),
            (part.content_type().to_string(), rels(part.rels())),
        );
    }
    result
}
fn equivalent(a: &[u8], b: &[u8], extension: &str) {
    if extension == "xls" {
        assert!(streams(a) == streams(b));
    } else {
        let a = opc(a);
        let b = opc(b);
        assert_eq!(graph(&a), graph(&b));
        for part in a.iter_parts() {
            assert!(
                a.get_part(part.partname()).unwrap().blob()
                    == b.get_part(part.partname()).unwrap().blob(),
                "changed {}",
                part.partname()
            );
        }
    }
}

#[cfg(all(feature = "doc", feature = "docx", feature = "ppt", feature = "pptx"))]
#[test]
fn native_office_signature_derivatives_verify_integrity_without_claiming_trust() {
    for (name, extension) in [
        ("word/word-text-structure.doc", "doc"),
        ("word/word-text-structure.docx", "docx"),
        ("powerpoint/powerpoint-text-table.ppt", "ppt"),
        ("powerpoint/powerpoint-text-table.pptx", "pptx"),
        ("excel/excel-metadata.xls", "xls"),
        ("excel/excel-metadata.xlsx", "xlsx"),
        ("excel/excel-metadata.xlsb", "xlsb"),
    ] {
        let source = native_file(name);
        let signed = sign(&source, extension);
        let r = reports(&signed, extension);
        assert_eq!(r.len(), 1);
        assert_eq!(
            (
                r[0].integrity(),
                r[0].signature(),
                r[0].coverage(),
                r[0].trust()
            ),
            (
                Status::Valid,
                Status::Valid,
                Coverage::Complete,
                Trust::NotChecked
            )
        );
        assert_eq!(r[0].time(), Some(TIME));
        assert!(!r[0].uses_sha1());
        assert!(r[0].certificates().is_empty());
        if cfb_format(extension).is_some() {
            let before = streams(&source);
            let after = streams(&signed);
            for (name, data) in before {
                assert!(after[&name] == data);
            }
        } else {
            let before = opc(&source);
            let after = opc(&signed);
            for part in before.iter_parts() {
                assert!(
                    before.get_part(part.partname()).unwrap().blob()
                        == after.get_part(part.partname()).unwrap().blob()
                );
            }
        }
        let original = litchi_ops::open_bytes(source.clone()).unwrap();
        let reopened = litchi_ops::open_bytes(signed.clone()).unwrap();
        assert!(reopened.is_signed());
        assert_eq!(reopened.text().unwrap(), original.text().unwrap());
        assert!(reopened.to_bytes().unwrap() == signed);
        assert!(native_file(name) == source);
    }
}

#[test]
fn native_signed_workbooks_require_explicit_unsign_on_outer_and_inner_handles() {
    for extension in ["xls", "xlsx", "xlsb"] {
        let source = native(extension);
        let signed = sign(&source, extension);
        let original = litchi_ops::open_bytes(source.clone()).unwrap();
        let expected = original
            .set("Sheet1!Z100", props! {"value" => 42})
            .unwrap()
            .to_bytes()
            .unwrap();
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join(format!("signed.{extension}"));
        std::fs::write(&path, &signed).unwrap();
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let options = OpenOptions::new().source_mode(mode);
            let file = if mode == SourceMode::Auto {
                litchi_ops::open_with(&path, &options).unwrap()
            } else {
                litchi_ops::open_bytes_with(signed.clone(), &options).unwrap()
            };
            let File::Signed(state) = &file else {
                panic!("{extension}: signature metadata was not detected")
            };
            assert_eq!(
                state.origin(),
                if extension == "xls" {
                    litchi_ops::SignatureOrigin::Cfb
                } else {
                    litchi_ops::SignatureOrigin::Opc
                }
            );
            for f in [&file, state.file()] {
                assert_eq!(
                    f.get("Sheet1!A7")
                        .unwrap()
                        .view()
                        .text()
                        .unwrap()
                        .as_deref(),
                    Some("30")
                );
                assert!(f.edit(|_| Ok(())).unwrap().file().to_bytes().unwrap() == signed);
                assert_eq!(
                    f.set("Sheet1!Z100", props! {"value"=>42})
                        .unwrap_err()
                        .code(),
                    Code::DocumentProtected
                );
                assert!(f.to_bytes().unwrap() == signed);
            }
            let unsigned = state.clone().unsign();
            assert!(!unsigned.is_signed());
            let changed = unsigned
                .set("Sheet1!Z100", props! {"value"=>42})
                .unwrap()
                .to_bytes()
                .unwrap();
            assert!(reports(&changed, extension).is_empty());
            equivalent(&expected, &changed, extension);
            let reopened = litchi_ops::open_bytes(changed).unwrap();
            assert!(!reopened.is_signed());
            assert_eq!(
                reopened
                    .get("Sheet1!Z100")
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("42")
            );
            assert!(file.to_bytes().unwrap() == signed);
        }
        assert!(std::fs::read(path).unwrap() == signed);
        assert!(native(extension) == source);
    }
}

#[test]
fn native_workbook_payload_tampering_separates_signature_math_from_integrity_and_trust() {
    for extension in ["xls", "xlsx", "xlsb"] {
        let source = native(extension);
        let signed = sign(&source, extension);
        let changed = litchi_ops::open_bytes(source.clone())
            .unwrap()
            .set("Sheet1!Z100", props! {"value"=>42})
            .unwrap()
            .to_bytes()
            .unwrap();
        let tampered = if extension == "xls" {
            let mut parts = streams(&signed);
            let replacement = streams(&changed);
            parts.insert(
                vec!["Workbook".into()],
                replacement[&vec!["Workbook".into()]].clone(),
            );
            cfb_bytes(&parts)
        } else {
            let mut parts = members(&signed);
            let replacement = members(&changed);
            let name = if extension == "xlsx" {
                "xl/worksheets/sheet1.xml"
            } else {
                "xl/worksheets/sheet1.bin"
            };
            parts.insert(name.into(), replacement[name].clone());
            zip_bytes(&parts)
        };
        let r = reports(&tampered, extension);
        assert_eq!(r.len(), 1);
        assert_eq!(
            (r[0].integrity(), r[0].signature(), r[0].trust()),
            (Status::Invalid, Status::Valid, Trust::NotChecked)
        );
        // Signed is presence/policy metadata, not automatic verification.
        assert!(litchi_ops::open_bytes(tampered).unwrap().is_signed());
        assert_eq!(reports(&signed, extension)[0].integrity(), Status::Valid);
        assert!(native(extension) == source);
    }
}
