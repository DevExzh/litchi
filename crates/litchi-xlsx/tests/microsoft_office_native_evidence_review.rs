//! Corpus evidence audit only. Absence here does NOT qualify feature support.
#![cfg(feature = "encryption")]
use litchi_xlsx::Package;
use quick_xml::{events::Event, name::ResolveResult, reader::NsReader};
use std::path::Path;

#[derive(Default)]
struct Signals {
    content_parts: usize,
    pivot_data: usize,
    feature_bags: usize,
    pictures: usize,
    shapes: usize,
    pivots: usize,
    controls: usize,
}
fn inspect(package: &litchi_opc::OpcPackage) -> Signals {
    let mut out = Signals::default();
    for part in package.iter_parts() {
        assert!(
            !part
                .content_type()
                .to_ascii_lowercase()
                .contains("featurepropertybag")
        );
        for rel in part.rels().iter() {
            assert!(
                !rel.reltype()
                    .to_ascii_lowercase()
                    .contains("featurepropertybag")
            );
        }
        if !part.partname().as_str().ends_with(".xml")
            && !part.content_type().ends_with("+xml")
            && !matches!(part.content_type(), "application/xml" | "text/xml")
        {
            continue;
        }
        let part = package.get_part(part.partname()).unwrap();
        let bytes = part.blob();
        // DataMashup's custom XML is native UTF-16. Decode it for inspection;
        // the source payload is never rewritten or normalized in the package.
        let xml = if bytes.starts_with(&[0xff, 0xfe]) || bytes.starts_with(&[0xfe, 0xff]) {
            assert_eq!((bytes.len() - 2) % 2, 0);
            let little = bytes[0] == 0xff;
            let units = bytes[2..]
                .chunks_exact(2)
                .map(|b| {
                    if little {
                        u16::from_le_bytes([b[0], b[1]])
                    } else {
                        u16::from_be_bytes([b[0], b[1]])
                    }
                })
                .collect::<Vec<_>>();
            std::borrow::Cow::Owned(String::from_utf16(&units).unwrap())
        } else {
            std::borrow::Cow::Borrowed(std::str::from_utf8(bytes).unwrap())
        };
        let mut reader = NsReader::from_str(&xml);
        loop {
            let (namespace, event) = reader.read_resolved_event().unwrap();
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let uri = match &namespace {
                        ResolveResult::Bound(uri) => uri.as_ref(),
                        ResolveResult::Unbound => b"",
                        ResolveResult::Unknown(_) => {
                            panic!("unbound XML prefix in {}", part.partname())
                        },
                    };
                    let local = e.local_name();
                    let local = local.as_ref();
                    let xdr = uri
                        == b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                        || uri == b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
                    out.content_parts += usize::from(xdr && local == b"contentPart");
                    out.pictures += usize::from(xdr && local == b"pic");
                    out.shapes += usize::from(xdr && local == b"sp");
                    out.pivot_data += usize::from(local == b"pivotTableData");
                    out.feature_bags+=usize::from(uri==b"http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag" || matches!(local,b"xfControls"|b"xfControl"|b"xfComplement"|b"xfComplements"|b"dxfComplement"|b"dxfComplements"));
                    out.pivots += usize::from(local == b"pivotTableDefinition");
                    out.controls += usize::from(local == b"control");
                    for attr in e.attributes() {
                        let attr =
                            attr.unwrap_or_else(|error| panic!("{}: {error}", part.partname()));
                        if attr.key.local_name().as_ref() == b"uri" {
                            assert!(
                                !attr.value.as_ref().eq_ignore_ascii_case(
                                    b"{44433962-1CF7-4059-B4EE-95C3D5FFCF73}"
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
    out
}
#[test]
fn native_excel_corpus_does_not_evidence_content_parts_pivot_data_or_feature_bags() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut files = 0;
    let mut encrypted = 0;
    let mut neighboring = [false; 3];
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("excel/")
            || !matches!(
                Path::new(name).extension().and_then(|s| s.to_str()),
                Some("xlsx" | "xlsm" | "xltx" | "xltm")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let package = if bytes.starts_with(b"PK") {
            Package::from_bytes(&bytes).unwrap()
        } else {
            assert_eq!(name, "excel/excel-encrypted.xlsx");
            encrypted += 1;
            Package::from_bytes_with_password(&bytes, "OfficeFixture42!").unwrap()
        };
        let signals = inspect(&package.into_plain_opc());
        files += 1;
        assert_eq!(
            (
                signals.content_parts,
                signals.pivot_data,
                signals.feature_bags
            ),
            (0, 0, 0),
            "new feature evidence in {name}; update the evidence review instead of counting absence as coverage"
        );
        if name == "excel/excel-conditional-drawings.xlsx" {
            assert_eq!((signals.pictures, signals.shapes), (1, 1));
            neighboring[0] = true;
        }
        if name == "excel/excel-visual-analytics.xlsx" {
            assert!(signals.pivots > 0);
            neighboring[1] = true;
        }
        if name == "excel/excel-form-controls.xlsx" {
            assert!(signals.controls > 0);
            neighboring[2] = true;
        }
    }
    assert_eq!((files, encrypted), (19, 1));
    assert_eq!(neighboring, [true; 3]);
}
