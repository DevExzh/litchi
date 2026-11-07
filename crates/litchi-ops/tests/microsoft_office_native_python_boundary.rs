//! Python parts and PY formulas are explicit library derivatives, never executed.
#![cfg(all(feature = "xlsx", feature = "eval"))]
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceLimits, TargetMode, part::BlobPart};
use litchi_ops::{Code, OpenOptions, SourceMode, calc, props};
use std::{collections::BTreeMap, path::Path};
const PYTHON: &[u8]=br#"<?xml version="1.0"?><python xmlns="http://schemas.microsoft.com/office/spreadsheetml/2023/python"><environmentDefinition id="{11111111-2222-3333-4444-555555555555}"><initialization userModified="false"><code>fixture_marker = 62</code></initialization></environmentDefinition><pythonScripts><pythonScript><code>40 + 2</code></pythonScript></pythonScripts></python>"#;
const NAME: &str = "xl/python/python.xml";
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
#[test]
fn native_python_part_derivatives_are_preserved_while_uncached_py_evaluation_refuses() {
    for name in ["excel-metadata.xlsx", "excel-recorded-macro.xlsm"] {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name);
        let native = std::fs::read(&path).unwrap();
        let mut opc = OpcPackage::from_bytes(&native, SourceLimits::default()).unwrap();
        assert!(
            !opc.iter_parts()
                .any(|p| p.content_type() == "application/vnd.ms-excel.python+xml")
        );
        opc.try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{NAME}")).unwrap(),
            "application/vnd.ms-excel.python+xml".into(),
            PYTHON.to_vec(),
        )))
        .unwrap();
        opc.get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "http://schemas.microsoft.com/office/2023/09/relationships/Python".into(),
                "python/python.xml".into(),
                "rIdNativePython062".into(),
                TargetMode::Internal,
            )
            .unwrap();
        let derivative = PackageWriter::to_bytes(&opc).unwrap();
        let data = members(&derivative);
        for (n, b) in members(&native) {
            if !matches!(
                n.as_str(),
                "[Content_Types].xml" | "xl/_rels/workbook.xml.rels"
            ) {
                assert!(data.get(&n) == Some(&b), "changed {n}");
            }
        }
        for mode in [SourceMode::Auto, SourceMode::Eager] {
            let file = tempfile::NamedTempFile::with_suffix(if name.ends_with("xlsm") {
                ".xlsm"
            } else {
                ".xlsx"
            })
            .unwrap();
            std::fs::write(file.path(), &derivative).unwrap();
            let p =
                litchi_ops::open_with(file.path(), &OpenOptions::new().source_mode(mode)).unwrap();
            assert!(p.to_bytes().unwrap() == derivative);
            let changed = p
                .set("Sheet1!Z100", props! {"formula"=>"_xlfn.PY(0,0)"})
                .unwrap();
            let before = changed.to_bytes().unwrap();
            assert_eq!(members(&before)[NAME], PYTHON);
            if name.ends_with("xlsm") {
                assert!(data.contains_key("xl/vbaProject.bin"));
                assert_eq!(
                    members(&before)["xl/vbaProject.bin"],
                    data["xl/vbaProject.bin"]
                );
            }

            let error = changed
                .evaluate("Sheet1!Z100", &calc::Options::new().recalculate(true))
                .unwrap_err();
            assert_eq!(error.code(), Code::UnsupportedCapability);
            assert!(matches!(
                error.leaf(),
                litchi_ops::error::Source::Calc(calc::Error::Unsupported)
            ));
            assert!(changed.to_bytes().unwrap() == before);
            let report = changed.recalculate(&calc::Options::new()).unwrap();
            let formula = report
                .formulas()
                .iter()
                .find(|f| (f.sheet(), f.row(), f.column()) == (0, 99, 25))
                .unwrap();
            assert_eq!(formula.status(), calc::Status::Unsupported);
            let other = changed.set("Sheet1!Z101", props! {"value"=>62}).unwrap();
            let saved = other.to_bytes().unwrap();
            assert_eq!(members(&saved)[NAME], PYTHON);
            if name.ends_with("xlsm") {
                assert_eq!(
                    members(&saved)["xl/vbaProject.bin"],
                    data["xl/vbaProject.bin"]
                );
            }

            assert_eq!(
                members(&saved)["xl/_rels/workbook.xml.rels"],
                members(&before)["xl/_rels/workbook.xml.rels"]
            );
            let opened = litchi_ops::open_bytes(saved.clone()).unwrap();
            assert!(opened.to_bytes().unwrap() == saved);
            assert!(p.to_bytes().unwrap() == derivative);
        }
        assert!(std::fs::read(path).unwrap() == native);
    }
}
