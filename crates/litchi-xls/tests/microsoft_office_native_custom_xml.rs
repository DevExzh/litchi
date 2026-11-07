//! Native BIFF Custom XML derivatives and ordinary cell edits preserve independent CFB owners.
#[path = "../../../tools/test_support/native_custom_xml.rs"]
mod stores;
use litchi_core::sheet::CellValue;
use litchi_xls::{
    Package,
    cell_values::{Reference, Value},
};
#[test]
fn native_xls_custom_xml_lifecycle_and_cell_edits_preserve_opaque_storage_metadata() {
    for name in ["excel/excel-metadata.xls", "excel/excel-recorded-macro.xlt"] {
        let bytes = stores::native(name);
        let seeded = stores::exercise(&bytes);
        let p = Package::from_bytes(&seeded).unwrap();
        let original = Package::from_bytes(&bytes).unwrap();
        assert_eq!(
            p.workbook().worksheet_names().unwrap(),
            original.workbook().worksheet_names().unwrap()
        );
        let at = Reference::new(120, 25).unwrap();
        let mut e = p.edit().unwrap();
        e.insert_cell("Sheet1", at, Value::Number(42.5)).unwrap();
        let c = e.commit().unwrap();
        let out = c.package().to_bytes().unwrap();
        stores::preserved_owned(&seeded, &out);
        let before = stores::streams(&seeded);
        let after = stores::streams(&out);
        for (path, data) in before {
            if path != ["Workbook"] {
                assert!(after.get(&path) == Some(&data), "changed {path:?}");
            }
        }
        let f = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
        c.package().save(f.path()).unwrap();
        let reopened = Package::open(f.path()).unwrap();
        assert_eq!(
            reopened
                .workbook()
                .worksheet("Sheet1")
                .unwrap()
                .unwrap()
                .cell_value(at)
                .unwrap(),
            Some(CellValue::Float(42.5))
        );
        stores::preserved_owned(&seeded, &reopened.to_bytes().unwrap());
        assert!(
            reopened
                .apply(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == seeded
        );
        assert!(p.to_bytes().unwrap() == seeded);
        assert!(stores::native(name) == bytes);
    }
}
