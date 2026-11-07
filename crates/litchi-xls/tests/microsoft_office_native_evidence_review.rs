//! Evidence-only audit of absent WebPub and XCB owners in the native XLS corpus.
use litchi_biff::{Limits, Records};
use litchi_xls::{OpenOptions, Package};
use std::io::Cursor;
#[test]
fn native_xls_corpus_has_no_webpub_records_or_office_toolbar_streams() {
    let root =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut files = 0;
    let mut encrypted = 0;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !name.starts_with("excel/")
            || !matches!(
                std::path::Path::new(name)
                    .extension()
                    .and_then(|s| s.to_str()),
                Some("xls" | "xlt")
            )
        {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        let mut ole =
            litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
                .unwrap();
        assert!(
            ole.list_streams()
                .iter()
                .all(|p| p.iter().all(|n| !n.eq_ignore_ascii_case("XCB"))),
            "{name}: new toolbar evidence"
        );
        let stream = ole.open_stream(&["Workbook"]).unwrap();
        let records = Records::new(&stream, Limits::default())
            .unwrap()
            .collect::<Result<Vec<_>, _>>()
            .unwrap();
        assert!(
            records.iter().all(|r| r.kind().get() != 0x0801),
            "{name}: new WebPub evidence"
        );
        assert!(records.iter().any(|r| r.kind().get() == 0x809));
        let protected = name == "excel/excel-encrypted.xls";
        let p = if protected {
            encrypted += 1;
            Package::from_bytes_with_options(
                &bytes,
                OpenOptions::new().password("OfficeFixture42!"),
            )
            .unwrap()
        } else {
            Package::from_bytes(&bytes).unwrap()
        };
        assert!(!p.workbook().worksheet_names().unwrap().is_empty());
        assert!(p.reader_ref().unwrap().web_publications().is_empty());
        if protected {
            assert!(p.to_encrypted_bytes().unwrap() == bytes);
        } else {
            assert!(p.to_bytes().unwrap() == bytes);
        }
        files += 1;
    }
    assert_eq!((files, encrypted), (17, 1));
}
