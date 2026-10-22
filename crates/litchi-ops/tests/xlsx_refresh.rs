//! Workbook refresh (record 0279): recalculation through the `eval`
//! engine, and formula cache write-back that keeps the formula markup.
#![cfg(all(feature = "xlsx", feature = "eval"))]

use std::io::{Cursor, Read as _};

use litchi_core::FileFormat;
use litchi_ops::refresh::{Options, Status};
use litchi_ops::{Code, props};

fn sheet_xml(file: &litchi_ops::File) -> String {
    let bytes = file.to_bytes().unwrap();
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut member = archive.by_name("xl/worksheets/sheet1.xml").unwrap();
    let mut contents = String::new();
    member.read_to_string(&mut contents).unwrap();
    contents
}

#[test]
fn stale_caches_are_reported_then_written_once() {
    let file = litchi_ops::create(FileFormat::Xlsx)
        .unwrap()
        .set("/Sheet1/A1", props! { "value" => 20 })
        .unwrap()
        .set("/Sheet1/A2", props! { "formula" => "A1*2" })
        .unwrap()
        .set(
            "/Sheet1/A3",
            props! { "formula" => "IF(A1>10,\"big\",\"small\")" },
        )
        .unwrap()
        .set("/Sheet1/A4", props! { "formula" => "A1>5" })
        .unwrap()
        .set("/Sheet1/A5", props! { "formula" => "1/0" })
        .unwrap();
    let report = file.refresh(&Options::new()).unwrap();
    assert!(report.patch().is_noop());
    assert_eq!(report.value().count(Status::Stale), 4);
    assert_eq!(report.value().fields()[0].location(), Some("Sheet1!A2"));
    let written = file.refresh(&Options::new().write_cache(true)).unwrap();
    assert_eq!(written.value().refreshed(), 4);
    assert_eq!(written.diagnostics().commits(), 1);
    let xml = sheet_xml(written.file());
    for expected in [
        "<f>A1*2</f><v>40</v>",
        r#"t="str"><f>IF(A1&gt;10,&quot;big&quot;,&quot;small&quot;)</f><v>big</v>"#,
        r#"t="b"><f>A1&gt;5</f><v>1</v>"#,
        r#"t="e"><f>1/0</f><v>#DIV/0!</v>"#,
    ] {
        assert!(xml.contains(expected), "{expected}: {xml}");
    }
    let reopened = litchi_ops::open_bytes(written.file().to_bytes().unwrap()).unwrap();
    let again = reopened.refresh(&Options::new().write_cache(true)).unwrap();
    assert!(again.patch().is_noop());
    assert_eq!(again.value().count(Status::Unchanged), 4);
}

#[test]
fn write_back_is_xlsx_only() {
    #[cfg(feature = "docx")]
    {
        let docx = litchi_ops::create(FileFormat::Docx).unwrap();
        assert!(
            docx.refresh(&Options::new().write_cache(true))
                .unwrap()
                .patch()
                .is_noop()
        );
    }
    let error = litchi_ops::create(FileFormat::Xlsx)
        .unwrap()
        .refresh_with_options(
            &Options::new(),
            litchi_ops::EditOptions::new()
                .semantic_limits(litchi_ops::transaction::SemanticLimits::default()),
        )
        .map(|commit| commit.value().fields().len());
    assert_eq!(error.unwrap(), 0);
    let _ = Code::UnsupportedCapability;
}

/// Record 0295: XLSB and XLS write stale results back as the formula
/// records' cached values, once.
#[cfg(all(feature = "xlsb", feature = "xls"))]
#[test]
fn legacy_and_binary_workbooks_write_caches_back() {
    for format in [FileFormat::Xlsb, FileFormat::Xls] {
        let file = litchi_ops::create(format)
            .unwrap()
            .set("/Sheet1/A1", props! { "value" => 20 })
            .unwrap()
            .set("/Sheet1/A2", props! { "formula" => "A1*2" })
            .unwrap()
            .set("/Sheet1/A3", props! { "formula" => "A1+1" })
            .unwrap();
        let report = file.refresh(&Options::new()).unwrap();
        assert!(report.patch().is_noop(), "{format:?}");
        let stale = report.value().count(Status::Stale);
        let written = file
            .refresh(&Options::new().write_cache(true))
            .unwrap_or_else(|error| panic!("{format:?}: {error}"));
        assert_eq!(
            written.value().refreshed(),
            stale,
            "{format:?}: {:?}",
            written
                .value()
                .fields()
                .iter()
                .map(|field| (field.status(), field.reason()))
                .collect::<Vec<_>>()
        );
        if stale > 0 {
            assert!(written.diagnostics().changed(), "{format:?}");
        }
        let reopened = litchi_ops::open_bytes(written.file().to_bytes().unwrap()).unwrap();
        let again = reopened.refresh(&Options::new().write_cache(true)).unwrap();
        assert!(again.patch().is_noop(), "{format:?}");
        assert_eq!(again.value().count(Status::Unchanged), 2, "{format:?}");
        let restored = written.patch().inverse().apply(written.file()).unwrap();
        assert_eq!(
            restored.to_bytes().unwrap(),
            file.to_bytes().unwrap(),
            "{format:?}"
        );
    }
}
