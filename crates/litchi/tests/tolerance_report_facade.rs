//! The unified facades expose the owning format crate's repair report.
//!
//! Each format family opens one defect-carrying file through
//! `Document`/`Presentation`/`Workbook::open` and asserts the repair is
//! listed, and one conformant file and asserts a clean report.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
#![cfg_attr(
    not(all(
        feature = "doc",
        feature = "docx",
        feature = "rtf",
        feature = "odt",
        feature = "ppt",
        feature = "pptx",
        feature = "odp",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb",
        feature = "ods"
    )),
    allow(
        dead_code,
        unused_imports,
        reason = "helpers are shared by feature-gated tests"
    )
)]

use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// Rewrite a ZIP package: `patch` may replace a member's bytes, and the
/// `mimetype` member (when present) is written first, compressed with
/// `mimetype_method`. Every other member is deflated.
fn rezip(
    source: &Path,
    mimetype_method: zip::CompressionMethod,
    patch: impl Fn(&str, Vec<u8>) -> Vec<u8>,
) -> Vec<u8> {
    let bytes = std::fs::read(source).expect("fixture");
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).expect("zip fixture");
    let mut members = Vec::new();
    for index in 0..archive.len() {
        let mut member = archive.by_index(index).expect("member");
        if member.is_dir() {
            continue;
        }
        let name = member.name().to_owned();
        let mut data = Vec::new();
        member.read_to_end(&mut data).expect("member bytes");
        let data = patch(&name, data);
        members.push((name, data));
    }
    members.sort_by_key(|(name, _)| name != "mimetype");
    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    for (name, data) in members {
        let method = if name == "mimetype" {
            mimetype_method
        } else {
            zip::CompressionMethod::Deflated
        };
        writer
            .start_file(
                name,
                zip::write::SimpleFileOptions::default().compression_method(method),
            )
            .expect("start member");
        writer.write_all(&data).expect("member bytes");
    }
    writer.finish().expect("finish zip");
    output.into_inner()
}

/// Drop the `xsi:type` producer annotation from the core-properties dates
/// (OPC M4.5), which the reader repairs as `MissingDateType`.
fn strip_core_date_types(name: &str, data: Vec<u8>) -> Vec<u8> {
    if name != "docProps/core.xml" {
        return data;
    }
    let xml = String::from_utf8(data).expect("UTF-8 core properties");
    assert!(xml.contains(r#" xsi:type="dcterms:W3CDTF""#), "{xml}");
    xml.replace(r#" xsi:type="dcterms:W3CDTF""#, "")
        .into_bytes()
}

fn write_temp(directory: &tempfile::TempDir, name: &str, bytes: &[u8]) -> PathBuf {
    let path = directory.path().join(name);
    std::fs::write(&path, bytes).expect("write fixture");
    path
}

fn assert_clean(report: &litchi::ToleranceReport) {
    assert!(report.is_tracked());
    assert!(report.is_clean(), "{report:?}");
    assert_eq!(report.total(), 0);
    assert_eq!(report.unrecorded(), 0);
    assert!(report.repairs().is_empty());
}

fn assert_repaired(report: &litchi::ToleranceReport, repair: &str) {
    assert!(!report.is_clean(), "{report:?}");
    assert!(report.total() >= 1, "{report:?}");
    assert!(
        report.repairs().contains(&repair),
        "expected {repair:?} in {:?}",
        report.repairs()
    );
}

#[cfg(feature = "ooxml-common")]
fn core_properties_defect(
    defect: litchi::ooxml_common::leniency::CorePropertiesDefect,
) -> litchi::ooxml_common::leniency::Defect {
    litchi::ooxml_common::leniency::Defect::CoreProperties(defect)
}

// ---------------------------------------------------------------- documents

#[cfg(feature = "doc")]
#[test]
fn doc_repairs_are_reported_through_the_document_facade() {
    let document = litchi::Document::open(fixture("ole/doc/watermark.doc")).expect("open");
    let report = document.tolerance_report().expect("report");
    assert_eq!(report.format(), litchi::FileFormat::Doc);
    let litchi::ToleranceReport::Doc(doc) = &report else {
        panic!("expected a DOC report: {report:?}");
    };
    assert!(
        doc.defects()
            .iter()
            .any(|entry| matches!(entry.defect, litchi::doc::Defect::Bookmark(_))),
        "{doc:?}"
    );

    let clean = litchi::Document::open(fixture(
        "libreoffice-core/embeddedobj/qa/cppunit/data/insert-file-config.doc",
    ))
    .expect("open");
    assert_clean(&clean.tolerance_report().expect("report"));
}

#[cfg(all(feature = "docx", feature = "ooxml-common"))]
#[test]
fn docx_core_properties_repairs_are_reported_through_the_document_facade() {
    use litchi::ooxml_common::leniency::CorePropertiesDefect;

    let path = fixture(
        "poi/test-data/openxml4j/OPCCompliance_CoreProperties_LimitedXSITypeAttribute_NotPresentFAIL.docx",
    );
    // Filesystem open (deferred parts) and byte open (eager package).
    let opened = litchi::Document::open(&path).expect("open");
    let from_bytes = litchi::Document::from_bytes(std::fs::read(&path).unwrap()).expect("bytes");
    for document in [opened, from_bytes] {
        // Twice: a deferred metadata read is recorded once, not per call.
        let _ = document.metadata().expect("metadata");
        for _ in 0..2 {
            let report = document.tolerance_report().expect("report");
            let litchi::ToleranceReport::Docx(docx) = &report else {
                panic!("expected a DOCX report: {report:?}");
            };
            let missing = docx
                .defects()
                .iter()
                .filter(|entry| {
                    entry.defect
                        == litchi::docx::Defect::Properties(core_properties_defect(
                            CorePropertiesDefect::MissingDateType,
                        ))
                })
                .count();
            assert_eq!(missing, 1, "{docx:?}");
        }
    }

    let clean_path = fixture("libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx");
    let clean = litchi::Document::open(&clean_path).expect("open");
    assert_clean(&clean.tolerance_report().expect("report"));
    let clean = litchi::Document::from_bytes(std::fs::read(&clean_path).unwrap()).expect("bytes");
    assert_clean(&clean.tolerance_report().expect("report"));
}

#[cfg(feature = "rtf")]
#[test]
fn rtf_repairs_are_reported_through_the_document_facade() {
    let document = litchi::Document::open(fixture(
        "libreoffice-core/sw/qa/extras/rtfexport/data/tdf108955.rtf",
    ))
    .expect("open");
    let report = document.tolerance_report().expect("report");
    assert!(matches!(report, litchi::ToleranceReport::Rtf(_)));
    assert_repaired(
        &report,
        "ignored the content after the root group's closing brace",
    );

    let clean = litchi::Document::open(fixture(
        "libreoffice-core/sw/qa/extras/layout/data/A020-min.rtf",
    ))
    .expect("open");
    assert_clean(&clean.tolerance_report().expect("report"));
}

#[cfg(feature = "odf-common")]
fn assert_mimetype_deflated(report: &litchi::odf_common::ToleranceReport) {
    assert!(
        report.defects().iter().any(|entry| entry.defect
            == litchi::odf_common::Defect::Package(
                litchi::odf_common::PackageDefect::MimetypeDeflated
            )),
        "{report:?}"
    );
}

#[cfg(feature = "odt")]
#[test]
fn odt_repairs_are_reported_through_the_document_facade() {
    let source = fixture("odfpy/tests/examples/simpletable.odt");
    let directory = tempfile::tempdir().unwrap();
    let deflated = write_temp(
        &directory,
        "deflated.odt",
        &rezip(&source, zip::CompressionMethod::Deflated, |_, data| data),
    );
    let document = litchi::Document::open(&deflated).expect("open");
    let report = document.tolerance_report().expect("report");
    let litchi::ToleranceReport::Odt(odt) = &report else {
        panic!("expected an ODT report: {report:?}");
    };
    assert_mimetype_deflated(odt);

    assert_clean(
        &litchi::Document::open(&source)
            .expect("open")
            .tolerance_report()
            .expect("report"),
    );
}

#[cfg(feature = "pages")]
#[test]
fn pages_reports_are_untracked() {
    let report = litchi::ToleranceReport::Untracked(litchi::FileFormat::Pages);
    assert!(!report.is_tracked());
    assert!(!report.is_clean());
    assert_eq!(report.total(), 0);
}

// ------------------------------------------------------------ presentations

#[cfg(feature = "ppt")]
#[test]
fn ppt_repairs_are_reported_through_the_presentation_facade() {
    let presentation =
        litchi::Presentation::open(fixture("poi/test-data/slideshow/45543.ppt")).expect("open");
    let report = presentation.tolerance_report().expect("report");
    assert!(matches!(report, litchi::ToleranceReport::Ppt(_)));
    assert_repaired(&report, "ignored the nonzero reserved or padding bytes");

    let clean = litchi::Presentation::open(fixture("ole/ppt/empty.ppt")).expect("open");
    assert_clean(&clean.tolerance_report().expect("report"));
}

#[cfg(all(feature = "pptx", feature = "ooxml-common"))]
#[test]
fn pptx_core_properties_repairs_are_reported_through_the_presentation_facade() {
    use litchi::ooxml_common::leniency::CorePropertiesDefect;

    let source = fixture("libreoffice-core/oox/qa/unit/data/shape-text-alignment.pptx");
    let patched = rezip(
        &source,
        zip::CompressionMethod::Deflated,
        strip_core_date_types,
    );
    let directory = tempfile::tempdir().unwrap();
    let path = write_temp(&directory, "missing-date-type.pptx", &patched);
    let expected = litchi::pptx::Defect::Properties(core_properties_defect(
        CorePropertiesDefect::MissingDateType,
    ));
    for presentation in [
        litchi::Presentation::open(&path).expect("open"),
        litchi::Presentation::from_bytes(patched.clone()).expect("bytes"),
    ] {
        let _ = presentation.metadata().expect("metadata");
        let report = presentation.tolerance_report().expect("report");
        let litchi::ToleranceReport::Pptx(pptx) = &report else {
            panic!("expected a PPTX report: {report:?}");
        };
        let missing = pptx
            .defects()
            .iter()
            .filter(|entry| entry.defect == expected)
            .count();
        assert_eq!(missing, 1, "{pptx:?}");
    }

    assert_clean(
        &litchi::Presentation::open(&source)
            .expect("open")
            .tolerance_report()
            .expect("report"),
    );
    assert_clean(
        &litchi::Presentation::from_bytes(std::fs::read(&source).unwrap())
            .expect("bytes")
            .tolerance_report()
            .expect("report"),
    );
}

#[cfg(feature = "odp")]
#[test]
fn odp_repairs_are_reported_through_the_presentation_facade() {
    let source = fixture("libreoffice-core/sd/qa/unit/data/odp/ole_icon.odp");
    let directory = tempfile::tempdir().unwrap();
    let deflated = write_temp(
        &directory,
        "deflated.odp",
        &rezip(&source, zip::CompressionMethod::Deflated, |_, data| data),
    );
    let report = litchi::Presentation::open(&deflated)
        .expect("open")
        .tolerance_report()
        .expect("report");
    let litchi::ToleranceReport::Odp(odp) = &report else {
        panic!("expected an ODP report: {report:?}");
    };
    assert_mimetype_deflated(odp);

    assert_clean(
        &litchi::Presentation::open(&source)
            .expect("open")
            .tolerance_report()
            .expect("report"),
    );
}

// ---------------------------------------------------------------- workbooks

#[cfg(feature = "xls")]
#[test]
fn xls_repairs_are_reported_through_the_workbook_facade() {
    let path = fixture("ole/xls/WithCustomViews.xls");
    for workbook in [
        litchi::Workbook::open(&path).expect("open"),
        litchi::Workbook::from_bytes(std::fs::read(&path).unwrap()).expect("bytes"),
    ] {
        let report = workbook.tolerance_report().expect("report");
        assert!(matches!(report, litchi::ToleranceReport::Xls(_)));
        assert_repaired(&report, "ignored the nonzero reserved or padding bytes");
    }

    let clean = fixture("libreoffice-core/chart2/qa/extras/data/xls/chart.xls");
    assert_clean(
        &litchi::Workbook::open(&clean)
            .expect("open")
            .tolerance_report()
            .expect("report"),
    );
    assert_clean(
        &litchi::Workbook::from_bytes(std::fs::read(&clean).unwrap())
            .expect("bytes")
            .tolerance_report()
            .expect("report"),
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_core_properties_repairs_are_reported_through_the_workbook_facade() {
    use litchi::ooxml_common::leniency::CorePropertiesDefect;

    fn properties_count(report: &litchi::ToleranceReport, defect: CorePropertiesDefect) -> usize {
        let litchi::ToleranceReport::Xlsx(xlsx) = report else {
            panic!("expected an XLSX report: {report:?}");
        };
        let expected = litchi::xlsx::Defect::Properties(core_properties_defect(defect));
        xlsx.defects()
            .iter()
            .filter(|entry| entry.defect == expected)
            .count()
    }

    // A core-properties part no relationship selects.
    let orphaned = fixture("ooxml/xlsx/universal-content.xlsx");
    for workbook in [
        litchi::Workbook::open(&orphaned).expect("open"),
        litchi::Workbook::from_bytes(std::fs::read(&orphaned).unwrap()).expect("bytes"),
    ] {
        let report = workbook.tolerance_report().expect("report");
        assert_eq!(
            properties_count(&report, CorePropertiesDefect::OrphanedPart),
            1,
            "{report:?}"
        );
    }

    // A core-properties body deviation (dates without `xsi:type`).
    let clean = fixture("libreoffice-core/chart2/qa/extras/data/xlsx/boxWhisker.xlsx");
    let patched = rezip(
        &clean,
        zip::CompressionMethod::Deflated,
        strip_core_date_types,
    );
    let directory = tempfile::tempdir().unwrap();
    let path = write_temp(&directory, "missing-date-type.xlsx", &patched);
    for workbook in [
        litchi::Workbook::open(&path).expect("open"),
        litchi::Workbook::from_bytes(patched.clone()).expect("bytes"),
    ] {
        assert!(workbook.metadata().expect("metadata").created.is_some());
        let report = workbook.tolerance_report().expect("report");
        assert!(
            properties_count(&report, CorePropertiesDefect::MissingDateType) >= 1,
            "{report:?}"
        );
    }

    for workbook in [
        litchi::Workbook::open(&clean).expect("open"),
        litchi::Workbook::from_bytes(std::fs::read(&clean).unwrap()).expect("bytes"),
    ] {
        let _ = workbook.text().expect("text");
        assert_clean(&workbook.tolerance_report().expect("report"));
    }
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_core_properties_repairs_are_reported_through_the_workbook_facade() {
    use litchi::ooxml_common::leniency::CorePropertiesDefect;

    let path = fixture("ooxml/xlsb/universal-content.xlsb");
    let expected = litchi::xlsb::Defect::Properties(core_properties_defect(
        CorePropertiesDefect::DuplicatePart,
    ));
    for workbook in [
        litchi::Workbook::open(&path).expect("open"),
        litchi::Workbook::from_bytes(std::fs::read(&path).unwrap()).expect("bytes"),
    ] {
        let report = workbook.tolerance_report().expect("report");
        let litchi::ToleranceReport::Xlsb(xlsb) = &report else {
            panic!("expected an XLSB report: {report:?}");
        };
        let duplicates = xlsb
            .defects()
            .iter()
            .filter(|entry| entry.defect() == expected)
            .count();
        assert_eq!(duplicates, 1, "{xlsb:?}");
        let entry = xlsb
            .defects()
            .iter()
            .find(|entry| entry.defect() == expected)
            .unwrap();
        assert!(entry.part().is_some(), "{entry:?}");
    }

    let clean = fixture("libreoffice-core/sc/qa/unit/data/xlsb/tdf108017_calcProtection.xlsb");
    for workbook in [
        litchi::Workbook::open(&clean).expect("open"),
        litchi::Workbook::from_bytes(std::fs::read(&clean).unwrap()).expect("bytes"),
    ] {
        let _ = workbook.text().expect("text");
        assert_clean(&workbook.tolerance_report().expect("report"));
    }
}

#[cfg(feature = "ods")]
#[test]
fn ods_repairs_are_reported_through_the_workbook_facade() {
    let source = fixture("libreoffice-core/sc/qa/extras/testdocuments/ScDatabaseRangeObj.ods");
    let deflated = rezip(&source, zip::CompressionMethod::Deflated, |_, data| data);
    let directory = tempfile::tempdir().unwrap();
    let path = write_temp(&directory, "deflated.ods", &deflated);
    for workbook in [
        litchi::Workbook::open(&path).expect("open"),
        litchi::Workbook::from_bytes(deflated.clone()).expect("bytes"),
    ] {
        let report = workbook.tolerance_report().expect("report");
        let litchi::ToleranceReport::Ods(ods) = &report else {
            panic!("expected an ODS report: {report:?}");
        };
        assert_mimetype_deflated(ods);
    }

    assert_clean(
        &litchi::Workbook::open(&source)
            .expect("open")
            .tolerance_report()
            .expect("report"),
    );
}
