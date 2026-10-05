//! The facade's unified entry point: `litchi::open` and `litchi::File`
//! (ADR 0033, M0a step S7).
//!
//! The root names the `litchi-ops` API; `open` and `open_bytes` refuse Apple
//! iWork inputs with their detected format and a suggestion instead of
//! reporting them as unrecognized; and the reads the facade's format-family
//! types offered (tolerance reports, PPTX slide catalogs and sizes) come
//! from the format handles with the same values.
//!
//! S7 compared those reads with the format-family types directly. M0a step
//! S8c recorded the format-family types' values in `tests/golden/` (pinned
//! to those types by `markdown_file_parity.rs` until S9 removes them), and
//! these tests compare the handles with the recorded values.

#![cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "rtf",
    feature = "odt",
    feature = "ods",
    feature = "odp",
    feature = "pages",
    feature = "keynote",
    feature = "numbers"
))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
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
        feature = "ods",
        feature = "pages",
        feature = "keynote",
        feature = "numbers"
    )),
    allow(
        dead_code,
        unused_imports,
        reason = "helpers are shared by feature-gated tests"
    )
)]

use std::path::{Path, PathBuf};

use litchi::{Code, FileFormat};

fn fixture(relative: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

// ------------------------------------------------------------- root names --

#[cfg(any(feature = "docx", feature = "xlsx", feature = "pptx"))]
#[test]
fn the_root_names_the_unified_api() {
    fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<litchi::File>();
    assert_handle::<litchi::Locked>();
    assert_handle::<litchi::Signed>();
    assert_handle::<litchi::Error>();
    assert_eq!(size_of::<litchi::File>(), 2 * size_of::<usize>());

    let _: fn(&Path) -> litchi::Result<litchi::File> = |path| litchi::open(path);
    let _: fn(Vec<u8>) -> litchi::Result<litchi::File> = |bytes| litchi::open_bytes(bytes);
    let _: fn(&Path, &litchi::OpenOptions) -> litchi::Result<litchi::File> =
        |path, options| litchi::open_with(path, options);
    let _: fn(Vec<u8>, &litchi::OpenOptions) -> litchi::Result<litchi::File> =
        |bytes, options| litchi::open_bytes_with(bytes, options);
    let options = litchi::OpenOptions::new().source_mode(litchi::SourceMode::Eager);
    let _ = options;
    let _: Option<litchi::Family> = None;
    let _: Option<litchi::EncryptionKind> = None;
    let _: Option<litchi::SignatureOrigin> = None;
    let _: Option<litchi::ToleranceReport> = None;
    let _: Option<litchi::error::Source> = None;
    let _: Option<litchi::Code> = None;
    #[cfg(feature = "docx")]
    {
        let _: Option<litchi::DocumentView<'_>> = None;
        let _: Option<litchi::format::docx::Docx> = None;
        let _: Option<litchi::compat::word::Paragraph> = None;
    }
    #[cfg(feature = "pptx")]
    {
        let _: Option<litchi::PresentationView<'_>> = None;
        let _: Option<litchi::format::pptx::SlideDescriptor> = None;
        let _: Option<litchi::compat::slide::Slide> = None;
    }
    #[cfg(feature = "xlsx")]
    let _: Option<litchi::WorkbookView<'_>> = None;
    #[cfg(feature = "doc")]
    let _ = litchi::OpenOptions::new().doc(litchi::DocOptions::new());
    #[cfg(feature = "encryption")]
    let _: Option<litchi::Password> = None;
    #[cfg(all(feature = "docx", any(unix, windows)))]
    {
        let _: fn(litchi::PreparedDetection) -> litchi::Result<litchi::File> =
            litchi::open_prepared;
    }
}

#[cfg(feature = "docx")]
#[test]
fn the_root_error_is_the_unified_error_and_the_core_error_stays_nameable() {
    let error: litchi::Error = litchi::open(fixture("does/not/exist.docx")).unwrap_err();
    assert_eq!(error.code(), Code::FileNotFound);
    // The format crates' shared error stays nameable as `common::Error`, and
    // converts into the unified error with its code and typed source.
    let core = litchi::common::Error::Io(std::io::Error::from(std::io::ErrorKind::NotFound));
    assert!(matches!(core, litchi::common::Error::Io(_)), "{core:?}");
    let unified = litchi::Error::from(core);
    assert_eq!(unified.code(), Code::FileNotFound);
    assert!(
        matches!(
            unified.leaf(),
            litchi::error::Source::Core(litchi::common::Error::Io(_))
        ),
        "{unified:?}"
    );
}

// ------------------------------------------------------------------- iWork --

/// The iWork packages and package directories of each enabled iWork format.
#[allow(
    clippy::vec_init_then_push,
    unused_mut,
    reason = "each push is gated on its own iWork feature"
)]
fn iwork_fixtures() -> Vec<(FileFormat, PathBuf, PathBuf)> {
    let mut fixtures = Vec::new();
    #[cfg(feature = "pages")]
    fixtures.push((
        FileFormat::Pages,
        fixture("iwork/pages/basic.pages"),
        fixture("iwork/directory/pages/basic.pages"),
    ));
    #[cfg(feature = "keynote")]
    fixtures.push((
        FileFormat::Keynote,
        fixture("iwork/keynote/basic.key"),
        fixture("iwork/directory/keynote/basic.key"),
    ));
    #[cfg(feature = "numbers")]
    fixtures.push((
        FileFormat::Numbers,
        fixture("iwork/numbers/basic.numbers"),
        fixture("iwork/directory/numbers/basic.numbers"),
    ));
    fixtures
}

#[track_caller]
fn assert_iwork_refused(error: &litchi::Error, format: FileFormat) {
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), Some(format), "{error:?}");
    let suggestion = error
        .suggestion()
        .expect("an iWork refusal suggests the reader");
    assert!(
        suggestion.contains("litchi::iwork::Document::open"),
        "{suggestion}"
    );
    #[cfg(not(feature = "iwork"))]
    assert!(suggestion.contains("`iwork` cargo feature"), "{suggestion}");
}

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#[test]
fn iwork_packages_are_refused_with_their_format_on_every_route() {
    let eager = litchi::OpenOptions::new().source_mode(litchi::SourceMode::Eager);
    let fixtures = iwork_fixtures();
    assert!(!fixtures.is_empty());
    for (format, package, _) in fixtures {
        assert!(package.is_file(), "{}", package.display());
        let bytes = std::fs::read(&package).unwrap();
        for error in [
            litchi::open(&package).unwrap_err(),
            litchi::open_with(&package, &eager).unwrap_err(),
            litchi::open_bytes(bytes.clone()).unwrap_err(),
            litchi::open_bytes_with(bytes, &eager).unwrap_err(),
        ] {
            assert_iwork_refused(&error, format);
        }
        // The refused input is what the iWork reader opens.
        #[cfg(feature = "iwork")]
        assert_eq!(
            litchi::iwork::Document::open(&package)
                .unwrap()
                .format()
                .name(),
            format!("{format:?}")
        );
    }
}

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#[test]
fn iwork_package_directories_are_refused_with_their_format() {
    for (format, _, directory) in iwork_fixtures() {
        assert!(directory.is_dir(), "{}", directory.display());
        let error = litchi::open(&directory).unwrap_err();
        assert_iwork_refused(&error, format);
        #[cfg(feature = "iwork")]
        litchi::iwork::Document::open(&directory).expect("the iWork reader opens the directory");
    }
}

#[cfg(any(feature = "pages", feature = "keynote", feature = "numbers"))]
#[test]
fn a_directory_that_is_not_an_iwork_package_keeps_its_io_error() {
    let directory = tempfile::tempdir().unwrap();
    let error = litchi::open(directory.path()).unwrap_err();
    assert_eq!(error.code(), Code::IoError, "{error:?}");
    assert_eq!(error.detected(), None);
}

#[cfg(any(feature = "docx", feature = "pages"))]
#[test]
fn input_no_format_claims_is_unrecognized() {
    let error = litchi::open_bytes(b"plain text, not an Office file".to_vec()).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), None, "{error:?}");
    assert_eq!(error.suggestion(), None, "{error:?}");
}

/// A compound file whose host stream names a format this build lacks is
/// reported with that format and the feature to enable, not read as a
/// workbook (M0a step S6; the facade's `Workbook` fell back to XLS).
#[cfg(all(feature = "xls", not(feature = "doc")))]
#[test]
fn a_disabled_format_is_reported_with_its_feature() {
    let error = litchi::open(fixture("ole/doc/watermark.doc")).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), Some(FileFormat::Doc), "{error:?}");
    assert_eq!(error.suggestion(), Some("enable the `doc` cargo feature"));
}

/// The facade's detection functions and `open` agree on iWork: detection
/// names the format, and `open` refuses it with that format.
#[cfg(feature = "pages")]
#[test]
fn detection_and_open_agree_on_iwork() {
    let path = fixture("iwork/pages/basic.pages");
    assert_eq!(litchi::detect_file_format(&path), Some(FileFormat::Pages));
    let error = litchi::open(&path).unwrap_err();
    assert_eq!(error.detected(), litchi::detect_file_format(&path));
}

// ----------------------------------------------------- family view refusals --

#[cfg(all(feature = "docx", feature = "xlsx", feature = "pptx"))]
#[test]
fn a_view_of_another_family_reports_the_detected_format() {
    let docx = litchi::open(fixture("ooxml/docx/documentProperties.docx")).unwrap();
    let error = docx.workbook().unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedFormat);
    assert_eq!(error.detected(), Some(FileFormat::Docx));
    let error = docx.presentation().unwrap_err();
    assert_eq!(error.detected(), Some(FileFormat::Docx));

    let xlsx = litchi::open(fixture("ooxml/xlsx/styles.xlsx")).unwrap();
    assert_eq!(
        xlsx.document().unwrap_err().detected(),
        Some(FileFormat::Xlsx)
    );
}

#[cfg(all(feature = "docx", any(unix, windows)))]
#[test]
fn a_prepared_detection_opens_without_indexing_again() {
    let path = fixture("ooxml/docx/documentProperties.docx");
    let detection = litchi::detect_and_prepare(&path).expect("a DOCX is detected");
    assert!(detection.is_prepared());
    let file = litchi::open_prepared(detection).unwrap();
    assert_eq!(file.format(), Some(FileFormat::Docx));
    assert_eq!(
        file.text().unwrap(),
        litchi::open(&path).unwrap().text().unwrap()
    );
}

// ------------------------------------------------------- tolerance reports --

/// The golden file of `name`.
fn golden(name: &str) -> String {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/golden")
        .join(name);
    std::fs::read_to_string(&path).unwrap_or_else(|error| panic!("{}: {error}", path.display()))
}

/// The rows of a tab-separated golden file, comments skipped.
fn golden_rows(name: &str) -> Vec<Vec<String>> {
    golden(name)
        .lines()
        .filter(|line| !line.starts_with('#') && !line.is_empty())
        .map(|line| line.split('\t').map(str::to_owned).collect())
        .collect()
}

/// Compare the `File` tolerance report of every `family` row of
/// `tolerance_reports.tsv` with the facade type's recorded report, after
/// `prepare` performed the reads the facade test performed.
fn assert_recorded_reports(family: &str, prepare: impl Fn(&litchi::File)) -> usize {
    let mut compared = 0;
    for row in golden_rows("tolerance_reports.tsv") {
        let [row_family, route, relative, expected] = row.as_slice() else {
            panic!("malformed tolerance golden row: {row:?}");
        };
        if row_family != family {
            continue;
        }
        let path = fixture(relative);
        let file = match route.as_str() {
            "path" => litchi::open(&path),
            "bytes" => litchi::open_bytes(std::fs::read(&path).unwrap()),
            other => panic!("unknown route {other:?}"),
        }
        .unwrap_or_else(|error| panic!("{relative} ({route}): {error:?}"));
        prepare(&file);
        let report = file
            .tolerance_report()
            .unwrap_or_else(|error| panic!("{relative} ({route}): File report: {error:?}"));
        assert_eq!(
            &format!("{report:?}"),
            expected,
            "{relative} ({route}): File report differs from the facade type's"
        );
        compared += 1;
    }
    compared
}

/// Every fixture `tolerance_report_facade.rs` reads, through `File`: the
/// same report as the facade type, defects included.
#[cfg(all(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
#[test]
fn document_tolerance_reports_come_from_the_handles() {
    // Read the metadata first, as the facade test does, so deferred
    // core-properties repairs are part of the report.
    let compared = assert_recorded_reports("document", |file| {
        let _ = file.metadata().unwrap();
    });
    // Seven fixtures, by path and by bytes.
    assert_eq!(compared, 14);

    let report = litchi::open(fixture("ole/doc/watermark.doc"))
        .unwrap()
        .tolerance_report()
        .unwrap();
    let litchi::ToleranceReport::Doc(doc) = &report else {
        panic!("expected a DOC report: {report:?}");
    };
    assert!(
        doc.defects()
            .iter()
            .any(|entry| matches!(entry.defect, litchi::doc::Defect::Bookmark(_))),
        "{doc:?}"
    );
}

#[cfg(all(feature = "ppt", feature = "pptx", feature = "odp"))]
#[test]
fn presentation_tolerance_reports_come_from_the_handles() {
    let compared = assert_recorded_reports("presentation", |file| {
        let _ = file.metadata().unwrap();
    });
    assert_eq!(compared, 4);
    let report = litchi::open(fixture("poi/test-data/slideshow/45543.ppt"))
        .unwrap()
        .tolerance_report()
        .unwrap();
    assert!(matches!(report, litchi::ToleranceReport::Ppt(_)));
    assert!(
        report
            .repairs()
            .contains(&"ignored the nonzero reserved or padding bytes"),
        "{report:?}"
    );
}

#[cfg(all(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
#[test]
fn workbook_tolerance_reports_come_from_the_handles() {
    let compared = assert_recorded_reports("workbook", |file| {
        let _ = file.text().unwrap();
        let _ = file.metadata().unwrap();
    });
    assert_eq!(compared, 7);
}

/// Pages, which `open` refuses, keeps no repair report: the untracked report
/// names it.
#[cfg(feature = "pages")]
#[test]
fn pages_reports_stay_untracked() {
    let report = litchi::ToleranceReport::Untracked(FileFormat::Pages);
    assert!(!report.is_tracked());
    assert!(!report.is_clean());
    assert_eq!(report.total(), 0);
    assert_eq!(report.format(), FileFormat::Pages);
}

// ------------------------------------------------ PPTX catalog and size --

/// The slide catalog and size the facade's `Presentation` reported for
/// every `test-data/ooxml/pptx` fixture it opened, read from the `Pptx`
/// handle (a signed package through the file it wraps).
#[cfg(feature = "pptx")]
#[test]
fn pptx_slide_catalog_and_size_come_from_the_handle() {
    let directory = fixture("ooxml/pptx");
    let mut compared = 0;
    for row in golden_rows("pptx_slide_catalog.tsv") {
        let [name, width, height, catalog] = row.as_slice() else {
            panic!("malformed slide catalog golden row: {row:?}");
        };
        let path = directory.join(name);
        let file = litchi::open(&path).unwrap_or_else(|error| panic!("{name}: {error:?}"));
        let file = match &file {
            litchi::File::Signed(signed) => signed.file().clone(),
            other => other.clone(),
        };
        let litchi::File::Pptx(pptx) = &file else {
            panic!("{name}: expected a PPTX, got {file:?}");
        };
        let handle = pptx
            .slide_catalog()
            .unwrap()
            .iter()
            .map(|slide| format!("{}:{}", slide.position(), slide.slide_id()))
            .collect::<Vec<_>>()
            .join(",");
        assert_eq!(&handle, catalog, "{name}");
        let (handle_width, handle_height) = pptx.slide_size().unwrap();
        assert_eq!(&handle_width.to_string(), width, "{name}");
        assert_eq!(&handle_height.to_string(), height, "{name}");
        compared += 1;
    }
    assert!(compared > 0, "no PPTX fixture compared");
}
