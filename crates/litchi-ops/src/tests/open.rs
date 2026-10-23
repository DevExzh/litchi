//! The single-pass openers: the `File` layout, one container index per
//! open, wrong-family refusals, read limits, cancellation, owner selection
//! and the lowest-precedence ZIP tail.

use crate::{Error, File, Locked, Signed};

// `File`, every handle, `Locked`, `Signed` and `Error` are `Send + Sync +
// Clone`; a `File` is a tag and one pointer, and the others one pointer.
const _: () = {
    const fn handle<T: Send + Sync + Clone>() {}
    handle::<File>();
    handle::<Locked>();
    handle::<Signed>();
    handle::<Error>();
    #[cfg(feature = "docx")]
    handle::<crate::format::docx::Docx>();
    #[cfg(feature = "xlsx")]
    handle::<crate::format::xlsx::Xlsx>();
    #[cfg(feature = "pptx")]
    handle::<crate::format::pptx::Pptx>();
    #[cfg(feature = "xlsb")]
    handle::<crate::format::xlsb::Xlsb>();
    #[cfg(feature = "doc")]
    handle::<crate::format::doc::Doc>();
    #[cfg(feature = "xls")]
    handle::<crate::format::xls::Xls>();
    #[cfg(feature = "ppt")]
    handle::<crate::format::ppt::Ppt>();
    #[cfg(feature = "rtf")]
    handle::<crate::format::rtf::Rtf>();
    #[cfg(feature = "odt")]
    handle::<crate::format::odt::Odt>();
    #[cfg(feature = "ods")]
    handle::<crate::format::ods::Ods>();
    #[cfg(feature = "odp")]
    handle::<crate::format::odp::Odp>();
    assert!(size_of::<File>() == 2 * size_of::<usize>());
    assert!(size_of::<Locked>() == size_of::<usize>());
    assert!(size_of::<Signed>() == size_of::<usize>());
    assert!(size_of::<Error>() == size_of::<usize>());
};
#[cfg(target_pointer_width = "64")]
const _: () = assert!(size_of::<File>() == 16);

/// One fixture per enabled format: its path under `test-data`, its format
/// and its family.
#[cfg(any(
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
))]
fn fixtures() -> Vec<(&'static str, litchi_core::FileFormat)> {
    use litchi_core::FileFormat;

    vec![
        #[cfg(feature = "docx")]
        ("ooxml/docx/comment.docx", FileFormat::Docx),
        #[cfg(feature = "doc")]
        ("ole/doc/NoHeadFoot.doc", FileFormat::Doc),
        #[cfg(feature = "rtf")]
        ("rtf/testUnicode.rtf", FileFormat::Rtf),
        #[cfg(feature = "odt")]
        ("odf/corpus/writer-paragraph-styles.odt", FileFormat::Odt),
        #[cfg(feature = "pptx")]
        ("ooxml/pptx/shapes.pptx", FileFormat::Pptx),
        #[cfg(feature = "ppt")]
        ("ole/ppt/text_shapes.ppt", FileFormat::Ppt),
        #[cfg(feature = "odp")]
        ("odf/corpus/impress-basic.odp", FileFormat::Odp),
        #[cfg(feature = "xlsx")]
        ("ooxml/xlsx/styles.xlsx", FileFormat::Xlsx),
        #[cfg(feature = "xls")]
        ("ole/xls/Simple.xls", FileFormat::Xls),
        #[cfg(feature = "xlsb")]
        ("ooxml/xlsb/Simple.xlsb", FileFormat::Xlsb),
        #[cfg(feature = "ods")]
        ("odf/ods/grid-padding-package.ods", FileFormat::Ods),
    ]
}

/// Wrong-family refusals: every family view of every enabled format, and
/// the unrecognized input (`unexpected_format.rs` for the facade's openers).
#[cfg(any(
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
))]
mod families {
    use litchi_core::FileFormat;

    use super::super::test_data;
    use crate::{Code, Family, File};

    /// Each family view this build has, applied to `file`.
    fn views(file: &File) -> Vec<(Family, crate::Result<()>)> {
        vec![
            #[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
            (Family::Document, file.document().map(drop)),
            #[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
            (Family::Presentation, file.presentation().map(drop)),
            #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
            (Family::Workbook, file.workbook().map(drop)),
        ]
    }

    #[track_caller]
    fn assert_views(file: &File, format: FileFormat) {
        assert_eq!(file.format(), Some(format));
        assert_eq!(file.family(), Some(Family::of(format)));
        for (family, result) in views(file) {
            if family == Family::of(format) {
                result.unwrap_or_else(|error| panic!("{format:?} {family:?} view: {error:?}"));
            } else {
                let error = result.expect_err("a view of another family is refused");
                assert_eq!(
                    error.code(),
                    Code::UnsupportedFormat,
                    "{format:?} {family:?}"
                );
                assert_eq!(error.detected(), Some(format), "{format:?} {family:?}");
            }
        }
    }

    #[test]
    fn every_view_refuses_every_other_family_with_the_detected_format() {
        for (relative, format) in super::fixtures() {
            let path = test_data(relative);
            let file = crate::open(&path).unwrap_or_else(|error| panic!("{relative}: {error:?}"));
            assert_views(&file, format);
            let bytes = std::fs::read(&path).unwrap();
            let file =
                crate::open_bytes(bytes).unwrap_or_else(|error| panic!("{relative}: {error:?}"));
            assert_views(&file, format);
        }
    }

    /// The facade's `the_workbook_opener_separates_every_wrong_format_it_can_name`,
    /// generalized: every wrong format a view refuses is named distinctly.
    #[test]
    fn refusals_name_each_wrong_format_distinctly() {
        let mut named = Vec::new();
        for (relative, format) in super::fixtures() {
            let file = crate::open(test_data(relative)).unwrap();
            if views(&file)
                .into_iter()
                .any(|(_, result)| result.is_err_and(|error| error.detected() == Some(format)))
            {
                named.push(format);
            }
        }
        let mut unique = named.clone();
        unique.sort_by_key(|format| format!("{format:?}"));
        unique.dedup();
        assert_eq!(unique.len(), named.len());
    }

    #[test]
    fn a_non_office_file_reports_no_detected_format() {
        let path = test_data("images/png/lena.png");
        for error in [
            crate::open(&path).expect_err("a PNG is not an Office file"),
            crate::open_bytes(std::fs::read(&path).unwrap())
                .expect_err("a PNG is not an Office file"),
        ] {
            assert_eq!(error.code(), Code::UnsupportedFormat);
            assert_eq!(error.detected(), None);
            assert_eq!(error.format(), None);
        }
    }

    #[test]
    fn short_and_empty_inputs_are_not_office_files() {
        for bytes in [Vec::new(), b"PK".to_vec(), b"not an office file".to_vec()] {
            let error = crate::open_bytes(bytes).expect_err("not an Office file");
            assert_eq!(error.code(), Code::UnsupportedFormat);
            assert_eq!(error.detected(), None);
        }
    }

    #[test]
    fn a_missing_path_is_file_not_found() {
        let error = crate::open(test_data("no/such/file.docx")).expect_err("missing");
        assert_eq!(error.code(), Code::FileNotFound);
    }

    /// The facade's `non_powerpoint_ole_keeps_the_existing_facade_fallback`
    /// and `malformed_ole_keeps_the_existing_facade_fallback`: the path and
    /// byte routes agree on a compound file.
    #[cfg(any(unix, windows))]
    #[test]
    fn path_and_byte_routes_agree_on_compound_files() {
        let mut malformed = litchi_core::detection::utils::OLE2_SIGNATURE.to_vec();
        malformed.resize(512, 0);
        let temporary = super::super::temporary_with(&malformed, "");
        let from_path = crate::open(temporary.path()).expect_err("malformed compound file");
        let from_bytes = crate::open_bytes(malformed).expect_err("malformed compound file");
        assert_eq!(from_path.code(), Code::UnsupportedFormat);
        assert_eq!(from_path.code(), from_bytes.code());
        assert_eq!(from_path.detected(), None);
        assert_eq!(from_path.detected(), from_bytes.detected());
    }
}

/// One OPC container index per OOXML open (record 0003 §2.5), counted by
/// the detection probe counter.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
mod one_index {
    use super::super::test_data;
    use crate::detect::{opc_probe_count, reset_opc_probe_count};
    use crate::{File, OpenOptions, SourceMode};

    fn ooxml_fixtures() -> Vec<&'static str> {
        vec![
            #[cfg(feature = "docx")]
            "ooxml/docx/FancyFoot.docx",
            #[cfg(feature = "pptx")]
            "ooxml/pptx/shapes.pptx",
            #[cfg(feature = "xlsx")]
            "ooxml/xlsx/styles.xlsx",
            #[cfg(feature = "xlsb")]
            "ooxml/xlsb/Simple.xlsb",
        ]
    }

    #[track_caller]
    fn assert_one_index(open: impl FnOnce() -> crate::Result<File>) -> File {
        reset_opc_probe_count();
        let file = open().expect("open OOXML fixture");
        assert_eq!(opc_probe_count(), 1, "{file:?}");
        // Reads use the package indexed at open.
        file.text().expect("read the opened package");
        assert_eq!(opc_probe_count(), 1, "{file:?}");
        file
    }

    #[test]
    fn every_ooxml_route_indexes_the_container_once() {
        let eager = OpenOptions::new().source_mode(SourceMode::Eager);
        for relative in ooxml_fixtures() {
            let path = test_data(relative);
            let bytes = std::fs::read(&path).unwrap();
            assert_one_index(|| crate::open(&path));
            assert_one_index(|| crate::open_with(&path, &eager));
            assert_one_index(|| crate::open_bytes(bytes.clone()));
            assert_one_index(|| crate::open_bytes_with(bytes.clone(), &eager));
        }
    }

    #[cfg(any(unix, windows))]
    #[test]
    fn open_prepared_adopts_the_index_detection_built() {
        for relative in ooxml_fixtures() {
            reset_opc_probe_count();
            let detection = crate::detect::detect_and_prepare(
                test_data(relative),
                &crate::detect::SourceLimits::default(),
            )
            .expect("detect the fixture");
            assert!(detection.is_prepared(), "{relative}");
            let detected = opc_probe_count();
            let format = detection.format();
            let file = crate::open_prepared(detection).expect("adopt the prepared package");
            assert_eq!(file.format(), Some(format));
            assert_eq!(opc_probe_count(), detected, "{relative}: no second index");
        }
    }

    #[cfg(all(any(unix, windows), feature = "rtf"))]
    #[test]
    fn open_prepared_refuses_a_detection_without_a_package() {
        let detection = crate::detect::detect_and_prepare(
            test_data("rtf/testUnicode.rtf"),
            &crate::detect::SourceLimits::default(),
        )
        .expect("detect the RTF fixture");
        assert!(!detection.is_prepared());
        let error = crate::open_prepared(detection).expect_err("nothing was prepared");
        assert_eq!(error.code(), crate::Code::InvalidArgument);
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Rtf));
    }

    /// A ZIP that is not OOXML is probed once, then handed to the
    /// lower-precedence candidates without a second probe.
    #[test]
    fn a_zip_that_is_not_ooxml_is_probed_once() {
        use std::io::Write as _;

        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        writer
            .start_file("plain.txt", zip::write::SimpleFileOptions::default())
            .unwrap();
        writer.write_all(b"not an office package").unwrap();
        let bytes = writer.finish().unwrap().into_inner();
        reset_opc_probe_count();
        let error = crate::open_bytes(bytes.clone()).expect_err("not an Office file");
        assert_eq!(error.code(), crate::Code::UnsupportedFormat);
        assert_eq!(opc_probe_count(), 1);
        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(&bytes, "");
            reset_opc_probe_count();
            crate::open(temporary.path()).expect_err("not an Office file");
            assert_eq!(opc_probe_count(), 1);
        }
    }

    #[cfg(any(feature = "doc", feature = "rtf"))]
    #[test]
    fn non_zip_inputs_build_no_opc_index() {
        let fixtures = [
            #[cfg(feature = "doc")]
            "ole/doc/FancyFoot.doc",
            #[cfg(feature = "rtf")]
            "rtf/testUnicode.rtf",
        ];
        for relative in fixtures {
            reset_opc_probe_count();
            crate::open(test_data(relative)).unwrap();
            crate::open_bytes(std::fs::read(test_data(relative)).unwrap()).unwrap();
            assert_eq!(opc_probe_count(), 0, "{relative}");
        }
    }

    /// An ODF package whose ZIP catalog the cheap central-directory probe
    /// cannot vouch for (LibreOffice writes data descriptors) gets the one
    /// bounded OPC probe the facade ran, which declines it; it is never
    /// probed twice.
    #[cfg(feature = "odt")]
    #[test]
    fn an_odt_gets_at_most_one_declined_opc_probe() {
        let relative = "odf/odt/table-cell-column-span.odt";
        reset_opc_probe_count();
        assert!(matches!(
            crate::open(test_data(relative)).unwrap(),
            File::Odt(_)
        ));
        assert!(opc_probe_count() <= 1, "{}", opc_probe_count());
        reset_opc_probe_count();
        let bytes = std::fs::read(test_data(relative)).unwrap();
        assert!(matches!(crate::open_bytes(bytes).unwrap(), File::Odt(_)));
        assert!(opc_probe_count() <= 1, "{}", opc_probe_count());
    }
}

/// Read limits through [`OpenOptions`](crate::OpenOptions), mirroring the
/// facade's `read_limits_facade.rs`.
mod limits {
    #[cfg(any(
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
    ))]
    #[test]
    fn the_input_ceiling_bounds_every_route() {
        use crate::{Code, OpenOptions};

        for (relative, _) in super::fixtures() {
            let path = super::super::test_data(relative);
            let bytes = std::fs::read(&path).unwrap();
            let tight = OpenOptions::new().max_input_bytes(bytes.len() as u64 - 1);
            let exact = OpenOptions::new().max_input_bytes(bytes.len() as u64);
            for error in [
                crate::open_with(&path, &tight).expect_err("over the ceiling"),
                crate::open_bytes_with(bytes.clone(), &tight).expect_err("over the ceiling"),
            ] {
                assert_eq!(error.code(), Code::LimitExceeded, "{relative}: {error:?}");
                assert!(
                    matches!(
                        error.leaf(),
                        crate::error::Source::Core(litchi_core::Error::ResourceLimit(limit))
                            if limit.resource == litchi_core::Resource::InputBytes
                                && limit.observed == bytes.len() as u64
                                && limit.limit == bytes.len() as u64 - 1
                    ),
                    "{relative}: {error:?}"
                );
            }
            crate::open_with(&path, &exact).unwrap_or_else(|error| panic!("{relative}: {error:?}"));
            crate::open_bytes_with(bytes, &exact)
                .unwrap_or_else(|error| panic!("{relative}: {error:?}"));
        }
    }

    /// A new package saved by its leaf, as `read_limits_facade.rs` builds it.
    #[cfg(any(feature = "docx", feature = "pptx"))]
    fn saved(extension: &str, save: impl FnOnce(&std::path::Path)) -> Vec<u8> {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join(format!("bounded.{extension}"));
        save(&path);
        std::fs::read(&path).unwrap()
    }

    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
    fn new_packages() -> Vec<(&'static str, Vec<u8>)> {
        vec![
            #[cfg(feature = "docx")]
            (
                "docx",
                saved("docx", |path| {
                    litchi_docx::Package::new().unwrap().save(path).unwrap();
                }),
            ),
            #[cfg(feature = "pptx")]
            (
                "pptx",
                saved("pptx", |path| {
                    litchi_pptx::Package::new().unwrap().save(path).unwrap();
                }),
            ),
            #[cfg(feature = "xlsx")]
            (
                "xlsx",
                std::fs::read(super::super::test_data("ooxml/xlsx/styles.xlsx")).unwrap(),
            ),
        ]
    }

    /// `document_limits_bound_path_and_byte_ingress` and its presentation
    /// twin: the OPC input limit bounds the path and byte routes, with and
    /// without an OOXML file name.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
    #[test]
    fn ooxml_limits_bound_path_and_byte_ingress() {
        use crate::{Code, OpenOptions};

        for (extension, bytes) in new_packages() {
            let limits = litchi_opc::SourceLimits::builder()
                .max_input_bytes(bytes.len() as u64 - 1)
                .unwrap()
                .build()
                .unwrap();
            let options = OpenOptions::new().ooxml_limits(limits);
            let directory = tempfile::tempdir().unwrap();
            let named = directory.path().join(format!("bounded.{extension}"));
            let unnamed = directory.path().join("bounded-package");
            std::fs::write(&named, &bytes).unwrap();
            std::fs::write(&unnamed, &bytes).unwrap();

            crate::open(&named).unwrap_or_else(|error| panic!("{extension}: {error:?}"));
            for error in [
                crate::open_with(&named, &options).expect_err("over the OPC input limit"),
                crate::open_with(&unnamed, &options).expect_err("over the OPC input limit"),
                crate::open_bytes_with(bytes.clone(), &options)
                    .expect_err("over the OPC input limit"),
            ] {
                assert_eq!(error.code(), Code::LimitExceeded, "{extension}: {error:?}");
                assert!(
                    matches!(
                        error.leaf(),
                        crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                            kind: litchi_opc::SourceLimitKind::InputBytes,
                            observed: actual,
                            ..
                        }) if *actual == bytes.len() as u64
                    ),
                    "{extension}: {error:?}"
                );
            }

            // An OOXML name on input that is not a ZIP is held to the same
            // limit before its ZIP magic is checked.
            let invalid = directory.path().join(format!("oversized.{extension}"));
            std::fs::write(&invalid, [b'x'; 32]).unwrap();
            let eight = litchi_opc::SourceLimits::builder()
                .max_input_bytes(8)
                .unwrap()
                .build()
                .unwrap();
            let error = crate::open_with(&invalid, &OpenOptions::new().ooxml_limits(eight))
                .expect_err("over the OPC input limit");
            assert_eq!(error.code(), Code::LimitExceeded, "{extension}: {error:?}");
        }
    }

    /// A part-count limit stops the OPC index with the exact bounds
    /// (`path_part_limit_returns_typed_opc_error_with_exact_bounds`).
    #[cfg(all(feature = "pptx", any(unix, windows)))]
    #[test]
    fn the_part_limit_stops_the_index_with_exact_bounds() {
        use std::io::Write as _;

        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="txt" ContentType="text/plain"/></Types>"#,
            )
            .unwrap();
        writer.start_file("one.txt", options).unwrap();
        writer.write_all(b"one").unwrap();
        writer.start_file("two.txt", options).unwrap();
        writer.write_all(b"two").unwrap();
        let temporary =
            super::super::temporary_with(&writer.finish().unwrap().into_inner(), ".pptx");
        let limits = litchi_opc::SourceLimits::builder()
            .max_parts(1)
            .unwrap()
            .build()
            .unwrap();
        let error = crate::open_with(
            temporary.path(),
            &crate::OpenOptions::new().ooxml_limits(limits),
        )
        .expect_err("the part limit stops the index");
        assert_eq!(error.code(), crate::Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Opc(litchi_opc::OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::Parts,
                observed: 2,
                maximum: 1,
            })
        ));
    }

    /// `generic_non_zip_over_neutral_ceiling_reports_exact_limit`: without
    /// an explicit ceiling every route applies the facade's neutral 2 GiB
    /// ceiling, reported with its exact bounds.
    #[cfg(all(
        any(unix, windows),
        any(
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
        )
    ))]
    #[test]
    fn the_neutral_ceiling_bounds_an_unnamed_path() {
        let temporary = tempfile::NamedTempFile::new().unwrap();
        let neutral = 2 * 1024 * 1024 * 1024_u64;
        temporary.as_file().set_len(neutral + 1).unwrap();
        let error = crate::open(temporary.path()).expect_err("over the neutral ceiling");
        assert_eq!(error.code(), crate::Code::LimitExceeded);
        assert!(matches!(
            error.leaf(),
            crate::error::Source::Core(litchi_core::Error::ResourceLimit(limit))
                if limit.resource == litchi_core::Resource::InputBytes
                    && limit.observed == neutral + 1
                    && limit.limit == neutral
        ));
    }
}

/// An execution context bounds an open by its cancellation token only
/// (record 0008, finding A1).
#[cfg(any(
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
))]
mod execution {
    use std::num::{NonZeroU64, NonZeroUsize};

    use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};

    use super::super::test_data;
    use crate::{Code, OpenOptions};

    fn context() -> (CancellationSource, ExecutionContext) {
        let budget = Budget::root(
            "litchi-ops-open-test",
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (source, token) = CancellationSource::pair();
        let limits =
            ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap();
        (source, ExecutionContext::new(budget, token, limits))
    }

    /// Record 0008 A1: the managed source-backed DOCX route refuses to read
    /// the Word-authored fixtures, so an open under a context keeps the
    /// ordinary owners. Every DOCX fixture reads under one.
    #[cfg(feature = "docx")]
    #[test]
    fn every_word_fixture_reads_under_an_execution_context() {
        let (_cancellation, context) = context();
        let options = OpenOptions::new().execution(context);
        let mut read = 0;
        for entry in std::fs::read_dir(test_data("ooxml/docx")).unwrap() {
            let path = entry.unwrap().path();
            if path.extension().and_then(std::ffi::OsStr::to_str) != Some("docx") {
                continue;
            }
            let file = crate::open_with(&path, &options)
                .unwrap_or_else(|error| panic!("{}: {error:?}", path.display()));
            let crate::File::Docx(docx) = &file else {
                panic!("{}: expected DOCX, got {file:?}", path.display());
            };
            assert!(docx.is_source_backed());
            file.text()
                .unwrap_or_else(|error| panic!("{}: {error:?}", path.display()));
            file.document()
                .unwrap()
                .paragraphs()
                .unwrap_or_else(|error| panic!("{}: {error:?}", path.display()));
            let bytes = crate::open_bytes_with(std::fs::read(&path).unwrap(), &options).unwrap();
            assert_eq!(bytes.text().unwrap(), file.text().unwrap());
            read += 1;
        }
        assert_eq!(read, 19, "the Word-authored DOCX fixtures of record 0008");
    }

    #[test]
    fn a_cancelled_context_refuses_the_open() {
        let (cancellation, context) = context();
        cancellation.cancel();
        let options = OpenOptions::new().execution(context);
        for (relative, _) in super::fixtures() {
            let path = test_data(relative);
            for error in [
                crate::open_with(&path, &options).expect_err("cancelled"),
                crate::open_bytes_with(std::fs::read(&path).unwrap(), &options)
                    .expect_err("cancelled"),
            ] {
                assert_eq!(error.code(), Code::Cancelled, "{relative}: {error:?}");
            }
        }
    }
}

/// [`SourceMode::Eager`](crate::SourceMode::Eager) builds the eager owner
/// from the eager detection route (record 0008, finding A5).
#[cfg(feature = "docx")]
mod eager {
    use super::super::test_data;
    use crate::format::docx::DocxOwner;
    use crate::{File, OpenOptions, SourceMode};

    fn docx_owner(file: &File) -> &DocxOwner {
        match file {
            File::Docx(docx) => docx.owner(),
            other => panic!("expected DOCX, got {other:?}"),
        }
    }

    #[test]
    fn eager_mode_docx_keeps_the_eager_package_it_indexed() {
        let path = test_data("ooxml/docx/table-alignment.docx");
        let eager = OpenOptions::new().source_mode(SourceMode::Eager);
        let source = crate::open(&path).unwrap();
        assert!(matches!(docx_owner(&source), DocxOwner::Source(_)));
        for file in [
            crate::open_with(&path, &eager).unwrap(),
            crate::open_bytes_with(std::fs::read(&path).unwrap(), &eager).unwrap(),
        ] {
            assert!(matches!(docx_owner(&file), DocxOwner::Eager(_)));
            assert_eq!(file.text().unwrap(), source.text().unwrap());
            // The eager open validated the main story and recorded its
            // repairs, as the facade's eager route did.
            file.tolerance_report().unwrap();
        }
    }

    /// `open_prepared` under `Eager` materializes the source-backed index
    /// it adopts (`to_owned_package`); detection prepared no eager package.
    #[cfg(any(unix, windows))]
    #[test]
    fn eager_mode_open_prepared_materializes_the_prepared_index() {
        let path = test_data("ooxml/docx/FancyFoot.docx");
        let detection =
            crate::detect::detect_and_prepare(&path, &crate::detect::SourceLimits::default())
                .unwrap();
        let file = crate::open_prepared_with(
            detection,
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )
        .unwrap();
        assert!(matches!(docx_owner(&file), DocxOwner::Eager(_)));
        assert_eq!(
            file.text().unwrap(),
            crate::open(&path).unwrap().text().unwrap()
        );
    }
}

/// The lowest-precedence ZIP candidate: only a caller's tail names iWork
/// packages, and `litchi-ops` refuses them.
#[cfg(any(
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
mod tail {
    use std::io::{Read, Seek, Write as _};

    use litchi_core::FileFormat;

    use crate::detect::DetectedFormat;
    use crate::detect::tail::ZipTail;
    use crate::{Code, OpenOptions};

    /// Claims every ZIP it is offered as a Pages document.
    struct PagesTail;

    impl ZipTail for PagesTail {
        fn classify_bytes(&self, _bytes: &[u8]) -> Option<FileFormat> {
            Some(FileFormat::Pages)
        }

        fn classify_reader<R: Read + Seek>(&self, _reader: &mut R) -> Option<FileFormat> {
            Some(FileFormat::Pages)
        }

        fn detect_owned(&self, bytes: Vec<u8>) -> Option<DetectedFormat> {
            drop(bytes);
            None
        }
    }

    fn plain_zip() -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        writer
            .start_file(
                "Index/Document.iwa",
                zip::write::SimpleFileOptions::default(),
            )
            .unwrap();
        writer.write_all(b"not really iWork").unwrap();
        writer.finish().unwrap().into_inner()
    }

    #[test]
    fn a_zip_the_tail_names_is_refused_as_iwork() {
        let options = OpenOptions::new();
        let error = crate::open_bytes_with_tail(plain_zip(), &options, &PagesTail)
            .expect_err("iWork does not open in litchi-ops");
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), Some(FileFormat::Pages));
        assert!(error.suggestion().is_some());

        let error = crate::open_bytes(plain_zip()).expect_err("no tail names it");
        assert_eq!(error.code(), Code::UnsupportedFormat);
        assert_eq!(error.detected(), None);

        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(&plain_zip(), ".pages");
            let error = crate::open_with_tail(temporary.path(), &options, &PagesTail)
                .expect_err("iWork does not open in litchi-ops");
            assert_eq!(error.detected(), Some(FileFormat::Pages));
            let error = crate::open(temporary.path()).expect_err("no tail names it");
            assert_eq!(error.detected(), None);
        }
    }

    /// The tail is the lowest-precedence candidate: inputs another candidate
    /// claims never reach it.
    #[cfg(feature = "rtf")]
    #[test]
    fn the_tail_never_sees_inputs_another_candidate_claims() {
        let file = crate::open_bytes_with_tail(
            br"{\rtf1 Hello}".to_vec(),
            &OpenOptions::new(),
            &PagesTail,
        )
        .unwrap();
        assert_eq!(file.format(), Some(FileFormat::Rtf));
    }
}

/// ZIP inputs at the edges of OOXML and OpenDocument detection, through the
/// openers. M0a step S9 moved these cases here from the tests of the
/// per-family probes (`detect_docx_source_bytes`,
/// `detect_workbook_source_bytes`, `detect_presentation_source_bytes`,
/// `detect_prepared_ods`, `detect_prepared_odp` and
/// `detect_workbook_source_path`) that the format-family types used, which
/// S9 deleted with those types.
#[cfg(any(
    feature = "docx",
    feature = "pptx",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "odp"
))]
mod zip_candidates {
    use std::io::Write as _;

    /// A ZIP of stored members.
    #[allow(
        dead_code,
        reason = "used by the cases of whichever formats are enabled"
    )]
    fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        for (name, data) in entries {
            writer.start_file(*name, options).unwrap();
            writer.write_all(data).unwrap();
        }
        writer.finish().unwrap().into_inner()
    }

    /// An ODF package of `mimetype` whose `content.xml` is `content`, plus
    /// the `extra` members.
    #[cfg(any(feature = "ods", feature = "odp"))]
    fn odf(mimetype: &str, content: &[u8], extra: &[(&str, &[u8])]) -> Vec<u8> {
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(litchi_odf_common::constants::ODF_CONTENT, content)
            .unwrap();
        for (name, data) in extra {
            writer.add_file(name, data).unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    #[cfg(feature = "ods")]
    const ODS_CONTENT: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#;

    /// An OOXML/ODF polyglot whose OPC catalog the cheap probe cannot vouch
    /// for: a harmless `./custom.xml` alias makes it answer "uncertain",
    /// while the full OPC reader normalizes the alias and sees a valid DOCX.
    #[cfg(all(
        any(feature = "ods", feature = "odp"),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    fn uncertain_ooxml_odf_polyglot(mimetype: &str) -> Vec<u8> {
        zip(&[
            ("mimetype", mimetype.as_bytes()),
            (
                "content.xml",
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#,
            ),
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            ),
            (
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
            (
                "word/document.xml",
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#,
            ),
            ("./custom.xml", b"alias"),
        ])
    }

    /// The polyglot keeps OOXML precedence: it opens as the DOCX its catalog
    /// names (refused as a disabled format without `docx`) after one OPC
    /// probe, never as the ODF package its `mimetype` names.
    #[cfg(all(
        any(feature = "ods", feature = "odp"),
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[track_caller]
    fn assert_ooxml_precedence(mimetype: &str) {
        use crate::detect::{opc_probe_count, reset_opc_probe_count};

        let bytes = uncertain_ooxml_odf_polyglot(mimetype);
        assert_eq!(
            litchi_odf_common::detect::packaged_has_ooxml_catalog(
                &bytes,
                litchi_odf_common::detect::CatalogProbeLimits::default()
            ),
            None
        );
        reset_opc_probe_count();
        let opened = crate::open_bytes(bytes);
        assert_eq!(opc_probe_count(), 1);
        #[cfg(feature = "docx")]
        assert!(matches!(opened, Ok(crate::File::Docx(_))), "{opened:?}");
        #[cfg(not(feature = "docx"))]
        {
            let error = opened.expect_err("DOCX is not enabled");
            assert_eq!(error.code(), crate::Code::UnsupportedFormat);
            assert_eq!(error.detected(), Some(litchi_core::FileFormat::Docx));
        }
    }

    /// An ordinary ODS opens from the package detection prepared.
    #[cfg(feature = "ods")]
    #[test]
    fn an_ordinary_ods_opens_as_ods() {
        let bytes = odf(
            litchi_odf_common::constants::ODF_SPREADSHEET,
            ODS_CONTENT,
            &[],
        );
        assert!(matches!(
            crate::detect::detect_format_smart(
                bytes.clone(),
                &crate::detect::SourceLimits::default(),
            ),
            Some(crate::detect::DetectedFormat::Ods(_))
        ));
        let file = crate::open_bytes(bytes).unwrap();
        assert!(matches!(file, crate::File::Ods(_)), "{file:?}");
        assert_eq!(
            file.workbook().unwrap().worksheet_names().unwrap(),
            ["Sheet1"]
        );
    }

    /// An ODS whose `[Content_Types].xml` is not an OPC catalog is still an
    /// ODS: the declined OPC probe falls back to the ODF owner, by path and
    /// from bytes. An ordinary ODS needs no OPC probe at all.
    #[cfg(all(
        feature = "ods",
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[test]
    fn ods_opens_with_ordinary_and_malformed_ooxml_catalogs() {
        use crate::detect::{
            DetectedFormat, detect_format_smart, opc_probe_count, reset_opc_probe_count,
        };

        let ordinary = odf(
            litchi_odf_common::constants::ODF_SPREADSHEET,
            ODS_CONTENT,
            &[],
        );
        reset_opc_probe_count();
        assert!(matches!(
            detect_format_smart(ordinary.clone(), &crate::detect::SourceLimits::default()),
            Some(DetectedFormat::Ods(_))
        ));
        assert_eq!(opc_probe_count(), 0);
        let limits = crate::opc::SourceLimits::builder()
            .max_members(1)
            .expect("test member limit")
            .max_parts(1)
            .expect("test part limit")
            .max_relationship_parts(1)
            .expect("test relationship limit")
            .build()
            .expect("consistent test limits");
        reset_opc_probe_count();
        assert!(matches!(
            detect_format_smart(
                ordinary.clone(),
                &crate::detect::SourceLimits::default().ooxml_source_limits(limits)
            ),
            Some(DetectedFormat::Ods(_))
        ));
        assert_eq!(opc_probe_count(), 0);
        reset_opc_probe_count();
        assert!(matches!(
            crate::open_bytes(ordinary).unwrap(),
            crate::File::Ods(_)
        ));
        assert_eq!(opc_probe_count(), 0);

        let malformed = odf(
            litchi_odf_common::constants::ODF_SPREADSHEET,
            ODS_CONTENT,
            &[("[Content_Types].xml", b"<not-an-opc-types-root/>")],
        );
        reset_opc_probe_count();
        assert!(matches!(
            detect_format_smart(malformed.clone(), &crate::detect::SourceLimits::default()),
            Some(DetectedFormat::Ods(_))
        ));
        assert_eq!(opc_probe_count(), 1);
        reset_opc_probe_count();
        assert!(matches!(
            crate::open_bytes(malformed.clone()).unwrap(),
            crate::File::Ods(_)
        ));
        assert_eq!(opc_probe_count(), 1);
        #[cfg(any(unix, windows))]
        {
            let temporary = super::super::temporary_with(&malformed, "");
            assert!(matches!(
                crate::open(temporary.path()).unwrap(),
                crate::File::Ods(_)
            ));
        }
    }

    #[cfg(all(
        feature = "ods",
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[test]
    fn ods_keeps_ooxml_precedence_for_an_uncertain_alias_catalog() {
        assert_ooxml_precedence(litchi_odf_common::constants::ODF_SPREADSHEET);
    }

    /// An ordinary ODP opens from the package detection prepared.
    #[cfg(feature = "odp")]
    #[test]
    fn an_ordinary_odp_opens_as_odp() {
        let bytes = odf(
            litchi_odf_common::constants::ODF_PRESENTATION,
            br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
            &[],
        );
        assert!(matches!(
            crate::detect::detect_format_smart(
                bytes.clone(),
                &crate::detect::SourceLimits::default(),
            ),
            Some(crate::detect::DetectedFormat::Odp(_))
        ));
        let file = crate::open_bytes(bytes).unwrap();
        assert!(matches!(file, crate::File::Odp(_)), "{file:?}");
        assert_eq!(file.presentation().unwrap().slide_count().unwrap(), 0);
    }

    #[cfg(all(
        feature = "odp",
        any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb")
    ))]
    #[test]
    fn odp_keeps_ooxml_precedence_for_an_uncertain_alias_catalog() {
        assert_ooxml_precedence(litchi_odf_common::constants::ODF_PRESENTATION);
    }

    /// A valid OPC package whose main part is a presentation is classified
    /// as PPTX for good: it is opened, or refused, as PPTX, and never handed
    /// to the lower-precedence candidates.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[test]
    fn a_valid_non_docx_opc_package_is_classified_terminally() {
        let bytes = super::super::minimal_ooxml(
            "ppt/presentation.xml",
            super::super::content_type::PPTX,
            br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#,
            None,
        );
        match crate::open_bytes(bytes) {
            Ok(file) => assert_eq!(
                file.format(),
                Some(litchi_core::FileFormat::Pptx),
                "{file:?}"
            ),
            Err(error) => assert_eq!(
                error.format().or(error.detected()),
                Some(litchi_core::FileFormat::Pptx),
                "{error:?}"
            ),
        }
    }

    /// A DOCX catalog whose package lacks the officeDocument relationship
    /// is still a DOCX: the DOCX owner's refusal is reported, not a
    /// lower-precedence fallback.
    #[cfg(feature = "docx")]
    #[test]
    fn a_docx_catalog_without_its_relationship_reports_the_docx_error() {
        let bytes = zip(&[
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            ),
            (
                "word/document.xml",
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
            ),
        ]);
        let error = crate::open_bytes(bytes).expect_err("no officeDocument relationship");
        assert_eq!(
            error.format(),
            Some(litchi_core::FileFormat::Docx),
            "{error:?}"
        );
        assert_eq!(error.code(), crate::Code::CorruptFile, "{error:?}");
    }

    /// Bytes that are not a ZIP, a ZIP without an OPC catalog and a valid OPC
    /// package with no Office main part are not Office files.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[test]
    fn zip_inputs_without_an_office_main_part_are_unsupported() {
        let inputs = [
            b"not an XLSX package".to_vec(),
            zip(&[("plain.txt", b"not an OPC package")]),
            zip(&[
                (
                    "[Content_Types].xml",
                    br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="txt" ContentType="text/plain"/></Types>"#,
                ),
                ("plain.txt", b"not a workbook"),
            ]),
        ];
        for bytes in inputs {
            let error = crate::open_bytes(bytes).expect_err("not an Office file");
            assert_eq!(error.code(), crate::Code::UnsupportedFormat, "{error:?}");
            assert_eq!(error.detected(), None, "{error:?}");
        }
    }

    /// The OOXML read limits refuse a ZIP candidate with the exact resource,
    /// count and bound.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[test]
    fn ooxml_limits_refuse_a_zip_candidate_with_the_exact_bound() {
        use crate::error::Source;
        use crate::opc::{OpcError, SourceLimitKind, SourceLimits};

        let input = SourceLimits::builder()
            .max_input_bytes(1)
            .expect("test input limit")
            .build()
            .expect("consistent test limits");
        let error = crate::open_bytes_with(
            b"PK\x03\x04".to_vec(),
            &crate::OpenOptions::new().ooxml_limits(input),
        )
        .expect_err("the input exceeds the limit");
        assert_eq!(error.code(), crate::Code::LimitExceeded, "{error:?}");
        assert!(
            matches!(
                error.leaf(),
                Source::Opc(OpcError::Limit {
                    kind: SourceLimitKind::InputBytes,
                    observed: 4,
                    maximum: 1,
                })
            ),
            "{error:?}"
        );

        let parts = SourceLimits::builder()
            .max_parts(1)
            .expect("test part limit")
            .build()
            .expect("consistent test limits");
        let bytes = zip(&[
            (
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="txt" ContentType="text/plain"/></Types>"#,
            ),
            ("one.txt", b"one"),
            ("two.txt", b"two"),
        ]);
        let error = crate::open_bytes_with(bytes, &crate::OpenOptions::new().ooxml_limits(parts))
            .expect_err("the package exceeds the part limit");
        assert_eq!(error.code(), crate::Code::LimitExceeded, "{error:?}");
        assert!(
            matches!(
                error.leaf(),
                Source::Opc(OpcError::Limit {
                    kind: SourceLimitKind::Parts,
                    observed: 2,
                    maximum: 1,
                })
            ),
            "{error:?}"
        );
    }

    /// A truncated ZIP candidate is a corrupt file, with the OPC reader's ZIP
    /// error as the leaf; it is not handed to a lower-precedence candidate.
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    #[test]
    fn a_malformed_zip_candidate_is_refused() {
        let error = crate::open_bytes(b"PK\x03\x04malformed".to_vec()).expect_err("malformed");
        assert_eq!(error.code(), crate::Code::CorruptFile, "{error:?}");
        assert!(
            matches!(
                error.leaf(),
                crate::error::Source::Opc(crate::opc::OpcError::ZipError(_))
            ),
            "{error:?}"
        );
    }
}
