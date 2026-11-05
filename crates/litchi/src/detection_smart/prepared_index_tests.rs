//! Evidence that one detect-and-prepare plus one adopt builds the container
//! index exactly once, where detect-then-open builds it twice.
//!
//! These tests detect through the facade's detection functions, which add the
//! iWork tail, so they stay in the facade. They open through `crate::open`
//! and adopt through `crate::open_prepared` (M0a step S8c ported them from
//! the format-family types' `open` and `from_prepared`).

use litchi_core::detection::FileFormat;

fn fixture(relative: &str) -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// Archive constructions recorded while running `body`.
#[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx"))]
fn archive_constructions(body: impl FnOnce()) -> usize {
    crate::detection_smart::reset_opc_probe_count();
    body();
    crate::detection_smart::opc_probe_count()
}

#[test]
#[cfg(feature = "docx")]
fn docx_detect_then_open_indexes_twice_and_prepared_open_indexes_once() {
    let path = fixture("ooxml/docx/comment.docx");

    let two_call = archive_constructions(|| {
        assert_eq!(
            crate::detection_smart::detect_file_format(
                &path,
                &crate::detection_smart::SourceLimits::default(),
            ),
            Some(FileFormat::Docx)
        );
        let file = crate::open(&path).expect("open");
        assert_eq!(file.format(), Some(FileFormat::Docx));
    });

    let single_call = archive_constructions(|| {
        let detected = crate::detection_smart::detect_and_prepare(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .expect("detect");
        assert_eq!(detected.format(), FileFormat::Docx);
        assert!(detected.is_prepared(), "prepared handle");
        let file = crate::open_prepared(detected).expect("adopt");
        assert_eq!(file.format(), Some(FileFormat::Docx));
    });

    assert_eq!(
        two_call, 2,
        "detect-then-open must index the container twice"
    );
    assert_eq!(
        single_call, 1,
        "detect-and-prepare plus adopt must index the container once"
    );
}

#[test]
#[cfg(feature = "pptx")]
fn pptx_detect_then_open_indexes_twice_and_prepared_open_indexes_once() {
    let path = fixture("ooxml/pptx/shapes.pptx");

    let two_call = archive_constructions(|| {
        assert_eq!(
            crate::detection_smart::detect_file_format(
                &path,
                &crate::detection_smart::SourceLimits::default(),
            ),
            Some(FileFormat::Pptx)
        );
        let file = crate::open(&path).expect("open");
        assert_eq!(file.format(), Some(FileFormat::Pptx));
    });

    let single_call = archive_constructions(|| {
        let detected = crate::detection_smart::detect_and_prepare(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .expect("detect");
        assert_eq!(detected.format(), FileFormat::Pptx);
        assert!(detected.is_prepared(), "prepared handle");
        let file = crate::open_prepared(detected).expect("adopt");
        assert_eq!(file.format(), Some(FileFormat::Pptx));
    });

    assert_eq!(
        two_call, 2,
        "detect-then-open must index the container twice"
    );
    assert_eq!(
        single_call, 1,
        "detect-and-prepare plus adopt must index the container once"
    );
}

#[test]
#[cfg(feature = "xlsx")]
fn xlsx_detect_then_open_indexes_twice_and_prepared_open_indexes_once() {
    let path = fixture("ooxml/xlsx/styles.xlsx");

    let two_call = archive_constructions(|| {
        assert_eq!(
            crate::detection_smart::detect_file_format(
                &path,
                &crate::detection_smart::SourceLimits::default(),
            ),
            Some(FileFormat::Xlsx)
        );
        let file = crate::open(&path).expect("open");
        assert_eq!(file.format(), Some(FileFormat::Xlsx));
    });

    let single_call = archive_constructions(|| {
        let detected = crate::detection_smart::detect_and_prepare(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .expect("detect");
        assert_eq!(detected.format(), FileFormat::Xlsx);
        assert!(detected.is_prepared(), "prepared handle");
        let file = crate::open_prepared(detected).expect("adopt");
        assert_eq!(file.format(), Some(FileFormat::Xlsx));
    });

    assert_eq!(
        two_call, 2,
        "detect-then-open must index the container twice"
    );
    assert_eq!(
        single_call, 1,
        "detect-and-prepare plus adopt must index the container once"
    );
}

/// A non-OOXML input keeps its classification and receives no handle.
#[test]
fn non_ooxml_inputs_are_still_classified_without_a_prepared_handle() {
    #[allow(unused_mut, reason = "a build with no non-OOXML format pushes no case")]
    let mut cases: Vec<(std::path::PathBuf, FileFormat)> = Vec::new();
    #[cfg(feature = "doc")]
    cases.push((fixture("ole/doc/NoHeadFoot.doc"), FileFormat::Doc));
    #[cfg(feature = "odt")]
    cases.push((
        fixture("odf/corpus/writer-paragraph-styles.odt"),
        FileFormat::Odt,
    ));
    #[cfg(feature = "odp")]
    cases.push((fixture("odf/corpus/impress-basic.odp"), FileFormat::Odp));
    #[cfg(feature = "ods")]
    cases.push((fixture("odf/corpus/calc-formulas.ods"), FileFormat::Ods));

    for (path, expected) in cases {
        let detected = crate::detection_smart::detect_and_prepare(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .unwrap_or_else(|| panic!("{} should be classified", path.display()));
        assert_eq!(detected.format(), expected, "{}", path.display());
        assert!(
            !detected.is_prepared(),
            "{} must not carry an OOXML handle",
            path.display()
        );
        assert_eq!(
            crate::detection_smart::detect_file_format(
                &path,
                &crate::detection_smart::SourceLimits::default(),
            ),
            Some(expected),
            "{} must agree with the neutral detector",
            path.display()
        );
    }
}

/// A file that is not an Office document is reported the same way by both
/// detectors.
#[test]
fn a_non_office_file_is_rejected_by_both_detectors() {
    let path = fixture("images/png/lena.png");
    assert!(
        crate::detection_smart::detect_file_format(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .is_none()
    );
    assert!(
        crate::detection_smart::detect_and_prepare(
            &path,
            &crate::detection_smart::SourceLimits::default(),
        )
        .is_none()
    );
}

/// The prepared handle refuses a source rewritten between preparation and
/// adoption, matching the crate's source-version contract.
#[test]
#[cfg(feature = "docx")]
fn a_prepared_handle_refuses_a_source_rewritten_before_adoption() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("prepared.docx");
    let docx = std::fs::read(fixture("ooxml/docx/comment.docx")).expect("fixture");
    std::fs::write(&path, &docx).expect("write fixture");

    let detected = crate::detection_smart::detect_and_prepare(
        &path,
        &crate::detection_smart::SourceLimits::default(),
    )
    .expect("detect");
    assert_eq!(detected.format(), FileFormat::Docx);
    assert!(detected.is_prepared(), "prepared handle");

    // Rewrite the file under the handle. `FileSource` derives its revision
    // from length and modification time, so a different payload is a
    // different revision.
    let mut rewritten = docx.clone();
    rewritten.extend_from_slice(b"trailing bytes that change the length");
    std::fs::write(&path, &rewritten).expect("rewrite fixture");

    let Err(error) = crate::open_prepared(detected) else {
        panic!("adoption must refuse a source rewritten after preparation");
    };
    assert_eq!(error.code(), crate::Code::IoError, "{error:?}");
    assert!(
        matches!(
            error.leaf(),
            crate::error::Source::Core(litchi_core::Error::SourceChanged { .. })
                | crate::error::Source::Opc(crate::opc::OpcError::SourceChanged { .. })
        ),
        "{error:?}"
    );
}

/// A prepared handle is refused by the wrong family view, which names the
/// format it actually carries.
#[test]
#[cfg(all(feature = "docx", feature = "pptx"))]
fn a_prepared_docx_handle_is_refused_by_the_presentation_opener() {
    let path = fixture("ooxml/docx/comment.docx");
    let detected = crate::detection_smart::detect_and_prepare(
        &path,
        &crate::detection_smart::SourceLimits::default(),
    )
    .expect("detect");
    assert!(detected.is_prepared(), "prepared handle");
    let file = crate::open_prepared(detected).expect("adopt");
    let Err(error) = file.presentation() else {
        panic!("a DOCX handle must be refused by the presentation view");
    };
    assert_eq!(error.code(), crate::Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), Some(FileFormat::Docx), "{error:?}");
}
