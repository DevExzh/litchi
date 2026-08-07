//! Compile-only coverage for contextual OOXML ingestion limits exposed through
//! the umbrella facade.

#[cfg(feature = "docx")]
#[test]
fn docx_limits_are_nameable() {
    let _: litchi::docx::ReadLimits = Default::default();
    let _ = litchi::Document::open_with_limits::<&std::path::Path>;
    let _ = litchi::Document::from_bytes_with_limits;
    let _ = litchi::detection_smart::detect_format_smart_with_limits;
}

#[cfg(feature = "docx")]
#[test]
fn document_limits_bound_path_and_byte_ingress() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("bounded.docx");
    let mut package = litchi::docx::Package::new().expect("new DOCX package");
    package.save(&path).expect("save DOCX package");

    let bytes = std::fs::read(&path).expect("read DOCX fixture");
    let limits = litchi::docx::ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("valid limits");

    assert!(litchi::Document::open(&path).is_ok());
    assert!(litchi::Document::open_with_limits(&path, limits).is_err());
    assert!(litchi::Document::from_bytes_with_limits(bytes, limits).is_err());
}

#[cfg(feature = "docx")]
#[test]
fn document_limits_reject_invalid_suffix_and_unknown_extension_ooxml() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let invalid = directory.path().join("oversized.docx");
    std::fs::write(&invalid, vec![b'x'; 32]).expect("write invalid DOCX");
    let limits = litchi::docx::ReadLimits::builder()
        .max_input_bytes(8)
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert!(litchi::Document::open_with_limits(&invalid, limits).is_err());

    let unknown = directory.path().join("bounded-package");
    let mut package = litchi::docx::Package::new().expect("new DOCX package");
    package.save(&unknown).expect("save DOCX package");
    let bytes = std::fs::read(&unknown).expect("read DOCX fixture");
    let tight = litchi::docx::ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert!(litchi::Document::open_with_limits(&unknown, tight).is_err());
    assert!(
        litchi::detection_smart::ooxml::detect_ooxml_format_with_limits(&bytes, tight).is_none()
    );
}

#[cfg(feature = "pptx")]
#[test]
fn pptx_limits_are_nameable() {
    let _: litchi::pptx::ReadLimits = Default::default();
    let _ = litchi::Presentation::open_with_limits::<&std::path::Path>;
    let _ = litchi::Presentation::from_bytes_with_limits;
    let _ = litchi::detection_smart::detect_format_smart_with_limits;
}

#[cfg(feature = "pptx")]
#[test]
fn presentation_limits_bound_path_and_byte_ingress() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory.path().join("bounded.pptx");
    let mut package = litchi::pptx::Package::new().expect("new PPTX package");
    package.save(&path).expect("save PPTX package");

    let bytes = std::fs::read(&path).expect("read PPTX fixture");
    let limits = litchi::pptx::ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("valid limits");

    assert!(litchi::Presentation::open(&path).is_ok());
    assert!(litchi::Presentation::open_with_limits(&path, limits).is_err());
    assert!(litchi::Presentation::from_bytes_with_limits(bytes, limits).is_err());
}

#[cfg(feature = "pptx")]
#[test]
fn presentation_limits_reject_invalid_suffix_and_unknown_extension_ooxml() {
    let directory = tempfile::tempdir().expect("temporary directory");
    let invalid = directory.path().join("oversized.pptx");
    std::fs::write(&invalid, vec![b'x'; 32]).expect("write invalid PPTX");
    let limits = litchi::pptx::ReadLimits::builder()
        .max_input_bytes(8)
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert!(litchi::Presentation::open_with_limits(&invalid, limits).is_err());

    let unknown = directory.path().join("bounded-package");
    let mut package = litchi::pptx::Package::new().expect("new PPTX package");
    package.save(&unknown).expect("save PPTX package");
    let bytes = std::fs::read(&unknown).expect("read PPTX fixture");
    let tight = litchi::pptx::ReadLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len() - 1).expect("fixture length fits u64"))
        .expect("positive input limit")
        .build()
        .expect("valid limits");
    assert!(litchi::Presentation::open_with_limits(&unknown, tight).is_err());
    assert!(
        litchi::detection_smart::ooxml::detect_ooxml_format_from_bytes_with_limits(&bytes, tight)
            .is_none()
    );
}

#[cfg(feature = "xlsx")]
#[test]
fn xlsx_limits_and_bounded_sheet_helpers_are_nameable() {
    let _: litchi::xlsx::ReadLimits = Default::default();
    let _ = litchi::sheet::open_workbook_with_limits::<&std::path::Path>;
    let _ = litchi::sheet::open_workbook_from_bytes_with_limits;
}

#[cfg(feature = "xlsb")]
#[test]
fn xlsb_limits_and_bounded_sheet_helpers_are_nameable() {
    let _: litchi::xlsb::ReadLimits = Default::default();
    let _ = litchi::sheet::open_xlsb_workbook_with_limits::<&std::path::Path>;
    let _ = litchi::sheet::open_xlsb_workbook_from_bytes_with_limits;
    let _ = litchi::sheet::open_xlsb_workbook_dyn_with_limits::<&std::path::Path>;
    let _ = litchi::sheet::open_xlsb_workbook_from_bytes_dyn_with_limits;
}
