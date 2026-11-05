//! `File` tests: the single-pass openers, the family views, the `Locked` and
//! `Signed` states, and the facade's in-source document, presentation and
//! workbook tests, ported onto `File` with the same semantics (record 0003
//! §4, step S6).

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic by design"
)]

#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
mod document;
#[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
mod odf_policy;
mod open;
#[cfg(any(feature = "ppt", feature = "pptx"))]
mod presentation;
mod security;
#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "pptx",
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb"
))]
mod signature;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
mod workbook;

use std::io::{Cursor, Write as _};

use crate::error::{Error, Source};

/// A path under the workspace's `test-data`.
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
    feature = "odg"
))]
pub(crate) fn test_data(relative: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// A minimal OOXML package with one main part, optionally marked as an ODF
/// package too (an OOXML/ODF polyglot).
#[allow(
    dead_code,
    reason = "used by the OOXML suites of whichever formats are enabled"
)]
pub(crate) fn minimal_ooxml(
    main_part: &str,
    content_type: &str,
    main_xml: &[u8],
    odf_mime: Option<&str>,
) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{main_part}" ContentType="{content_type}"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{main_part}"/></Relationships>"#
    );
    writer.start_file("[Content_Types].xml", options).unwrap();
    writer.write_all(content_types.as_bytes()).unwrap();
    writer.start_file("_rels/.rels", options).unwrap();
    writer.write_all(root_relationships.as_bytes()).unwrap();
    writer.start_file(main_part, options).unwrap();
    writer.write_all(main_xml).unwrap();
    if let Some(odf_mime) = odf_mime {
        writer.start_file("mimetype", options).unwrap();
        writer.write_all(odf_mime.as_bytes()).unwrap();
        let manifest = format!(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="{odf_mime}"/></manifest:manifest>"#
        );
        writer.start_file("META-INF/manifest.xml", options).unwrap();
        writer.write_all(manifest.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

/// The main-part content types of the OOXML families.
#[allow(
    dead_code,
    reason = "used by the OOXML suites of whichever formats are enabled"
)]
pub(crate) mod content_type {
    pub(crate) const DOCX: &str =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
    pub(crate) const PPTX: &str =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
    pub(crate) const XLSX: &str =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
}

/// Copy `fixture` into a fresh temporary file with `suffix`.
#[cfg(any(unix, windows))]
#[allow(
    dead_code,
    reason = "used by the path suites of whichever formats are enabled"
)]
pub(crate) fn temporary_copy(fixture: &std::path::Path, suffix: &str) -> tempfile::NamedTempFile {
    let temporary = tempfile::Builder::new()
        .suffix(suffix)
        .tempfile()
        .expect("temporary file");
    std::fs::copy(fixture, temporary.path()).expect("copy fixture");
    temporary
}

/// Write `bytes` into a fresh temporary file with `suffix`.
#[cfg(any(unix, windows))]
#[allow(
    dead_code,
    reason = "used by the path suites of whichever formats are enabled"
)]
pub(crate) fn temporary_with(bytes: &[u8], suffix: &str) -> tempfile::NamedTempFile {
    let temporary = tempfile::Builder::new()
        .suffix(suffix)
        .tempfile()
        .expect("temporary file");
    std::fs::write(temporary.path(), bytes).expect("write temporary file");
    temporary
}

/// Append bytes to a file so its source revision changes.
#[cfg(any(unix, windows))]
#[allow(
    dead_code,
    reason = "used by the source-change suites of whichever formats are enabled"
)]
pub(crate) fn mutate(path: &std::path::Path) {
    let mut file = std::fs::OpenOptions::new()
        .append(true)
        .open(path)
        .expect("reopen source");
    file.write_all(b"source mutation").expect("mutate source");
}

/// The typed source revisions a source-change error carries, whichever
/// leaf reported it.
#[allow(
    dead_code,
    reason = "used by the source-change suites of whichever formats are enabled"
)]
pub(crate) fn source_change(
    error: &Error,
) -> Option<(litchi_core::SourceVersion, litchi_core::SourceVersion)> {
    let core = |error: &litchi_core::Error| match error {
        litchi_core::Error::SourceChanged { expected, observed } => Some((*expected, *observed)),
        _ => None,
    };
    #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
    let opc = |error: &litchi_opc::OpcError| match error {
        litchi_opc::OpcError::SourceChanged { expected, actual } => Some((*expected, *actual)),
        _ => None,
    };
    match error.leaf() {
        Source::Core(error) => core(error),
        #[cfg(any(feature = "docx", feature = "pptx", feature = "xlsx", feature = "xlsb"))]
        Source::Opc(error) => opc(error),
        #[cfg(feature = "docx")]
        Source::Docx(litchi_docx::Error::Opc(error)) => opc(error),
        #[cfg(feature = "pptx")]
        Source::Pptx(litchi_pptx::Error::Opc(error)) => opc(error),
        #[cfg(feature = "xlsx")]
        Source::Xlsx(litchi_xlsx::Error::Package(error)) => opc(error),
        _ => None,
    }
}

/// Assert `error` reports a changed source opened at `expected`.
#[track_caller]
#[allow(
    dead_code,
    reason = "used by the source-change suites of whichever formats are enabled"
)]
pub(crate) fn assert_source_changed(error: &Error, expected: litchi_core::SourceVersion) {
    assert_eq!(error.code(), crate::Code::IoError, "{error:?}");
    let (reported, observed) =
        source_change(error).unwrap_or_else(|| panic!("expected a typed source change: {error:?}"));
    assert_eq!(reported, expected);
    assert_ne!(observed, expected);
}
