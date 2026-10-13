use super::*;
use crate::Code;
use crate::detect::{DetectedFormat, detect_format_smart};
use std::path::{Path, PathBuf};

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

const FIXTURES: [&str; 3] = [
    "odf/odt/table-cell-column-span.odt",
    "odf/corpus/writer-user-fields.odt",
    "odf/corpus/writer-paragraph-styles.odt",
];

/// The facade's byte route: smart detection hands back the prepared package.
fn open_prepared(bytes: Vec<u8>) -> Odt {
    match detect_format_smart(bytes) {
        Some(DetectedFormat::Odt(prepared)) => {
            Odt::from_prepared(prepared, &OpenOptions::default()).expect("open prepared ODT")
        },
        other => panic!("expected ODT detection, got {other:?}"),
    }
}

/// The path route: [`crate::open`] classifies the package by its MIME and
/// builds the source-backed owner from the source-pinned candidate.
#[cfg(any(unix, windows))]
fn open_source(path: &Path) -> Odt {
    match crate::open(path).expect("open ODT path") {
        crate::File::Odt(handle) => handle,
        other => panic!("expected ODT, got {other:?}"),
    }
}

fn assert_matches_leaf(odt: &Odt, leaf: &litchi_odt::Document, name: &str) {
    assert_eq!(
        odt.text().expect("text"),
        leaf.text().expect("leaf text"),
        "{name}"
    );
    assert_eq!(
        odt.paragraph_count().expect("paragraph count"),
        leaf.paragraph_count().expect("leaf paragraph count"),
        "{name}"
    );
    assert_eq!(
        format!("{:?}", odt.metadata().expect("metadata").expect("Some")),
        format!("{:?}", leaf.metadata().expect("leaf metadata")),
        "{name}"
    );
    assert_eq!(
        &odt.tolerance_report().expect("report"),
        leaf.tolerance_report(),
        "{name}"
    );
}

#[test]
fn handle_is_one_pointer_and_shares_its_state() {
    assert_eq!(size_of::<Odt>(), size_of::<usize>());
    let odt = open_prepared(std::fs::read(fixture(FIXTURES[0])).expect("read"));
    let clone = odt.clone();
    assert!(Arc::ptr_eq(&odt.0, &clone.0));
    assert_eq!(format!("{odt:?}"), r#"Odt { owner: "eager", .. }"#);
}

#[test]
fn prepared_route_matches_the_leaf_reader() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).expect("read ODT fixture");
        let leaf = litchi_odt::Package::from_bytes(&bytes)
            .expect("leaf ODT")
            .document()
            .expect("leaf ODT");
        let odt = open_prepared(bytes);
        assert!(matches!(odt.owner(), OdtOwner::Eager(_)));
        assert_matches_leaf(&odt, &leaf, name);
    }
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_matches_the_leaf_reader() {
    for name in FIXTURES {
        let path = fixture(name);
        let leaf = litchi_odt::Package::from_bytes(&std::fs::read(&path).expect("read"))
            .expect("leaf")
            .document()
            .expect("leaf");
        let odt = open_source(&path);
        assert!(matches!(odt.owner(), OdtOwner::Source(_)));
        assert_eq!(format!("{odt:?}"), r#"Odt { owner: "source", .. }"#);
        // The source owner records body repairs on first read; read first so
        // both reports cover the same work.
        assert_eq!(
            odt.text().expect("text"),
            leaf.text().expect("leaf text"),
            "{name}"
        );
        assert_eq!(
            odt.paragraph_count().expect("paragraph count"),
            leaf.paragraph_count().expect("leaf paragraph count"),
            "{name}"
        );
        assert_eq!(
            format!("{:?}", odt.metadata().expect("metadata").expect("Some")),
            format!("{:?}", leaf.metadata().expect("leaf metadata")),
            "{name}"
        );
        odt.owner()
            .check_source()
            .expect("unchanged source is current");
    }
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_reports_a_changed_source_as_io_error() {
    let temporary = tempfile::Builder::new()
        .suffix(".odt")
        .tempfile()
        .expect("temporary ODT");
    let bytes = std::fs::read(fixture(FIXTURES[0])).expect("read");
    std::fs::write(temporary.path(), &bytes).expect("write ODT");
    let odt = open_source(temporary.path());

    let mut changed = bytes;
    changed.extend_from_slice(b"trailing bytes change the source revision");
    std::fs::write(temporary.path(), changed).expect("rewrite ODT");

    for error in [
        odt.text().expect_err("text after change"),
        odt.paragraph_count().expect_err("count after change"),
        odt.metadata().expect_err("metadata after change"),
        odt.owner().check_source().expect_err("check after change"),
    ] {
        assert_eq!(error.code(), Code::IoError, "{error}");
        assert!(matches!(error.leaf(), crate::error::Source::Core(_)));
        assert_eq!(error.format(), Some(FileFormat::Odt));
    }
}

/// The facade's `malformed_odt` fixture: a valid ODT package whose body is
/// truncated XML.
fn malformed_odt() -> Vec<u8> {
    use std::io::Write;

    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("mimetype", options).expect("mimetype");
    writer
        .write_all(litchi_odf_common::constants::ODF_TEXT.as_bytes())
        .expect("write mimetype");
    writer
        .start_file("META-INF/manifest.xml", options)
        .expect("manifest");
    writer
        .write_all(
            br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#,
        )
        .expect("write manifest");
    writer.start_file("content.xml", options).expect("content");
    writer
        .write_all(b"<office:document-content>")
        .expect("write content");
    writer.finish().expect("finish ZIP").into_inner()
}

#[test]
fn malformed_body_is_a_typed_core_error() {
    let prepared = match detect_format_smart(malformed_odt()) {
        Some(DetectedFormat::Odt(prepared)) => prepared,
        other => panic!("expected ODT detection, got {other:?}"),
    };
    let error = match Odt::from_prepared(prepared, &OpenOptions::default()) {
        Ok(odt) => odt.text().expect_err("malformed body must not read"),
        Err(error) => error,
    };
    assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
    assert!(
        matches!(error.leaf(), crate::error::Source::Core(_)),
        "{error:?}"
    );
    assert_eq!(error.format(), Some(FileFormat::Odt));
}
