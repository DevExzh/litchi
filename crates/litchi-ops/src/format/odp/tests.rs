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
    "odf/corpus/impress-basic.odp",
    "odf/corpus/impress-master-layouts.odp",
    "odf/corpus/impress-embedded-spreadsheet.odp",
];

/// The facade's first byte-route arm: checked ODP preparation.
fn open_prepared(bytes: Vec<u8>) -> Odp {
    let prepared = crate::detect::detected::detect_prepared_odp(bytes)
        .unwrap_or_else(|_| panic!("expected a prepared ODP package"));
    Odp::from_prepared(prepared, &OpenOptions::default()).expect("open prepared ODP")
}

/// The facade's `from_detected` arm, which re-parsed the bytes detection
/// retained: the unified byte route adopts the prepared package instead.
fn open_detected(bytes: Vec<u8>) -> Odp {
    assert!(matches!(
        detect_format_smart(bytes.clone()),
        Some(DetectedFormat::Odp(_))
    ));
    match crate::open_bytes(bytes).expect("open detected ODP") {
        crate::File::Odp(handle) => handle,
        other => panic!("expected ODP, got {other:?}"),
    }
}

/// The facade's path route: a presentation path probe builds the source
/// owner (the PPTX-aware probe when PPTX is enabled).
#[cfg(any(unix, windows))]
fn open_source(path: &Path) -> Odp {
    #[cfg(feature = "pptx")]
    let presentation = match crate::detect::detected::detect_pptx_source_path_with_limits(
        path,
        crate::opc::ReadLimits::default(),
    )
    .expect("probe ODP path")
    {
        Some(crate::detect::detected::PptxSourcePathDetection::Odp(presentation)) => presentation,
        _ => panic!("expected an ODP source owner"),
    };
    #[cfg(not(feature = "pptx"))]
    let presentation = {
        // Builds without PPTX but with another OOXML leaf keep the eager
        // fallback in the probe, so build the same leaf owner directly.
        let source: Arc<dyn litchi_core::ReadAt> =
            Arc::new(litchi_core::FileSource::open(path).expect("open source"));
        litchi_odp::SourceBackedPresentation::from_read_at(source).expect("ODP source owner")
    };
    Odp::from_source(presentation, &OpenOptions::default()).expect("open ODP source")
}

fn assert_matches_leaf(odp: &Odp, leaf: &litchi_odp::Presentation, name: &str) {
    assert_eq!(
        odp.text().expect("text"),
        leaf.text().expect("leaf text"),
        "{name}"
    );
    assert_eq!(
        odp.slide_count().expect("slide count"),
        leaf.slide_count().expect("leaf slide count"),
        "{name}"
    );
    // The facade reports `Some(Metadata::default())` for ODP.
    assert_eq!(
        format!("{:?}", odp.metadata().expect("metadata")),
        format!("{:?}", Some(Metadata::default())),
        "{name}"
    );
    assert_eq!(
        &odp.tolerance_report().expect("report"),
        leaf.tolerance_report(),
        "{name}"
    );
}

#[test]
fn handle_is_one_pointer_and_shares_its_state() {
    assert_eq!(size_of::<Odp>(), size_of::<usize>());
    let odp = open_prepared(std::fs::read(fixture(FIXTURES[0])).expect("read"));
    let clone = odp.clone();
    assert!(Arc::ptr_eq(&odp.0, &clone.0));
    assert_eq!(format!("{odp:?}"), r#"Odp { owner: "eager", .. }"#);
}

#[test]
fn byte_routes_match_the_leaf_reader() {
    for name in FIXTURES {
        let bytes = std::fs::read(fixture(name)).expect("read ODP fixture");
        let leaf = litchi_odp::Presentation::from_bytes(bytes.clone()).expect("leaf ODP");
        assert!(leaf.slide_count().expect("slides") > 0, "{name}");
        for odp in [open_prepared(bytes.clone()), open_detected(bytes)] {
            assert!(matches!(odp.owner(), OdpOwner::Eager(_)));
            assert_matches_leaf(&odp, &leaf, name);
        }
    }
}

/// The facade's source-path fixture (`source_odp_path_tests::package`): a
/// two-slide package the ordinary-ODP path probe accepts.
fn source_package() -> Vec<u8> {
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page draw:name="one"><draw:frame><draw:text-box><text:p>root source</text:p></draw:text-box></draw:frame></draw:page><draw:page draw:name="two"><draw:frame><draw:text-box><text:p>second slide</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
    );
    let mut writer = litchi_odf_common::core::PackageWriter::new();
    writer
        .set_mimetype(litchi_odf_common::constants::ODF_PRESENTATION)
        .expect("ODP MIME");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("ODP content");
    writer.finish_to_bytes().expect("ODP package")
}

#[cfg(any(unix, windows))]
fn write_source_package() -> (tempfile::NamedTempFile, Vec<u8>) {
    let bytes = source_package();
    let temporary = tempfile::Builder::new()
        .suffix(".odp")
        .tempfile()
        .expect("temporary ODP");
    std::fs::write(temporary.path(), &bytes).expect("write ODP");
    (temporary, bytes)
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_matches_the_leaf_reader() {
    let (temporary, bytes) = write_source_package();
    let leaf = litchi_odp::Presentation::from_bytes(bytes).expect("leaf ODP");
    let odp = open_source(temporary.path());
    assert!(matches!(odp.owner(), OdpOwner::Source(_)));
    assert_eq!(format!("{odp:?}"), r#"Odp { owner: "source", .. }"#);
    assert_eq!(odp.slide_count().expect("slide count"), 2);
    assert_matches_leaf(&odp, &leaf, "synthetic source ODP");
}

#[cfg(any(unix, windows))]
#[test]
fn source_route_reports_a_changed_source_as_io_error() {
    let (temporary, bytes) = write_source_package();
    let odp = open_source(temporary.path());

    let mut changed = bytes;
    changed.extend_from_slice(b"trailing bytes change the source revision");
    std::fs::write(temporary.path(), changed).expect("rewrite ODP");

    for error in [
        odp.text().expect_err("text after change"),
        odp.slide_count().expect_err("count after change"),
        odp.metadata().expect_err("metadata after change"),
    ] {
        assert_eq!(error.code(), Code::IoError, "{error}");
        assert!(matches!(error.leaf(), crate::error::Source::Core(_)));
        assert_eq!(error.format(), Some(FileFormat::Odp));
    }
}

#[test]
fn malformed_package_is_a_typed_core_error() {
    let bytes = broken_package(litchi_odf_common::constants::ODF_PRESENTATION);
    let Err(leaf_error) = litchi_odp::Presentation::from_bytes(bytes.clone()) else {
        panic!("a broken package must not open in the leaf");
    };
    let error = crate::open_bytes(bytes).expect_err("a broken package must not open");
    assert!(
        matches!(error.leaf(), crate::error::Source::Core(_)),
        "{error:?}"
    );
    assert_eq!(error.format(), Some(FileFormat::Odp));
    assert_eq!(error.message(), leaf_error.to_string());
}

/// A package with a valid `mimetype` and manifest and an unparsable body.
fn broken_package(mimetype: &str) -> Vec<u8> {
    use std::io::Write as _;

    let mut writer = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    writer.start_file("mimetype", options).unwrap();
    writer.write_all(mimetype.as_bytes()).unwrap();
    writer.start_file("META-INF/manifest.xml", options).unwrap();
    writer
        .write_all(
            format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="{mimetype}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer.start_file("content.xml", options).unwrap();
    writer.write_all(b"<office:document-content>").unwrap();
    writer.finish().unwrap().into_inner()
}
