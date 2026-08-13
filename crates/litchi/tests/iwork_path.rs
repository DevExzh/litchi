#![cfg(feature = "iwork")]

use std::fs;
use std::io;
use std::path::PathBuf;

use litchi::iwork::{Document, ErrorKind, Format, Options, Resource, SourceLimits, Stage};
use litchi_keynote::Package as KeynotePackage;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(relative)
}

fn assert_path_parity(file: &str, directory: &str, expected: Format) {
    let packaged = Document::open(fixture(file))
        .unwrap_or_else(|error| panic!("packaged native fixture must open: {error}"));
    let unpacked = Document::open(fixture(directory))
        .unwrap_or_else(|error| panic!("directory native fixture must open: {error}"));

    assert_eq!(packaged.format(), expected);
    assert_eq!(unpacked.format(), expected);
    assert_eq!(packaged.snapshot().summary(), unpacked.snapshot().summary());
    assert_eq!(
        packaged.snapshot().all_text(),
        unpacked.snapshot().all_text()
    );
}

#[test]
fn packaged_files_and_native_directories_have_semantic_parity() {
    assert_path_parity(
        "pages/basic.pages",
        "directory/pages/basic.pages",
        Format::Pages,
    );
    assert_path_parity(
        "numbers/basic.numbers",
        "directory/numbers/basic.numbers",
        Format::Numbers,
    );
    assert_path_parity(
        "keynote/basic.key",
        "directory/keynote/basic.key",
        Format::Keynote,
    );
}

#[test]
fn keynote_directory_coordinator_matches_the_focused_zip_snapshot() {
    let directory = Document::open(fixture("directory/keynote/basic.key"))
        .unwrap_or_else(|error| panic!("directory Keynote fixture must open: {error}"));
    let focused = KeynotePackage::open(fixture("keynote/basic.key"))
        .unwrap_or_else(|error| panic!("focused Keynote fixture must open: {error}"));
    let snapshot = directory.snapshot();

    assert_eq!(directory.format(), Format::Keynote);
    assert_eq!(snapshot.format(), Format::Keynote);
    assert_eq!(snapshot.slide_count(), 1);
    assert_eq!(snapshot.summary().slides(), focused.slides().unwrap().len());
    assert_eq!(
        snapshot.all_text(),
        focused
            .show()
            .unwrap()
            .all_text()
            .into_iter()
            .collect::<Vec<_>>()
    );

    let slide = snapshot
        .slide(0)
        .unwrap_or_else(|| panic!("directory Keynote snapshot must have one slide"));
    let focused_slide = &focused.slides().unwrap()[0];
    assert_eq!(slide.position(), focused_slide.index());
    assert_eq!(slide.is_skipped(), focused_slide.is_skipped());
    assert_eq!(slide.name(), focused_slide.name());
    assert_eq!(slide.title(), focused_slide.title());
    assert_eq!(slide.build_count(), focused_slide.builds().len());
    assert_eq!(slide.has_transition(), focused_slide.transition().is_some());
    assert_eq!(
        slide
            .iter_text()
            .map(|text| text.value())
            .collect::<Vec<_>>(),
        [
            "Litchi native Keynote fixture",
            "Buffa lazy-view migration verification",
            "2026-08-07",
        ]
    );
}

#[test]
fn frozen_directory_semantics_survive_source_removal() {
    let temp = tempfile::tempdir().unwrap_or_else(|error| panic!("create temp directory: {error}"));
    let package = temp.path().join("detached.pages");
    fs::create_dir(&package).unwrap_or_else(|error| panic!("create package directory: {error}"));
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        package.join("Index.zip"),
    )
    .unwrap_or_else(|error| panic!("copy native index: {error}"));

    let document = Document::open(&package)
        .unwrap_or_else(|error| panic!("frozen directory must open: {error}"));
    fs::remove_dir_all(&package).unwrap_or_else(|error| panic!("remove captured source: {error}"));

    assert_eq!(document.format(), Format::Pages);
    assert_eq!(document.snapshot().summary().sections(), 1);
    assert_eq!(
        document.snapshot().all_text(),
        [
            "Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07"
                .to_owned()
        ]
    );
}

#[test]
fn loose_index_directory_uses_the_same_pages_snapshot() {
    let temp = tempfile::tempdir().unwrap_or_else(|error| panic!("create temp directory: {error}"));
    let package = temp.path().join("loose.pages");
    let index = package.join("Index");
    fs::create_dir_all(&index).unwrap_or_else(|error| panic!("create loose index: {error}"));

    let file = fs::File::open(fixture("directory/pages/basic.pages/Index.zip"))
        .unwrap_or_else(|error| panic!("open native index archive: {error}"));
    let mut archive = zip::ZipArchive::new(file)
        .unwrap_or_else(|error| panic!("parse native index archive: {error}"));
    for position in 0..archive.len() {
        let mut entry = archive
            .by_index(position)
            .unwrap_or_else(|error| panic!("read native index member {position}: {error}"));
        let name = entry
            .name()
            .strip_prefix("Index/")
            .filter(|name| !name.is_empty() && !name.contains('/'))
            .unwrap_or_else(|| panic!("Pages native index member must be flat and normalized"));
        let mut output = fs::File::create(index.join(name))
            .unwrap_or_else(|error| panic!("create loose index member: {error}"));
        io::copy(&mut entry, &mut output)
            .unwrap_or_else(|error| panic!("copy loose index member: {error}"));
    }

    let loose = Document::open(&package)
        .unwrap_or_else(|error| panic!("loose directory must open: {error}"));
    let packaged = Document::open(fixture("pages/basic.pages"))
        .unwrap_or_else(|error| panic!("packaged Pages fixture must open: {error}"));
    assert_eq!(loose.format(), Format::Pages);
    assert_eq!(loose.snapshot().summary(), packaged.snapshot().summary());
    assert_eq!(loose.snapshot().all_text(), packaged.snapshot().all_text());
}

#[test]
fn path_failures_are_content_free_and_typed() {
    let temp = tempfile::tempdir().unwrap_or_else(|error| panic!("create temp directory: {error}"));
    let missing = temp.path().join("missing.pages");
    let error = Document::open(&missing)
        .err()
        .unwrap_or_else(|| panic!("missing source must fail"));
    assert_eq!(error.kind(), ErrorKind::Io);
    assert_eq!(error.stage(), Stage::Input);

    let unrelated = temp.path().join("unrelated.pages");
    fs::write(&unrelated, b"not an iWork package")
        .unwrap_or_else(|error| panic!("write unrelated source: {error}"));
    let error = Document::open(&unrelated)
        .err()
        .unwrap_or_else(|| panic!("unrelated source must fail"));
    assert_eq!(error.kind(), ErrorKind::Unrecognized);
    assert_eq!(error.stage(), Stage::Detection);

    let mixed = temp.path().join("mixed.pages");
    fs::create_dir(&mixed).unwrap_or_else(|error| panic!("create mixed package: {error}"));
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        mixed.join("Index.zip"),
    )
    .unwrap_or_else(|error| panic!("copy mixed index: {error}"));
    fs::create_dir(mixed.join("Index"))
        .unwrap_or_else(|error| panic!("create conflicting index: {error}"));
    let error = Document::open(&mixed)
        .err()
        .unwrap_or_else(|| panic!("mixed index representations must fail"));
    assert_eq!(error.kind(), ErrorKind::InvalidData);
    assert_eq!(error.stage(), Stage::Detection);

    let marker_conflict = temp.path().join("marker-conflict.pages");
    fs::create_dir(&marker_conflict)
        .unwrap_or_else(|error| panic!("create marker-conflict package: {error}"));
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        marker_conflict.join("Index.zip"),
    )
    .unwrap_or_else(|error| panic!("copy marker-conflict index: {error}"));
    fs::write(marker_conflict.join("index.apxl"), [])
        .unwrap_or_else(|error| panic!("write conflicting marker: {error}"));
    let error = Document::open(&marker_conflict)
        .err()
        .unwrap_or_else(|| panic!("conflicting marker must fail"));
    assert_eq!(error.kind(), ErrorKind::InvalidData);
    assert_eq!(error.stage(), Stage::Detection);
}

#[cfg(unix)]
#[test]
fn path_ingress_rejects_links_and_special_nodes_without_blocking() {
    use std::os::unix::fs::symlink;
    use std::os::unix::net::UnixListener;

    let temp = tempfile::tempdir().unwrap_or_else(|error| panic!("create temp directory: {error}"));
    let link = temp.path().join("linked.pages");
    symlink(fixture("pages/basic.pages"), &link)
        .unwrap_or_else(|error| panic!("create source symlink: {error}"));
    let error = Document::open(&link)
        .err()
        .unwrap_or_else(|| panic!("symbolic source must fail"));
    assert_eq!(error.kind(), ErrorKind::InvalidData);

    let socket = temp.path().join("socket.pages");
    let _listener = UnixListener::bind(&socket)
        .unwrap_or_else(|error| panic!("create special source node: {error}"));
    let error = Document::open(&socket)
        .err()
        .unwrap_or_else(|| panic!("special source node must fail"));
    assert_eq!(error.kind(), ErrorKind::InvalidData);
}

#[test]
fn directory_index_input_limit_is_inclusive() {
    let directory = fixture("directory/pages/basic.pages");
    let length = fs::metadata(directory.join("Index.zip"))
        .unwrap_or_else(|error| panic!("read native index metadata: {error}"))
        .len();

    Document::open_with_options(&directory, options_with_input_limit(length))
        .unwrap_or_else(|error| panic!("exact directory input limit must pass: {error}"));
    let error = Document::open_with_options(&directory, options_with_input_limit(length - 1))
        .err()
        .unwrap_or_else(|| panic!("one-under directory input limit must fail"));
    assert_eq!(error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(error.stage(), Stage::Detection);
    assert_eq!(error.resource(), Some(Resource::InputBytes));
    assert_eq!(error.observed(), Some(length));
    assert_eq!(error.maximum(), Some(length - 1));
}

fn options_with_input_limit(max_input_bytes: u64) -> Options {
    let source = SourceLimits::new(
        max_input_bytes,
        SourceLimits::HARD_MAX_ENTRIES,
        SourceLimits::HARD_MAX_ENTRY_BYTES,
        SourceLimits::HARD_MAX_EXPANDED_BYTES,
        SourceLimits::HARD_MAX_DECODED_BYTES_PER_ITEM,
    )
    .unwrap_or_else(|error| panic!("path source limits must be valid: {error}"));
    Options::default().with_source(source)
}
