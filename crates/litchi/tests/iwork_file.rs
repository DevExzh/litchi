//! Apple iWork through the unified `File` (record 0306): the behaviours the
//! deleted read-only `litchi::iwork::Document` was tested for, carried over
//! to `litchi::open` — package-directory parity, frozen captures, loose
//! `Index/` bundles, typed path failures, inclusive limits and the Numbers
//! compatibility fixture.

#![cfg(feature = "iwork")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use litchi::view::Mode;
use litchi::{Code, File, FileFormat, OpenOptions};

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(relative)
}

fn outline(file: &File) -> String {
    file.view(Mode::Outline).expect("outline")
}

#[test]
fn packaged_files_and_package_directories_read_alike() {
    for (file, directory, format) in [
        (
            "pages/basic.pages",
            "directory/pages/basic.pages",
            FileFormat::Pages,
        ),
        (
            "numbers/basic.numbers",
            "directory/numbers/basic.numbers",
            FileFormat::Numbers,
        ),
        (
            "keynote/basic.key",
            "directory/keynote/basic.key",
            FileFormat::Keynote,
        ),
    ] {
        let packaged = litchi::open(fixture(file)).expect(file);
        let unpacked = litchi::open(fixture(directory)).expect(directory);
        assert_eq!(packaged.format(), Some(format));
        assert_eq!(unpacked.format(), Some(format));
        assert_eq!(packaged.text().unwrap(), unpacked.text().unwrap(), "{file}");
        assert_eq!(outline(&packaged), outline(&unpacked), "{file}");
    }
    let keynote = litchi::open(fixture("directory/keynote/basic.key")).unwrap();
    assert_eq!(
        keynote
            .get("/slide[1]/placeholder[title]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("Litchi native Keynote fixture")
    );
}

#[test]
fn a_captured_directory_survives_removal_of_its_source() {
    let temp = tempfile::tempdir().unwrap();
    let package = temp.path().join("detached.pages");
    fs::create_dir(&package).unwrap();
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        package.join("Index.zip"),
    )
    .unwrap();
    let file = litchi::open(&package).expect("a bundle with only Index.zip opens");
    fs::remove_dir_all(&package).unwrap();
    assert_eq!(file.format(), Some(FileFormat::Pages));
    assert_eq!(
        file.text().unwrap(),
        litchi::open(fixture("pages/basic.pages"))
            .unwrap()
            .text()
            .unwrap()
    );
    assert_eq!(
        file.get("/section[2]").map(|_| ()).unwrap_err().code(),
        Code::PathNotFound,
        "one section"
    );
}

/// Unpack `Index.zip` of the Pages bundle fixture into a loose `Index/`.
fn loose_bundle(package: &Path) {
    let index = package.join("Index");
    fs::create_dir_all(&index).unwrap();
    let mut archive = zip::ZipArchive::new(
        fs::File::open(fixture("directory/pages/basic.pages/Index.zip")).unwrap(),
    )
    .unwrap();
    for position in 0..archive.len() {
        let mut entry = archive.by_index(position).unwrap();
        let name = entry
            .name()
            .strip_prefix("Index/")
            .filter(|name| !name.is_empty() && !name.contains('/'))
            .expect("flat Index member")
            .to_owned();
        io::copy(&mut entry, &mut fs::File::create(index.join(name)).unwrap()).unwrap();
    }
}

#[test]
fn a_loose_index_directory_reads_like_the_packaged_document() {
    let temp = tempfile::tempdir().unwrap();
    let package = temp.path().join("loose.pages");
    loose_bundle(&package);
    let loose = litchi::open(&package).expect("a loose Index directory opens");
    let packaged = litchi::open(fixture("pages/basic.pages")).unwrap();
    assert_eq!(loose.format(), Some(FileFormat::Pages));
    assert_eq!(loose.text().unwrap(), packaged.text().unwrap());
    assert_eq!(outline(&loose), outline(&packaged));
    // It edits, and `save()` writes the `Index.zip` bundle layout back.
    let edited = loose
        .set("/body/p[1]", litchi::props! { "text" => "Loose edit" })
        .unwrap();
    edited.save().unwrap();
    assert!(package.join("Index.zip").is_file());
    assert!(!package.join("Index").exists());
    let reopened = litchi::open(&package).unwrap();
    assert_eq!(
        reopened
            .get("/body/p[1]")
            .unwrap()
            .view()
            .text()
            .unwrap()
            .as_deref(),
        Some("Loose edit")
    );
}

#[test]
fn path_failures_are_typed() {
    let temp = tempfile::tempdir().unwrap();
    let missing = litchi::open(temp.path().join("missing.pages")).unwrap_err();
    assert_eq!(missing.code(), Code::FileNotFound, "{missing:?}");

    let unrelated = temp.path().join("unrelated.pages");
    fs::write(&unrelated, b"not an iWork package").unwrap();
    let error = litchi::open(&unrelated).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), None);

    // Both index layouts, or a pre-2013 XML marker beside `Index.zip`, are
    // ambiguous: refused, not guessed.
    let mixed = temp.path().join("mixed.pages");
    fs::create_dir(&mixed).unwrap();
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        mixed.join("Index.zip"),
    )
    .unwrap();
    fs::create_dir(mixed.join("Index")).unwrap();
    assert!(litchi::open(&mixed).is_err());

    let marker = temp.path().join("marker-conflict.pages");
    fs::create_dir(&marker).unwrap();
    fs::copy(
        fixture("directory/pages/basic.pages/Index.zip"),
        marker.join("Index.zip"),
    )
    .unwrap();
    fs::write(marker.join("index.apxl"), []).unwrap();
    assert!(litchi::open(&marker).is_err());
}

#[cfg(unix)]
#[test]
fn special_nodes_are_refused_without_blocking() {
    use std::os::unix::net::UnixListener;

    // Unix socket paths are bounded by SUN_LEN; keep this fixture's parent
    // under /tmp even when the test harness has a long TMPDIR.
    let temp = tempfile::tempdir_in("/tmp").unwrap();
    let socket = temp.path().join("socket.pages");
    let _listener = UnixListener::bind(&socket).unwrap();
    assert!(litchi::open(&socket).is_err());
}

#[test]
fn directory_input_limits_are_inclusive_and_content_free() {
    let directory = fixture("directory/pages/basic.pages");
    let length = fs::read_dir(&directory)
        .unwrap()
        .flatten()
        .map(|entry| {
            let path = entry.path();
            if path.is_dir() {
                fs::read_dir(path)
                    .unwrap()
                    .flatten()
                    .map(|entry| entry.metadata().unwrap().len())
                    .sum::<u64>()
            } else {
                entry.metadata().unwrap().len()
            }
        })
        .sum::<u64>();
    let options = |limit: u64| {
        OpenOptions::new().pages_options(
            litchi::pages::OpenOptions::new().source_limits(
                litchi::pages::SourceLimits::builder()
                    .max_input_bytes(limit)
                    .unwrap()
                    .build()
                    .unwrap(),
            ),
        )
    };
    litchi::open_with(&directory, &options(length)).expect("the exact input limit passes");
    let error = litchi::open_with(&directory, &options(length - 1)).unwrap_err();
    assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
    assert!(!error.to_string().contains("basic.pages"));
}

/// The synthetic Numbers compatibility fixture: `File` reads the rooted
/// workbook's text (one sheet, its two canonical tables in drawable order,
/// the typed number). The deleted snapshot also listed a table model
/// detached from every sheet; detached models are not document content, and
/// no `File` element addresses them (record 0306).
#[test]
fn the_numbers_compatibility_fixture_reads_its_rooted_tables() {
    let hex = include_str!("../../../test-data/synthetic-iwork/numbers/compatibility-oracles.hex");
    let digits: Vec<u8> = hex
        .bytes()
        .filter(|byte| !byte.is_ascii_whitespace())
        .collect();
    let bytes: Vec<u8> = digits
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect();
    let file = litchi::open_bytes(bytes).unwrap();
    assert_eq!(file.format(), Some(FileFormat::Numbers));
    assert_eq!(
        file.text().unwrap(),
        "Rooted sheet\nSecond canonical\nFirst canonical\n-1234.5"
    );
    // The object tree refuses this synthetic source, typed: its table-merge
    // storage is not a native shape (follow-up in record 0306).
    let error = file.query("table").map(|_| ()).unwrap_err();
    assert_eq!(error.code(), Code::InvalidInput, "{error:?}");
}

/// Markdown goes through `File` (record 0306): Pages paragraphs, one
/// `# Slide N` section per Keynote slide; a Numbers workbook is refused
/// like every workbook.
#[cfg(feature = "markdown")]
#[test]
fn pages_and_keynote_convert_to_markdown_through_file() {
    use litchi::markdown::ToMarkdown;

    let pages = litchi::open(fixture("pages/basic.pages")).unwrap();
    assert_eq!(
        pages.to_markdown().unwrap(),
        "Litchi native Pages fixture\n\nBuffa lazy\\-view migration verification\n\n2026\\-08\\-07\n\n"
    );
    let directory = litchi::open(fixture("directory/pages/basic.pages")).unwrap();
    assert_eq!(
        directory.to_markdown().unwrap(),
        pages.to_markdown().unwrap()
    );

    let keynote = litchi::open(fixture("keynote/basic.key")).unwrap();
    let markdown = keynote.to_markdown().unwrap();
    assert!(
        markdown.starts_with("# Slide 1 Litchi native Keynote fixture\n\n"),
        "{markdown}"
    );
    assert!(
        markdown.contains("Buffa lazy\\-view migration verification"),
        "{markdown}"
    );

    let numbers = litchi::open(fixture("numbers/basic.numbers")).unwrap();
    assert_eq!(
        numbers.to_markdown().unwrap_err().code(),
        Code::UnsupportedCapability
    );
}
