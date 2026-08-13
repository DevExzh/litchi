use std::path::PathBuf;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_keynote::{
    Document, DocumentReadOptions, DocumentSourceLimits, Package, ReadOptions, SemanticLimitKind,
    SemanticLimits, SlideSelector, Stats, show::Mode,
};

const EXPECTED_TEXT: &str =
    "Litchi native Keynote fixture\nBuffa lazy-view migration verification\n2026-08-07";
const EXPECTED_OBJECTS: usize = 959;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

fn directory_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/directory/keynote/basic.key")
}

fn pages_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

fn pages_directory_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/directory/pages/basic.pages")
}

fn assert_expected_slide(package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    package.validate()?;
    let show = package.show()?;
    assert_eq!(show.slide_count(), 1);
    let selected = show
        .select_slide(SlideSelector::index(0))?
        .ok_or_else(|| std::io::Error::other("native Keynote fixture has no first slide"))?;
    assert_eq!(selected.index(), 0);
    if let Some(name) = selected.name() {
        assert_eq!(show.select_slide(name)?, Some(selected));
    }

    let slide = &show.slides()[0];
    assert!(!slide.is_skipped());
    assert_eq!(slide.name(), None);
    assert_eq!(slide.title(), Some("Litchi native Keynote fixture"));
    assert!(slide.text_content().is_empty());
    assert_eq!(slide.notes(), None);
    assert!(slide.builds().is_empty());
    let transition = slide
        .transition()
        .ok_or_else(|| std::io::Error::other("native slide has no transition"))?;
    assert_eq!(transition.duration().as_f64(), 1.0);
    assert_eq!(
        slide
            .text_storages()
            .iter()
            .map(|storage| storage.text())
            .collect::<Vec<_>>(),
        ["Buffa lazy-view migration verification", "2026-08-07"]
    );
    assert_eq!(slide.plain_text(), EXPECTED_TEXT);
    assert_eq!(show.all_text().join("\n"), EXPECTED_TEXT);
    assert_eq!(package.text()?, EXPECTED_TEXT);
    assert_eq!(package.slides()?, show.slides());

    assert_eq!(show.title(), None);
    let settings = *show.settings();
    assert_eq!(settings.size().width(), 1_920.0);
    assert_eq!(settings.size().height(), 1_080.0);
    assert_eq!(settings.slide_numbers_visible(), None);
    assert_eq!(settings.loop_presentation(), Some(false));
    assert_eq!(settings.mode(), Some(Mode::Normal));
    assert_eq!(
        settings
            .autoplay_transition_delay()
            .map(|value| value.as_f64()),
        Some(5.0)
    );
    assert_eq!(
        settings.autoplay_build_delay().map(|value| value.as_f64()),
        Some(2.0)
    );
    assert_eq!(settings.idle_timer_active(), Some(false));
    assert_eq!(
        settings.idle_timer_delay().map(|value| value.as_f64()),
        Some(900.0)
    );
    assert_eq!(settings.automatically_plays_upon_open(), Some(false));

    let stats = package.stats()?;
    assert_eq!(
        stats,
        Stats {
            total_objects: EXPECTED_OBJECTS,
            slide_count: 1,
        }
    );
    Ok(())
}

#[test]
fn native_keynote_fixture_opens_from_path_and_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let package = Package::open(&path)?;
    assert_expected_slide(&package)?;

    let bytes = std::fs::read(path)?;
    let from_bytes = Package::from_bytes(&bytes)?;
    assert_expected_slide(&from_bytes)?;
    assert_eq!(from_bytes.show()?, package.show()?);
    Ok(())
}

#[test]
fn focused_reader_surface_is_exact_shareable_and_archive_free()
-> Result<(), Box<dyn std::error::Error>> {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Package>();
    assert_send_sync::<Document>();
    assert_send_sync::<Stats>();

    let bytes = std::fs::read(fixture_path())?;
    let package = {
        let borrowed_source = bytes.clone();
        Package::from_bytes(&borrowed_source)?
    };
    let snapshot = package.snapshot();
    package.validate()?;
    snapshot.validate()?;

    let mut exact = Vec::new();
    snapshot.write_to(&mut exact)?;
    assert_eq!(exact, bytes);
    assert_eq!(snapshot.read_options(), package.read_options());
    assert_eq!(snapshot.stats()?, package.stats()?);
    assert_eq!(snapshot.text()?, EXPECTED_TEXT);
    assert_eq!(snapshot.show()?, package.show()?);
    assert_eq!(snapshot.slides()?.as_ptr(), package.slides()?.as_ptr());

    let semantic = package.semantic_snapshot()?;
    let semantic_snapshot = semantic.snapshot();
    assert_eq!(semantic.show(), package.show()?);
    assert_eq!(semantic.slides(), package.slides()?);
    assert_eq!(semantic.slides().as_ptr(), package.slides()?.as_ptr());
    assert_eq!(semantic_snapshot.show(), semantic.show());
    assert_eq!(semantic.stats(), None);
    assert!(semantic.metadata().is_none());
    assert_eq!(
        semantic_snapshot.slides().as_ptr(),
        semantic.slides().as_ptr()
    );

    let metadata = package
        .metadata()?
        .ok_or_else(|| std::io::Error::other("native Keynote fixture has no metadata"))?;
    assert_eq!(metadata.application.as_deref(), Some("Keynote"));
    assert_eq!(metadata.title, None);
    assert_eq!(metadata.author, None);
    assert_eq!(
        metadata.revision.as_deref(),
        Some("0::EFC057DE-5F02-4B6E-B3D4-B5A5C39D430A")
    );
    assert_eq!(
        metadata.content_status.as_deref(),
        Some("Keynote Format Version 14.4.1")
    );

    let worker = snapshot.snapshot();
    let (worker_text, worker_stats) = std::thread::spawn(move || {
        worker.validate()?;
        Ok::<_, litchi_keynote::ReadError>((worker.text()?, worker.stats()?))
    })
    .join()
    .map_err(|_panic| std::io::Error::other("Keynote reader worker panicked"))??;
    assert_eq!(worker_text, EXPECTED_TEXT);
    assert_eq!(worker_stats, package.stats()?);
    Ok(())
}

#[test]
fn focused_package_refuses_directory_backing_without_treating_index_as_complete()
-> Result<(), Box<dyn std::error::Error>> {
    let options = ReadOptions::default();
    let error = Package::open_with_options(directory_fixture_path(), options)
        .expect_err("focused exact-package reader must refuse a directory bundle");
    assert!(matches!(error, litchi_keynote::ReadError::Io(_)));
    Ok(())
}

#[test]
fn metadata_ignores_earlier_malformed_near_names_and_preserves_source()
-> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let catalog = Catalog::from_bytes(&source)?;
    let malformed = b"not a plist";
    let mut entries = Vec::with_capacity(catalog.len().saturating_add(2));
    entries.push(("A/Properties.plist", malformed.as_slice()));
    entries.push(("Properties.plist", malformed.as_slice()));
    entries.extend(catalog.iter().map(|entry| (entry.name(), entry.data())));
    let candidate =
        litchi_iwa_archive::package::to_bytes(entries.iter().copied(), Limits::default())?;

    let package = Package::from_bytes(&candidate)?;
    let metadata = package
        .metadata()?
        .ok_or_else(|| std::io::Error::other("candidate has no canonical metadata"))?;
    assert_eq!(metadata.application.as_deref(), Some("Keynote"));
    assert_eq!(metadata.title, None);
    assert_eq!(metadata.author, None);
    assert_eq!(
        metadata.revision.as_deref(),
        Some("0::EFC057DE-5F02-4B6E-B3D4-B5A5C39D430A")
    );
    assert_eq!(
        metadata.content_status.as_deref(),
        Some("Keynote Format Version 14.4.1")
    );

    let mut exact = Vec::new();
    package.write_to(&mut exact)?;
    assert_eq!(exact, candidate);

    let temp = tempfile::tempdir()?;
    let path = temp.path().join("canonical-metadata.key");
    std::fs::write(&path, candidate)?;
    let document = Document::open(path)?;
    assert_eq!(document.show(), package.show()?);
    let document_metadata = document
        .metadata()
        .ok_or_else(|| std::io::Error::other("document has no canonical metadata"))?;
    assert_eq!(document_metadata.title, metadata.title);
    assert_eq!(document_metadata.author, metadata.author);
    assert_eq!(document_metadata.revision, metadata.revision);
    assert_eq!(document_metadata.content_status, metadata.content_status);
    assert_eq!(document_metadata.application, metadata.application);
    Ok(())
}

#[test]
fn archive_free_document_rejects_other_iwork_formats_before_semantic_publication() {
    for path in [pages_fixture_path(), pages_directory_fixture_path()] {
        assert!(matches!(
            Document::open(path),
            Err(litchi_keynote::ReadError::NotKeynote)
        ));
    }
}

#[test]
fn archive_free_document_has_full_zip_and_directory_semantic_parity()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let zipped = Document::open(fixture_path())?;
    let directory = Document::open(directory_fixture_path())?;

    zipped.validate()?;
    directory.validate()?;
    assert_eq!(zipped.show(), package.show()?);
    assert_eq!(directory.show(), zipped.show());
    assert_eq!(zipped.text()?, EXPECTED_TEXT);
    assert_eq!(directory.text()?, EXPECTED_TEXT);
    assert_eq!(zipped.stats(), Some(package.stats()?));
    assert_eq!(directory.stats(), zipped.stats());
    assert_eq!(
        zipped.snapshot().slides().as_ptr(),
        zipped.slides().as_ptr()
    );
    assert_eq!(
        directory.snapshot().slides().as_ptr(),
        directory.slides().as_ptr()
    );

    let zipped_metadata = zipped
        .metadata()
        .ok_or_else(|| std::io::Error::other("ZIP semantic document has no metadata"))?;
    assert_eq!(zipped_metadata.application.as_deref(), Some("Keynote"));
    assert_eq!(
        zipped_metadata.revision.as_deref(),
        Some("0::EFC057DE-5F02-4B6E-B3D4-B5A5C39D430A")
    );
    let directory_metadata = directory
        .metadata()
        .ok_or_else(|| std::io::Error::other("directory semantic document has no metadata"))?;
    assert_eq!(directory_metadata.application.as_deref(), Some("Keynote"));
    assert_eq!(
        directory_metadata.revision.as_deref(),
        Some("0::D71B2AD6-F14B-4945-AD9A-7ABEC2907DB1")
    );
    assert_eq!(
        directory_metadata.content_status.as_deref(),
        Some("Keynote Format Version 14.4.1")
    );
    Ok(())
}

#[test]
fn archive_free_directory_is_frozen_and_enforces_source_and_semantic_limits()
-> Result<(), Box<dyn std::error::Error>> {
    let temp = tempfile::tempdir()?;
    let source = temp.path().join("detached.key");
    std::fs::create_dir_all(source.join("Metadata"))?;
    std::fs::copy(
        directory_fixture_path().join("Index.zip"),
        source.join("Index.zip"),
    )?;
    std::fs::copy(
        directory_fixture_path().join("Metadata/Properties.plist"),
        source.join("Metadata/Properties.plist"),
    )?;
    std::fs::create_dir(source.join("A"))?;
    std::fs::write(source.join("A/Properties.plist"), b"not a plist")?;
    std::fs::write(source.join("Properties.plist"), b"not a plist")?;

    let index_bytes = std::fs::metadata(source.join("Index.zip"))?.len();
    let properties_bytes = std::fs::metadata(source.join("Metadata/Properties.plist"))?.len();
    let exact_input = index_bytes + properties_bytes;
    let defaults = DocumentSourceLimits::default();
    let exact_source = DocumentSourceLimits::new(
        exact_input,
        defaults.max_files(),
        defaults.max_entry_size(),
        defaults.max_total_size(),
        defaults.max_iwa_stream_size(),
    )?;
    let exact = Document::open_with_options(
        &source,
        DocumentReadOptions::new(exact_source, SemanticLimits::default()),
    )?;
    assert_eq!(exact.text()?, EXPECTED_TEXT);
    assert_eq!(
        exact.stats().map(|stats| stats.total_objects),
        Some(EXPECTED_OBJECTS)
    );

    let too_small_source = DocumentSourceLimits::new(
        exact_input - 1,
        defaults.max_files(),
        defaults.max_entry_size(),
        defaults.max_total_size(),
        defaults.max_iwa_stream_size(),
    )?;
    let error = Document::open_with_options(
        &source,
        DocumentReadOptions::new(too_small_source, SemanticLimits::default()),
    )
    .expect_err("aggregate index plus properties max-minus-one must refuse");
    assert!(matches!(error, litchi_keynote::ReadError::Detection(_)));

    let semantic = SemanticLimits::new(
        EXPECTED_OBJECTS - 1,
        SemanticLimits::MAX_SLIDES,
        SemanticLimits::MAX_REFERENCES,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let error =
        Document::open_with_options(&source, DocumentReadOptions::new(exact_source, semantic))
            .expect_err("object max-minus-one must refuse before publishing a document");
    assert!(matches!(
        error,
        litchi_keynote::ReadError::SemanticLimit {
            kind: SemanticLimitKind::Objects,
            observed: EXPECTED_OBJECTS,
            maximum,
            ..
        } if maximum == EXPECTED_OBJECTS - 1
    ));

    std::fs::remove_dir_all(&source)?;
    assert_eq!(exact.text()?, EXPECTED_TEXT);
    assert_eq!(
        exact
            .metadata()
            .and_then(|metadata| metadata.revision.as_deref()),
        Some("0::D71B2AD6-F14B-4945-AD9A-7ABEC2907DB1")
    );
    Ok(())
}
