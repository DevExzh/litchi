use std::path::PathBuf;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_keynote::{
    OpenOptions, Package, Presentation, SemanticLimits, SlideSelector, SourceLimits, Stats,
    presentation::Error as PresentationError, presentation::LimitKind as PresentationLimitKind,
    presentation::Stats as PresentationStats, show::Mode,
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
fn archive_free_document_accepts_borrowed_package_bytes() -> Result<(), Box<dyn std::error::Error>>
{
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;
    let document = Presentation::from_bytes(&source)?;

    assert_eq!(document.show(), package.show()?);
    assert_eq!(document.text()?, EXPECTED_TEXT);
    assert_eq!(
        document.stats().map(|stats| stats.slide_count),
        Some(package.stats()?.slide_count)
    );
    assert_eq!(
        document
            .metadata()
            .and_then(|metadata| metadata.application.as_deref()),
        Some("Keynote")
    );

    // Semantic projection owns its values instead of borrowing package bytes.
    drop(source);
    assert_eq!(document.text()?, EXPECTED_TEXT);
    Ok(())
}

#[test]
fn archive_free_byte_ingress_honors_source_and_semantic_limits()
-> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let defaults = SourceLimits::default();
    let exact_source = SourceLimits::builder()
        .max_input_bytes(u64::try_from(source.len())?)?
        .max_entries(defaults.max_entries())?
        .max_entry_bytes(defaults.max_entry_bytes())?
        .max_aggregate_bytes(defaults.max_aggregate_bytes())?
        .max_component_bytes(defaults.max_component_bytes())?
        .build()?;
    let source_error = Presentation::from_bytes_with_options(
        &source,
        OpenOptions::new()
            .source_limits(
                SourceLimits::builder()
                    .max_input_bytes(u64::try_from(source.len().saturating_sub(1))?)?
                    .max_entries(defaults.max_entries())?
                    .max_entry_bytes(defaults.max_entry_bytes())?
                    .max_aggregate_bytes(defaults.max_aggregate_bytes())?
                    .max_component_bytes(defaults.max_component_bytes())?
                    .build()?,
            )
            .semantic_limits(SemanticLimits::default()),
    )
    .expect_err("source max-minus-one must reject complete package bytes");
    assert!(matches!(source_error, PresentationError::Limit { .. }));

    let semantic = SemanticLimits::builder()
        .max_objects(EXPECTED_OBJECTS - 1)?
        .build()?;
    let semantic_error = Presentation::from_bytes_with_options(
        &source,
        OpenOptions::new()
            .source_limits(exact_source)
            .semantic_limits(semantic),
    )
    .expect_err("object max-minus-one must reject before publishing a document");
    assert!(matches!(
        semantic_error,
        PresentationError::Limit {
            kind: PresentationLimitKind::Objects,
            maximum,
            ..
        } if maximum == (EXPECTED_OBJECTS - 1) as u64
    ));
    Ok(())
}

#[test]
fn focused_reader_surface_is_exact_shareable_and_archive_free()
-> Result<(), Box<dyn std::error::Error>> {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Package>();
    assert_send_sync::<Presentation>();
    assert_send_sync::<PresentationStats>();
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
    assert_eq!(snapshot.options(), package.options());
    assert_eq!(snapshot.stats()?, package.stats()?);
    assert_eq!(snapshot.text()?, EXPECTED_TEXT);
    assert_eq!(snapshot.show()?, package.show()?);
    assert_eq!(snapshot.slides()?.as_ptr(), package.slides()?.as_ptr());

    let semantic = package.presentation()?;
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

/// Every file below `root`, by `/`-separated relative name.
fn bundle_files(
    root: &std::path::Path,
) -> Result<std::collections::BTreeMap<String, Vec<u8>>, Box<dyn std::error::Error>> {
    fn walk(
        directory: &std::path::Path,
        prefix: &str,
        files: &mut std::collections::BTreeMap<String, Vec<u8>>,
    ) -> Result<(), Box<dyn std::error::Error>> {
        for entry in std::fs::read_dir(directory)? {
            let entry = entry?;
            let name = entry.file_name().to_string_lossy().into_owned();
            let name = if prefix.is_empty() {
                name
            } else {
                format!("{prefix}/{name}")
            };
            if entry.file_type()?.is_dir() {
                walk(&entry.path(), &name, files)?;
            } else {
                files.insert(name, std::fs::read(entry.path())?);
            }
        }
        Ok(())
    }
    let mut files = std::collections::BTreeMap::new();
    walk(root, "", &mut files)?;
    Ok(files)
}

#[test]
fn package_directories_convert_and_save_back_byte_exactly() -> Result<(), Box<dyn std::error::Error>>
{
    // A directory bundle converts to an exact-source single-file package.
    let package = Package::open_with_options(directory_fixture_path(), OpenOptions::default())?;
    package.validate()?;
    assert_eq!(package.text()?, EXPECTED_TEXT);
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    let reopened = Package::from_bytes(&bytes)?;
    assert_eq!(reopened.text()?, EXPECTED_TEXT);
    // Deterministic: a second capture converts to the same bytes.
    let mut again = Vec::new();
    Package::open(directory_fixture_path())?.write_to(&mut again)?;
    assert_eq!(again, bytes);

    // Saved back as a directory, every bundle file is the source's bytes.
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("bundle.key");
    package.save_directory(&path)?;
    assert_eq!(
        bundle_files(&path)?,
        bundle_files(&directory_fixture_path())?
    );
    let mut round_trip = Vec::new();
    Package::open(&path)?.write_to(&mut round_trip)?;
    assert_eq!(round_trip, bytes);

    // Saving over the existing directory replaces it, leaving no staging
    // or retired directory behind; a file destination is refused.
    let single = Package::open(fixture_path())?;
    single.save_directory(&path)?;
    assert_eq!(Package::open(&path)?.text()?, EXPECTED_TEXT);
    assert_eq!(std::fs::read_dir(temp.path())?.count(), 1);
    let file = temp.path().join("file.key");
    single.save(&file)?;
    assert!(matches!(
        single.save_directory(&file),
        Err(litchi_keynote::SaveError::Publication(_))
    ));

    // The capture is bounded by the selected source limits.
    let limits = SourceLimits::builder().max_input_bytes(1024)?.build()?;
    assert!(
        Package::open_with_options(
            directory_fixture_path(),
            OpenOptions::default().source_limits(limits)
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn metadata_ignores_earlier_malformed_near_names_and_preserves_source()
-> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let catalog = Catalog::from_bytes(&source, Default::default())?;
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
    let document = Presentation::open(path)?;
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
            Presentation::open(path),
            Err(PresentationError::UnexpectedFormat { .. })
        ));
    }
}

#[test]
fn archive_free_document_has_full_zip_and_directory_semantic_parity()
-> Result<(), Box<dyn std::error::Error>> {
    let package = Package::open(fixture_path())?;
    let zipped = Presentation::open(fixture_path())?;
    let directory = Presentation::open(directory_fixture_path())?;

    zipped.validate()?;
    directory.validate()?;
    assert_eq!(zipped.show(), package.show()?);
    assert_eq!(directory.show(), zipped.show());
    assert_eq!(zipped.text()?, EXPECTED_TEXT);
    assert_eq!(directory.text()?, EXPECTED_TEXT);
    assert_eq!(
        zipped.stats().map(|stats| stats.slide_count),
        Some(package.stats()?.slide_count)
    );
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
    let defaults = SourceLimits::default();
    let exact_source = SourceLimits::builder()
        .max_input_bytes(exact_input)?
        .max_entries(defaults.max_entries())?
        .max_entry_bytes(defaults.max_entry_bytes())?
        .max_aggregate_bytes(defaults.max_aggregate_bytes())?
        .max_component_bytes(defaults.max_component_bytes())?
        .build()?;
    let exact = Presentation::open_with_options(
        &source,
        OpenOptions::new()
            .source_limits(exact_source)
            .semantic_limits(SemanticLimits::default()),
    )?;
    assert_eq!(exact.text()?, EXPECTED_TEXT);
    assert_eq!(exact.stats().map(|stats| stats.slide_count), Some(1));

    let too_small_source = SourceLimits::builder()
        .max_input_bytes(exact_input - 1)?
        .max_entries(defaults.max_entries())?
        .max_entry_bytes(defaults.max_entry_bytes())?
        .max_aggregate_bytes(defaults.max_aggregate_bytes())?
        .max_component_bytes(defaults.max_component_bytes())?
        .build()?;
    let error = Presentation::open_with_options(
        &source,
        OpenOptions::new()
            .source_limits(too_small_source)
            .semantic_limits(SemanticLimits::default()),
    )
    .expect_err("aggregate index plus properties max-minus-one must refuse");
    assert!(matches!(error, PresentationError::Limit { .. }));

    let semantic = SemanticLimits::builder()
        .max_objects(EXPECTED_OBJECTS - 1)?
        .build()?;
    let error = Presentation::open_with_options(
        &source,
        OpenOptions::new()
            .source_limits(exact_source)
            .semantic_limits(semantic),
    )
    .expect_err("object max-minus-one must refuse before publishing a document");
    assert!(matches!(
        error,
        PresentationError::Limit {
            kind: PresentationLimitKind::Objects,
            observed,
            maximum,
            ..
        } if observed == EXPECTED_OBJECTS as u64 && maximum == (EXPECTED_OBJECTS - 1) as u64
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
