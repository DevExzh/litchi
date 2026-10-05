use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::detection::FileFormat;
use litchi_core::{Metadata, ReadAt};

use super::{EditorRefused, Pptx, PptxEditor, PptxSlide, SlideDescriptor};
use crate::error::{Code, Error, Source};
use crate::options::OpenOptions;
use crate::{opc, pptx};

/// Fixtures with metadata, without metadata, and with several slides.
const FIXTURES: [&str; 3] = ["backgrounds.pptx", "sample.pptx", "shapes.pptx"];

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/pptx")
        .join(name)
}

fn options() -> OpenOptions {
    OpenOptions::default()
}

// ---- constructor routes -------------------------------------------------

/// The path route: [`crate::open_with`].
#[cfg(any(unix, windows))]
fn open_path(path: &Path, options: &OpenOptions) -> Pptx {
    match crate::open_with(path, options).expect("path route") {
        crate::File::Pptx(pptx) => pptx,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

/// The byte route: [`crate::open_bytes`].
fn open_bytes(bytes: Vec<u8>) -> Pptx {
    match crate::open_bytes_with(bytes, &options()).expect("byte route") {
        crate::File::Pptx(pptx) => pptx,
        other => panic!("bytes opened as {other:?}"),
    }
}

fn open_source_package(bytes: Vec<u8>) -> Pptx {
    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::from_arc(Arc::new(bytes)));
    let package = opc::SourceBackedPackage::from_read_at(source).expect("indexed OPC package");
    Pptx::from_source_package(package, &options()).expect("source package route")
}

#[cfg(any(unix, windows))]
fn open_prepared(path: &Path) -> Pptx {
    let detection = crate::detect::detect_and_prepare(path).expect("detected");
    match crate::open_prepared_with(detection, &options()).expect("prepared route") {
        crate::File::Pptx(pptx) => pptx,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

fn open_eager(bytes: Vec<u8>) -> Pptx {
    match crate::detect::detect_format_smart(bytes).expect("detected") {
        crate::detect::DetectedFormat::Pptx(package) => {
            Pptx::from_opc_package(package, &options()).expect("eager route")
        },
        #[allow(
            unreachable_patterns,
            reason = "a pptx-only build detects no other format"
        )]
        other => panic!("expected the eager PPTX owner, got {:?}", other.format()),
    }
}

/// Every route that exists for this fixture, labelled.
fn every_route(name: &str) -> Vec<(&'static str, Pptx)> {
    let path = fixture(name);
    let bytes = std::fs::read(&path).expect("read fixture");
    #[allow(unused_mut, reason = "the path routes exist only on unix and windows")]
    let mut routes = vec![
        ("bytes", open_bytes(bytes.clone())),
        ("source package", open_source_package(bytes.clone())),
        ("eager", open_eager(bytes)),
    ];
    #[cfg(any(unix, windows))]
    {
        routes.push(("path", open_path(&path, &options())));
        routes.push(("prepared", open_prepared(&path)));
    }
    routes
}

// ---- leaf expectations (the computations the facade wraps) --------------
//
// `litchi::Presentation` cannot be called from here (no `litchi` dependency),
// so the expectations repeat the facade's own leaf calls. During S5b a
// throwaway facade test confirmed that `Presentation::open` and
// `Presentation::from_bytes` return exactly these values for `FIXTURES`.

struct Expected {
    text: String,
    slides: Vec<PptxSlide>,
    catalog: Vec<SlideDescriptor>,
    size: (i64, i64),
    metadata: Option<Metadata>,
    source_tolerance: pptx::ToleranceReport,
    eager_tolerance: pptx::ToleranceReport,
}

fn expected(name: &str) -> Expected {
    let bytes = std::fs::read(fixture(name)).expect("read fixture");

    let package = pptx::Package::from_bytes(&bytes).expect("eager leaf");
    let presentation = package.presentation().expect("eager presentation");
    let leaf_slides = presentation.slides().expect("eager slides");
    let mut texts = Vec::new();
    let mut slides = Vec::new();
    let mut eager_tolerance = package.tolerance_report().clone();
    for slide in &leaf_slides {
        let text = slide.text().expect("eager slide text");
        slide
            .text_with_tolerance(&mut eager_tolerance)
            .expect("eager slide repairs");
        slides.push(PptxSlide {
            text: text.clone(),
            name: slide.name().expect("eager slide name"),
        });
        if !text.is_empty() {
            texts.push(text);
        }
    }
    let catalog = presentation
        .slide_references()
        .expect("eager slide references")
        .into_iter()
        .enumerate()
        .map(|(position, reference)| SlideDescriptor::new(position, reference.id()))
        .collect();
    let size = presentation.slide_size().expect("eager slide size");
    let opc = opc::OpcPackage::from_bytes(&bytes).expect("eager OPC");
    let metadata = litchi_ooxml_common::properties::read_with_report(
        &opc,
        &mut litchi_ooxml_common::leniency::ToleranceReport::default(),
    )
    .expect("core properties")
    .map(Metadata::from)
    .filter(Metadata::has_data);

    let source =
        pptx::SourceBackedPresentation::from_reader(bytes.as_slice()).expect("source leaf");
    let mut source_tolerance = source.tolerance_report().clone();
    source
        .properties_with_tolerance(&mut source_tolerance)
        .expect("source properties");
    for slide in source.slides() {
        slide
            .text_with_tolerance(&mut source_tolerance)
            .expect("source slide repairs");
    }

    Expected {
        text: texts.join("\n\n"),
        slides,
        catalog,
        size,
        metadata,
        source_tolerance,
        eager_tolerance,
    }
}

#[test]
fn every_route_reads_like_the_leaf_the_facade_wraps() {
    for name in FIXTURES {
        let expected = expected(name);
        for (route, handle) in every_route(name) {
            let context = format!("{name} via {route}");
            assert_eq!(handle.is_source_backed(), route != "eager", "{context}");
            assert_eq!(handle.text().expect(&context), expected.text, "{context}");
            assert_eq!(
                handle.slide_count().expect(&context),
                expected.slides.len(),
                "{context}"
            );
            assert_eq!(
                handle.slides().expect(&context),
                expected.slides,
                "{context}"
            );
            for (position, slide) in expected.slides.iter().enumerate() {
                assert_eq!(
                    handle.slide(position).expect(&context).as_ref(),
                    Some(slide),
                    "{context}"
                );
            }
            assert_eq!(
                handle.slide(expected.slides.len()).expect(&context),
                None,
                "{context}"
            );
            assert_eq!(
                handle.slide_catalog().expect(&context),
                expected.catalog,
                "{context}"
            );
            assert_eq!(
                handle.slide_size().expect(&context),
                expected.size,
                "{context}"
            );
            // `Metadata` has no `PartialEq`; its `Debug` form names every field.
            let metadata = format!("{:?}", expected.metadata);
            assert_eq!(
                format!("{:?}", handle.metadata().expect(&context)),
                metadata,
                "{context}"
            );
            // The cached metadata answers the same.
            assert_eq!(
                format!("{:?}", handle.metadata().expect(&context)),
                metadata,
                "{context}"
            );
            let tolerance = if route == "eager" {
                &expected.eager_tolerance
            } else {
                &expected.source_tolerance
            };
            assert_eq!(
                &handle.tolerance_report().expect(&context),
                tolerance,
                "{context}"
            );
        }
    }
}

#[test]
fn fixture_values_match_the_package_content() {
    // The slide parts and the `docProps/core.xml` `dc:title`/`dc:creator`
    // values stored in these fixtures.
    let backgrounds = open_bytes(std::fs::read(fixture("backgrounds.pptx")).unwrap());
    assert_eq!(backgrounds.slide_count().unwrap(), 4);
    let metadata = backgrounds.metadata().unwrap().expect("core properties");
    assert_eq!(metadata.title.as_deref(), Some("Solid Fill"));
    assert_eq!(metadata.author.as_deref(), Some("yegor"));

    let sample = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    assert_eq!(sample.slide_count().unwrap(), 2);
    assert_eq!(sample.slide_catalog().unwrap().len(), 2);

    let shapes = open_eager(std::fs::read(fixture("shapes.pptx")).unwrap());
    assert_eq!(shapes.slide_count().unwrap(), 6);
}

#[test]
fn source_metadata_is_lazy_and_cached() {
    let handle = open_bytes(std::fs::read(fixture("backgrounds.pptx")).unwrap());
    let super::PptxOwner::Source(presentation) = &handle.0.owner else {
        panic!("byte route is source-backed");
    };
    assert!(handle.0.metadata.get().is_none(), "open reads no metadata");
    let opened = presentation.cache_diagnostics().cold_loads;
    assert!(handle.metadata().unwrap().is_some());
    let loaded = presentation.cache_diagnostics().cold_loads;
    assert!(loaded > opened, "the first call reads the core properties");
    assert!(handle.metadata().unwrap().is_some());
    assert_eq!(presentation.cache_diagnostics().cold_loads, loaded);
}

#[test]
fn debug_names_the_owner_state() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let handle = open_eager(bytes);
    assert!(!handle.is_source_backed());
    assert_eq!(
        format!("{handle:?}"),
        "Pptx { source_backed: false, retain_source: true, .. }"
    );
}

// ---- TOCTOU guard (finding F10) -----------------------------------------

fn assert_source_changed<T: std::fmt::Debug>(result: crate::Result<T>, operation: &str) {
    let error: Error = result.expect_err(operation);
    assert_eq!(error.code(), Code::IoError, "{operation}: {error}");
    assert_eq!(error.format(), Some(FileFormat::Pptx), "{operation}");
    assert!(
        matches!(
            error.leaf(),
            Source::Pptx(pptx::Error::Opc(opc::OpcError::SourceChanged { .. }))
        ),
        "{operation}: {:?}",
        error.leaf()
    );
}

#[cfg(any(unix, windows))]
#[test]
fn every_source_read_reports_a_typed_source_change() {
    let temporary = tempfile::Builder::new()
        .suffix(".pptx")
        .tempfile()
        .expect("temporary PPTX");
    std::fs::write(
        temporary.path(),
        std::fs::read(fixture("backgrounds.pptx")).unwrap(),
    )
    .expect("write fixture");
    let handle = open_path(temporary.path(), &options());
    // Fill the metadata cache: a cached answer must not hide the change.
    assert!(handle.metadata().unwrap().is_some());

    let mut file = std::fs::OpenOptions::new()
        .append(true)
        .open(temporary.path())
        .expect("reopen for mutation");
    file.write_all(b"source mutation").expect("mutate source");
    drop(file);

    assert_source_changed(handle.text(), "text");
    assert_source_changed(handle.slide_count(), "slide_count");
    assert_source_changed(handle.slides(), "slides");
    assert_source_changed(handle.slide(0), "slide");
    assert_source_changed(handle.slide_catalog(), "slide_catalog");
    assert_source_changed(handle.slide_size(), "slide_size");
    assert_source_changed(handle.metadata(), "metadata");
    assert_source_changed(handle.tolerance_report(), "tolerance_report");
}

#[test]
fn a_source_change_inside_the_properties_reader_stays_typed() {
    // `properties_with_tolerance` reports OPC failures of the core-properties
    // read through the shared OOXML error the leaf wraps.
    let changed = pptx::Error::CustomProperties(litchi_ooxml_common::Error::Opc(
        opc::OpcError::SourceChanged {
            expected: litchi_core::SourceVersion::new(7, 0),
            actual: litchi_core::SourceVersion::new(7, 1),
        },
    ));
    let error = Error::from(changed);
    assert_eq!(error.code(), Code::IoError);
    assert_eq!(error.format(), Some(FileFormat::Pptx));

    let cancelled =
        pptx::Error::CustomProperties(litchi_ooxml_common::Error::Opc(opc::OpcError::Cancelled));
    assert_eq!(Error::from(cancelled).code(), Code::Cancelled);
}

#[cfg(any(unix, windows))]
#[test]
fn cancellation_is_typed() {
    use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};

    let budget = Budget::root(
        "litchi-ops-pptx-cancellation-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0)
        .expect("valid execution limits");
    let presentation = pptx::SourceBackedPresentation::from_path_with_limits_and_execution_context(
        fixture("backgrounds.pptx"),
        pptx::ReadLimits::default(),
        ExecutionContext::new(budget, cancellation, limits),
    )
    .expect("managed source-backed PPTX");
    let handle = Pptx::from_source(presentation, &options());
    assert!(handle.metadata().unwrap().is_some());

    cancellation_source.cancel();
    assert_eq!(handle.slide_catalog().unwrap_err().code(), Code::Cancelled);
    assert_eq!(handle.metadata().unwrap_err().code(), Code::Cancelled);
}

// ---- editor construction (L-PPTX-1) -------------------------------------

fn expect_editor(result: Result<PptxEditor, EditorRefused>) -> PptxEditor {
    match result {
        Ok(editor) => editor,
        Err(refused) => panic!("no editor: {:?}", refused.error),
    }
}

fn expect_refused(result: Result<PptxEditor, EditorRefused>) -> EditorRefused {
    match result {
        Ok(_) => panic!("an editor was built"),
        Err(refused) => refused,
    }
}

#[test]
fn unique_source_handle_becomes_a_source_editor_without_reindexing() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let handle = open_bytes(bytes.clone());
    let PptxEditor::Source(editor) = expect_editor(handle.into_editor()) else {
        panic!("the source owner yields the source-backed editor");
    };
    assert_eq!(editor.slide_count(), 2);
    // The editor's package is the one indexed at open: its cache already
    // holds the presentation root, and an exact no-op republishes the input.
    let mut edit = editor.edit_slide(0).expect("edit slide");
    assert!(!edit.clear_transition().expect("no transition to clear"));
    let commit = edit.commit();
    let mut output = Vec::new();
    editor
        .publish_slide_commit_to_stream(&mut output, &commit)
        .expect("publish");
    assert_eq!(output, bytes);
}

#[test]
fn shared_source_handle_is_returned_readable() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let snapshot = handle.clone();
    let refused = expect_refused(handle.into_editor());
    assert!(refused.error.is_none(), "sharing is not an error");
    assert_eq!(refused.handle.slide_count().unwrap(), 2);
    assert_eq!(snapshot.slide_count().unwrap(), 2);
    drop(snapshot);
    assert!(matches!(
        expect_editor(refused.handle.into_editor()),
        PptxEditor::Source(_)
    ));
}

#[cfg(any(unix, windows))]
#[test]
fn changed_source_refuses_the_editor_with_a_typed_error() {
    let temporary = tempfile::Builder::new()
        .suffix(".pptx")
        .tempfile()
        .expect("temporary PPTX");
    std::fs::write(
        temporary.path(),
        std::fs::read(fixture("sample.pptx")).unwrap(),
    )
    .expect("write fixture");
    let handle = open_path(temporary.path(), &options());
    let mut file = std::fs::OpenOptions::new()
        .append(true)
        .open(temporary.path())
        .expect("reopen for mutation");
    file.write_all(b"source mutation").expect("mutate source");
    drop(file);

    let refused = expect_refused(handle.into_editor());
    assert_source_changed::<()>(Err(refused.error.expect("typed error")), "into_editor");
    assert!(refused.handle.is_source_backed());
}

#[test]
fn eager_handle_becomes_an_opened_transaction() {
    let handle = open_eager(std::fs::read(fixture("sample.pptx")).unwrap());
    let snapshot = handle.clone();
    let PptxEditor::Eager(transaction) = expect_editor(handle.into_editor()) else {
        panic!("the eager owner yields an opened transaction");
    };
    drop(transaction);
    // The eager editor is detached, so the snapshot stays readable.
    assert_eq!(snapshot.slide_count().unwrap(), 2);
}

#[test]
fn editor_needs_the_retained_source() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::from_arc(Arc::new(bytes)));
    let presentation = pptx::SourceBackedPresentation::from_read_at(source).unwrap();
    let handle = Pptx::from_source(presentation, &OpenOptions::new().retain_source(false));
    let refused = expect_refused(handle.into_editor());
    let error = refused.error.expect("typed refusal");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(error.suggestion().is_some());
    assert_eq!(refused.handle.slide_count().unwrap(), 2);
}

// ---- handle shape ---------------------------------------------------------

#[test]
fn handle_is_one_pointer_and_shares_state_on_clone() {
    assert_eq!(size_of::<Pptx>(), size_of::<usize>());
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let clone = handle.clone();
    assert!(Arc::ptr_eq(&handle.0, &clone.0));
}
