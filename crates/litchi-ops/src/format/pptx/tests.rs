use std::io::Write as _;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use litchi_core::detection::FileFormat;
use litchi_core::{Metadata, Position, ReadAt};

use super::{Editable, Pptx, Slide, SlideDescriptor};
use crate::error::{Code, Error, Source};
use crate::options::OpenOptions;
use crate::{IntoTarget, opc, pptx};

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
    let package = opc::SourceBackedPackage::from_read_at(
        source,
        litchi_opc::SourceLimits::default(),
        litchi_opc::SourceCacheLimits::default(),
        litchi_opc::SourceReadPolicy::exact(),
        None,
    )
    .expect("indexed OPC package");
    Pptx::from_source_package(package, &options()).expect("source package route")
}

#[cfg(any(unix, windows))]
fn open_prepared(path: &Path) -> Pptx {
    let detection =
        crate::detect::detect_and_prepare(path, &crate::detect::SourceLimits::default())
            .expect("detected");
    match crate::open_prepared_with(detection, &options()).expect("prepared route") {
        crate::File::Pptx(pptx) => pptx,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

fn open_eager(bytes: Vec<u8>) -> Pptx {
    match crate::detect::detect_format_smart(bytes, &crate::detect::SourceLimits::default())
        .expect("detected")
    {
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
    slides: Vec<Slide>,
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
        slides.push(Slide {
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
    let opc = opc::OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
        .expect("eager OPC");
    let metadata = litchi_ooxml_common::properties::read_with_report(
        &opc,
        &mut litchi_ooxml_common::leniency::ToleranceReport::default(),
    )
    .expect("core properties")
    .unwrap_or_default();
    // The application properties join the core ones (record 0253).
    let metadata =
        Some(crate::format::metadata::with_app_props(metadata, &opc)).filter(Metadata::has_data);

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
    let presentation = pptx::SourceBackedPresentation::open_with_options(
        fixture("backgrounds.pptx"),
        litchi_pptx::OpenOptions::default()
            .source_limits(pptx::SourceLimits::default())
            .execution_context(ExecutionContext::new(budget, cancellation, limits)),
    )
    .expect("managed source-backed PPTX");
    let handle = Pptx::from_source(presentation, &options());
    assert!(handle.metadata().unwrap().is_some());

    let _promoted = source_owner(&handle);
    cancellation_source.cancel();
    assert_eq!(handle.editable().err().unwrap().code(), Code::Cancelled);
    assert_eq!(handle.slide_catalog().unwrap_err().code(), Code::Cancelled);
    assert_eq!(handle.metadata().unwrap_err().code(), Code::Cancelled);
}

// ---- editor construction (L-PPTX-1) -------------------------------------

fn source_owner(handle: &Pptx) -> pptx::source::Package {
    match handle.editable().expect("editable native owner") {
        Editable::Source(package) => package,
        Editable::Eager(_) => panic!("expected deferred source"),
    }
}

/// Build a source-backed deck whose every ordered slide owns distinct related
/// parts.  The comments are intentionally added through the native package
/// graph so this fixture exercises the same relationship resolution as a
/// producer-authored presentation.
fn related_order_fixture() -> Vec<u8> {
    let slides = [
        ("Slide A", "Notes A", "Comment A", "Modern A"),
        ("Slide B", "Notes B", "Comment B", "Modern B"),
        ("Slide C", "Notes C", "Comment C", "Modern C"),
    ];
    let mut builder = pptx::builder::Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        for (name, notes, _, _) in slides {
            let slide = presentation.add_slide().unwrap();
            slide.set_title(name);
            slide.set_notes(notes);
            slide.add_text_box(name, 0, 0, 914_400, 914_400);
        }
    }
    let package = builder.build().unwrap();
    let mut opc_edit = package.edit_opc().unwrap();
    opc_edit
        .update(|opc| {
            // The builder's title is shape text; authored slide facts are the
            // independent p:cSld@name values exposed by the unified node.
            for (index, (name, _, _, _)) in slides.iter().enumerate() {
                let slide_name =
                    opc::PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).unwrap();
                let part = opc.get_part_mut(&slide_name)?;
                let xml = String::from_utf8(part.blob().to_vec()).unwrap();
                let root = xml.find("<p:cSld ").unwrap();
                let open_end = root + xml[root..].find('>').unwrap();
                let marker = " name=\"";
                let name_start = root + xml[root..open_end].find(marker).unwrap() + marker.len();
                let name_end = name_start + xml[name_start..open_end].find('"').unwrap();
                let mut named = xml.clone();
                named.replace_range(name_start..name_end, name);
                assert_ne!(named, xml, "builder slide has no cSld element");
                part.set_blob(named.into_bytes());
            }

            Ok(())
        })
        .unwrap();
    let package = opc_edit.commit().unwrap().into_package();
    let mut edit = package.edit_comments().unwrap();
    let conformance = pptx::comments::Conformance::Transitional;
    edit.add_author(pptx::comments::Author::new(1, "Ada", "AD"), conformance)
        .unwrap();
    for (index, (_, _, text, _)) in slides.iter().enumerate() {
        edit.add(
            index,
            pptx::comments::Comment::new(1, *text, 100, 200).with_index(index as u32 + 1),
            conformance,
        )
        .unwrap();
    }

    let package = edit.commit().unwrap().into_package();
    let mut edit = package.edit_modern_comments().unwrap();
    const MODERN_AUTHOR: &str = "{11111111-1111-4111-8111-111111111111}";
    edit.add_author(pptx::modern_comments::Author {
        id: MODERN_AUTHOR.into(),
        name: "Grace".into(),
        initials: Some("GR".into()),
        user_id: "grace@example.test".into(),
        provider_id: "litchi-test".into(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    })
    .unwrap();
    for (index, (_, _, _, title)) in slides.iter().enumerate() {
        edit.add(
            index,
            pptx::modern_comments::Comment {
                id: format!("{{22222222-2222-4222-8222-22222222220{}}}", index + 1),
                author_id: MODERN_AUTHOR.into(),
                status: None,
                created: format!("2026-10-10T10:0{}:00Z", index),
                start_date: None,
                due_date: None,
                assigned_to: None,
                complete: None,
                title: Some((*title).into()),
                namespace_declarations: Vec::new(),
                anchors: Vec::new(),
                position: None,
                reply_list_namespace_declarations: Vec::new(),
                replies: Vec::new(),
                reply_list_present: false,
                text_body_xml: None,
                extension_xml: None,
            },
        )
        .unwrap();
    }
    let package = edit.commit().unwrap().into_package();
    package.to_bytes().unwrap()
}

fn node_property(node: &crate::node::Node<'_>, key: &str) -> Option<String> {
    node.view()
        .format()
        .unwrap()
        .find(|property| property.key() == key)
        .map(|property| property.to_string())
}

fn node_text(node: &crate::node::Node<'_>) -> Option<String> {
    node.view().text().unwrap().map(|text| text.into_owned())
}

#[test]
fn source_promotion_reuses_admitted_catalog_and_preserves_exact_noop() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let handle = open_bytes(bytes.clone());
    let cold = handle.cold_loads();
    let package = source_owner(&handle);
    assert_eq!(handle.cold_loads(), cold);
    assert_eq!(package.presentation().unwrap().slide_count().unwrap(), 2);
    assert_eq!(
        package
            .edit()
            .unwrap()
            .commit()
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        bytes
    );
    assert_eq!(handle.slide_count().unwrap(), 2);
}

#[test]
fn native_pptx_commits_account_their_retained_proof() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let target = "/slide[1]/shape[1]".into_target().unwrap();
    let limits = crate::transaction::SemanticLimits::default();

    let text = handle
        .set_with_limits(
            target.clone(),
            crate::props! { "text" => "native snapshot accounting" },
            &limits,
        )
        .unwrap();
    assert!(text.snapshot_bytes() > 0);

    let link = handle
        .set_with_limits(
            target.clone(),
            crate::props! { "link" => "https://example.com/native-snapshot-charge" },
            &limits,
        )
        .unwrap();
    assert!(link.snapshot_bytes() > 0);

    // Click setters compose a text/property edit with a second hyperlink
    // edit.  Both native proofs remain charged on the unsigned path.
    let combined = handle
        .set_with_limits(
            target,
            crate::props! {
                "text" => "native snapshot accounting",
                "link" => "https://example.com/native-snapshot-charge",
            },
            &limits,
        )
        .unwrap();
    assert!(combined.snapshot_bytes() > text.snapshot_bytes());
    assert!(combined.snapshot_bytes() > link.snapshot_bytes());
}

#[test]
fn combined_click_edits_respect_the_remaining_snapshot_budget() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let target = "/slide[1]/shape[1]".into_target().unwrap();
    let text = handle
        .set_with_limits(
            target.clone(),
            crate::props! { "text" => "combined-link budget" },
            &crate::transaction::SemanticLimits::default(),
        )
        .unwrap();
    let limits = crate::transaction::SemanticLimits::builder()
        .max_snapshot_bytes(text.snapshot_bytes())
        .unwrap()
        .build()
        .unwrap();
    let error = match handle.set_with_limits(
        target,
        crate::props! {
            "text" => "combined-link budget",
            "link" => "https://example.com/combined-link-budget",
        },
        &limits,
    ) {
        Ok(_) => panic!("combined click edit unexpectedly succeeded"),
        Err(error) => error,
    };
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn metadata_noop_accounts_its_complete_retained_scope() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let target = "/metadata".into_target().unwrap();
    let default_limits = crate::transaction::SemanticLimits::default();
    let noop = handle
        .set_with_limits(target.clone(), crate::props! {}, &default_limits)
        .unwrap();
    let retained = noop.snapshot_bytes();
    assert!(retained > 1, "metadata scope was not retained: {retained}");

    let exact_limits = crate::transaction::SemanticLimits::builder()
        .max_snapshot_bytes(retained)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        handle
            .set_with_limits(target.clone(), crate::props! {}, &exact_limits)
            .is_ok()
    );

    let below_limits = crate::transaction::SemanticLimits::builder()
        .max_snapshot_bytes(retained - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = match handle.set_with_limits(target, crate::props! {}, &below_limits) {
        Ok(_) => panic!("metadata no-op unexpectedly exceeded its lower bound"),
        Err(error) => error,
    };
    assert_eq!(error.code(), Code::LimitExceeded);
}

#[test]
fn coalescer_uses_native_patch_budget_not_archive_size() {
    let mut large_text = String::with_capacity(256 * 1024);
    let mut state = 0x1234_5678_u32;
    for _ in 0..256 * 1024 {
        state = state.wrapping_mul(1_664_525).wrapping_add(1_013_904_223);
        large_text.push((b'a' + ((state >> 24) % 26) as u8) as char);
    }
    let mut builder = pptx::builder::Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        presentation
            .add_slide()
            .unwrap()
            .add_text_box("small", 0, 0, 914_400, 914_400);
        presentation
            .add_slide()
            .unwrap()
            .add_text_box(&large_text, 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};
    let budget = Budget::root(
        "litchi-ops-pptx-native-budget-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, cancellation) = CancellationSource::pair();
    let execution_limits =
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap();
    let source: Arc<dyn ReadAt> =
        Arc::new(litchi_core::OwnedSource::from_arc(Arc::new(bytes.clone())));
    let presentation = pptx::SourceBackedPresentation::from_read_at_with_options(
        source,
        litchi_pptx::OpenOptions::default()
            .source_limits(pptx::SourceLimits::default())
            .execution_context(ExecutionContext::new(
                budget,
                cancellation,
                execution_limits,
            )),
    )
    .unwrap();
    let handle = Pptx::from_source(presentation, &options());
    assert!(handle.is_source_backed());

    // The first-slide edit retains only the changed native patch.  The
    // source archive also contains the unrelated large second slide, so its
    // size must not be used as a proxy for this exact patch budget.
    let target = "/slide[1]/shape[1]".into_target().unwrap();
    let probe = handle
        .set_with_limits(
            target.clone(),
            crate::props! { "text" => "small patch" },
            &crate::transaction::SemanticLimits::default(),
        )
        .unwrap();
    let remaining = probe.snapshot_bytes();
    assert!(remaining > 0);
    assert!(bytes.len() > remaining);

    let limits = crate::transaction::SemanticLimits::builder()
        .max_snapshot_bytes(remaining)
        .unwrap()
        .build()
        .unwrap();
    let mut run =
        super::write::coalesce::Run::new(&handle, &limits, remaining).expect("native edit");
    let props = crate::Props::from_props(crate::props! { "text" => "small patch" }).unwrap();
    let plan = run
        .shape_plan(&target, &props, &limits)
        .unwrap()
        .expect("direct shape plan");
    assert!(run.stage(plan, 0).unwrap());
    run.commit().unwrap();
}

#[test]
fn shared_source_promotion_retains_one_lineage_and_live_readers() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let snapshot = handle.clone();
    let original = snapshot.text().unwrap();
    let package = source_owner(&handle);
    let repeated = source_owner(&snapshot);
    let slide = package.presentation().unwrap().slide(0).unwrap();
    let scene = slide.scene().unwrap();
    let shape = scene
        .iter()
        .position(|shape| shape.view().common().text().is_some())
        .unwrap();
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(0, shape, "promoted edit").unwrap();
    let commit = edit.commit().unwrap();
    let replayed = repeated.apply(commit.patch()).unwrap().into_package();
    assert_eq!(
        replayed.to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    assert_eq!(snapshot.text().unwrap(), original);
    assert_eq!(handle.text().unwrap(), original);
    assert_eq!(slide.scene().unwrap().len(), scene.len());
}

#[test]
fn concurrent_source_promotion_shares_native_patch_lineage() {
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let packages = std::thread::scope(|scope| {
        let tasks: Vec<_> = (0..8)
            .map(|_| scope.spawn(|| source_owner(&handle)))
            .collect();
        tasks
            .into_iter()
            .map(|task| task.join().unwrap())
            .collect::<Vec<_>>()
    });
    let patch = packages[0]
        .edit()
        .unwrap()
        .commit()
        .unwrap()
        .patch()
        .clone();
    for package in &packages {
        assert!(package.apply(&patch).is_ok());
    }
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
    let _promoted = source_owner(&handle);
    let mut file = std::fs::OpenOptions::new()
        .append(true)
        .open(temporary.path())
        .expect("reopen for mutation");
    file.write_all(b"source mutation").expect("mutate source");
    drop(file);

    assert_source_changed(handle.editable().map(|_| ()), "editable");
    assert!(handle.is_source_backed());
}

#[test]
fn eager_handle_becomes_an_opened_transaction() {
    let handle = open_eager(std::fs::read(fixture("sample.pptx")).unwrap());
    let snapshot = handle.clone();
    let Editable::Eager(package) = handle.editable().unwrap() else {
        panic!("the eager owner yields an opened transaction");
    };
    let Editable::Eager(repeated) = snapshot.editable().unwrap() else {
        panic!("eager");
    };
    assert!(Arc::ptr_eq(&package, &repeated));
    drop(package.edit().unwrap());
    // The eager editor is detached, so the snapshot stays readable.
    assert_eq!(snapshot.slide_count().unwrap(), 2);
}

#[test]
fn editor_needs_the_retained_source() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let source: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::from_arc(Arc::new(bytes)));
    let presentation = pptx::SourceBackedPresentation::from_read_at(source).unwrap();
    let handle = Pptx::from_source(presentation, &OpenOptions::new().retain_source(false));
    let error = handle.editable().err().expect("typed refusal");
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(error.suggestion().is_some());
    assert_eq!(handle.slide_count().unwrap(), 2);
}

// ---- handle shape ---------------------------------------------------------

#[test]
fn handle_is_one_pointer_and_shares_state_on_clone() {
    assert_eq!(size_of::<Pptx>(), size_of::<usize>());
    let handle = open_bytes(std::fs::read(fixture("sample.pptx")).unwrap());
    let clone = handle.clone();
    assert!(Arc::ptr_eq(&handle.0, &clone.0));
}

#[test]
fn query_items_stay_as_small_as_a_text_item() {
    // Record 0326: every node a query yields is moved several times, so its
    // size is paid per match. Record 0321's column, hyperlink and equation
    // items are boxed, and a chart or SmartArt frame keeps one lazy part
    // cell: 144 bytes, against 176 before.
    use super::node::{Item, TextItem, TreeItem};
    assert!(size_of::<Item>() <= size_of::<TextItem>() + size_of::<usize>());
    assert!(size_of::<TreeItem>() <= 88, "{}", size_of::<TreeItem>());
    assert!(
        size_of::<super::Node<'static>>() <= 144,
        "{}",
        size_of::<super::Node<'static>>()
    );
}

#[test]
fn transient_promotion_fence_failure_can_retry_without_reopening() {
    use litchi_core::{OwnedSource, SourceVersion};
    use std::sync::atomic::{AtomicUsize, Ordering};
    struct FlakySource {
        source: OwnedSource,
        fail_version: AtomicUsize,
    }
    impl ReadAt for FlakySource {
        fn len(&self) -> std::io::Result<u64> {
            self.source.len()
        }
        fn version(&self) -> std::io::Result<SourceVersion> {
            if self
                .fail_version
                .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |n| n.checked_sub(1))
                == Ok(1)
            {
                return Err(std::io::Error::other("injected transient version failure"));
            }
            self.source.version()
        }
        fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
            self.source.read_at(offset, output)
        }
    }
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let source = Arc::new(FlakySource {
        source: OwnedSource::new(bytes.clone()),
        fail_version: AtomicUsize::new(0),
    });
    let projection = pptx::SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let handle = Pptx::from_source(projection, &options());
    source.fail_version.store(2, Ordering::SeqCst);
    assert_eq!(handle.editable().err().unwrap().code(), Code::IoError);
    assert_eq!(source_owner(&handle).to_bytes().unwrap(), bytes);
    source.fail_version.store(2, Ordering::SeqCst);
    assert_eq!(handle.editable().err().unwrap().code(), Code::IoError);
    assert_eq!(source_owner(&handle).to_bytes().unwrap(), bytes);
}

#[test]
fn adopted_source_revision_drives_text_nodes_and_subsequent_edits() {
    let bytes = std::fs::read(fixture("sample.pptx")).unwrap();
    let before = open_bytes(bytes.clone());
    let original_text = before.text().unwrap();
    let package = source_owner(&before);
    let mut edit = package.edit().unwrap();
    edit.set_shape_text(0, 0_usize, "revised title").unwrap();
    let committed = edit.commit().unwrap().into_package();
    let reopened = open_bytes(committed.to_bytes().unwrap());
    let after = before.adopt_source(committed).unwrap();
    assert!(after.is_source_backed());
    assert_eq!(after.text().unwrap(), reopened.text().unwrap());
    assert!(after.text().unwrap().contains("revised title"));
    assert_eq!(after.slides().unwrap(), reopened.slides().unwrap());
    assert_eq!(
        after.tolerance_report().unwrap(),
        reopened.tolerance_report().unwrap()
    );
    assert_eq!(after.slide(0).unwrap(), reopened.slide(0).unwrap());
    assert!(after.slide(usize::MAX).unwrap().is_none());
    assert_eq!(
        after.slide_catalog().unwrap(),
        before.slide_catalog().unwrap()
    );
    assert_eq!(
        serde_json::to_value(after.metadata().unwrap()).unwrap(),
        serde_json::to_value(before.metadata().unwrap()).unwrap()
    );
    assert_eq!(after.slide_size().unwrap(), before.slide_size().unwrap());
    let file = crate::File::Pptx(after.clone());
    let reopened_file = crate::File::Pptx(reopened);
    for path in [
        "/slide[1]",
        "/slide[1]/shape[1]",
        "/slide[1]/shape[1]/paragraph[1]/run[1]",
    ] {
        assert_eq!(
            file.get(path).unwrap().view().text().unwrap(),
            reopened_file.get(path).unwrap().view().text().unwrap()
        );
    }
    let current = source_owner(&after);
    let mut next = current.edit().unwrap();
    next.set_shape_text(1, 0_usize, "second revision").unwrap();
    let final_handle = after
        .adopt_source(next.commit().unwrap().into_package())
        .unwrap();
    let text = final_handle.text().unwrap();
    assert!(text.contains("revised title"));
    assert!(text.contains("second revision"));
    assert_eq!(before.text().unwrap(), original_text);
    assert_eq!(source_owner(&before).to_bytes().unwrap(), bytes);
}

#[test]
fn unsigned_managed_source_respects_native_materialization_policy() {
    use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits};
    use std::num::{NonZeroU64, NonZeroUsize};
    let budget = Budget::root(
        "pptx-unsigned-policy",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, cancellation) = CancellationSource::pair();
    let limits =
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap();
    let presentation = pptx::SourceBackedPresentation::open_with_options(
        fixture("sample.pptx"),
        litchi_pptx::OpenOptions::default()
            .source_limits(pptx::SourceLimits::default())
            .execution_context(ExecutionContext::new(budget, cancellation, limits)),
    )
    .unwrap();
    let handle = Pptx::from_source(presentation, &options())
        .with_signature(super::Signature::RemovalAuthorized);
    let cold = handle.cold_loads();
    let error = handle.unsigned_package().unwrap_err();
    assert!(matches!(
        error.leaf(),
        Source::Opc(opc::OpcError::ManagedPackageMaterialization)
    ));
    assert_eq!(handle.cold_loads(), cold);
}

#[test]
fn source_order_rebinds_slide_facts_and_all_related_parts() {
    let bytes = related_order_fixture();
    let before = open_source_package(bytes);
    let before_ids: Vec<_> = before
        .slide_catalog()
        .unwrap()
        .into_iter()
        .map(|descriptor| descriptor.slide_id())
        .collect();
    assert_eq!(before_ids, [256, 257, 258]);

    let mut edit = source_owner(&before).edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    let after = before
        .adopt_source(edit.commit().unwrap().into_package())
        .unwrap();
    let after_ids: Vec<_> = after
        .slide_catalog()
        .unwrap()
        .into_iter()
        .map(|descriptor| descriptor.slide_id())
        .collect();
    assert_eq!(after_ids, [257, 258, 256]);

    let file = crate::File::Pptx(after.clone());
    for (position, (name, notes, comment, modern_title)) in [
        ("Slide B", "Notes B", "Comment B", "Modern B"),
        ("Slide C", "Notes C", "Comment C", "Modern C"),
        ("Slide A", "Notes A", "Comment A", "Modern A"),
    ]
    .into_iter()
    .enumerate()
    {
        let slide_number = position + 1;
        let slide_path = format!("/slide[{slide_number}]");
        let slide = file.get(slide_path.as_str()).unwrap();
        assert_eq!(node_property(&slide, "name"), Some(name.into()));

        let notes_path = format!("/slide[{slide_number}]/notes");
        let notes_node = file.get(notes_path.as_str()).unwrap();
        assert_eq!(node_text(&notes_node), Some(notes.into()));

        let comment_path = format!("/slide[{slide_number}]/comment[1]");
        let comment_node = file.get(comment_path.as_str()).unwrap();
        assert_eq!(node_text(&comment_node), Some(comment.into()));

        let modern_path = format!("/slide[{slide_number}]/modernComment[1]");
        let modern_node = file.get(modern_path.as_str()).unwrap();
        assert_eq!(
            node_property(&modern_node, "title"),
            Some(modern_title.into())
        );
    }

    assert!(after.slide(3).unwrap().is_none());
    assert!(after.slide(usize::MAX).unwrap().is_none());
}

#[test]
fn stock_builder_source_order_preserves_presentation_properties() {
    let bytes = related_order_fixture();
    let before_opc =
        opc::OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let pres_props_uri = opc::PackURI::new("/ppt/presProps.xml").unwrap();
    let before_pres_props = before_opc
        .get_part(&pres_props_uri)
        .unwrap()
        .blob()
        .to_vec();

    let source = open_source_package(bytes);
    let moved = crate::File::Pptx(source)
        .move_to("/slide[1]", crate::Placement::End)
        .unwrap();
    let after_opc = opc::OpcPackage::from_bytes(
        &moved.to_bytes().unwrap(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    let after_pres_props = after_opc.get_part(&pres_props_uri).unwrap().blob().to_vec();
    assert_eq!(after_pres_props, before_pres_props);
}

fn unknown_pres_props_extension_fixture() -> Vec<u8> {
    let bytes = related_order_fixture();
    let mut package =
        opc::OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let pres_props_uri = opc::PackURI::new("/ppt/presProps.xml").unwrap();
    let part = package.get_part_mut(&pres_props_uri).unwrap();
    let xml = String::from_utf8(part.blob().to_vec()).unwrap();
    let marker = "</p:extLst>";
    let marker_start = xml.find(marker).unwrap();
    let extension =
        r#"<p:ext uri="urn:litchi:test:unknown"><x:payload xmlns:x="urn:litchi:test"/></p:ext>"#;
    let mut updated = xml;
    updated.insert_str(marker_start, extension);
    part.set_blob(updated.into_bytes());
    opc::PackageWriter::to_bytes(&package).unwrap()
}

#[test]
fn unknown_presentation_properties_extension_remains_refused() {
    let source = open_source_package(unknown_pres_props_extension_fixture());
    let error = match crate::File::Pptx(source).move_to("/slide[1]", crate::Placement::End) {
        Ok(_) => panic!("unknown presentation-properties extension was silently rewritten"),
        Err(error) => error,
    };
    assert_eq!(error.code(), Code::UnsupportedCapability);
    assert!(matches!(
        error.leaf(),
        Source::Pptx(pptx::Error::Refused(refusal))
            if refusal.kind() == pptx::SlideOrderRefusal::PresentationProperties.kind()
                && refusal.operation() == "slide_order"
    ));
}

#[test]
fn source_adoption_rejects_a_revision_from_an_independently_opened_owner() {
    let bytes = related_order_fixture();
    let original = open_source_package(bytes.clone());
    let foreign = open_source_package(bytes);
    let error = match original.adopt_source(source_owner(&foreign)) {
        Ok(_) => panic!("independently opened source owner was adopted"),
        Err(error) => error,
    };
    // A revision of another base is the caller's input (record 0351).
    assert_eq!(error.code(), Code::InvalidInput);
    assert!(matches!(
        error.leaf(),
        Source::Pptx(pptx::Error::PatchConflict)
    ));
}

#[test]
fn reopened_structural_inverse_drives_effective_catalog_and_related_reads() {
    let mut builder = pptx::builder::Builder::new().unwrap();
    for text in ["A", "B", "C"] {
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box(text, 0, 0, 914_400, 914_400);
    }
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let original = open_source_package(bytes);
    let native = source_owner(&original);
    let mut edit = native.edit().unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    let (removed, patch) = edit.commit().unwrap().into_parts();
    let inverse = pptx::source::Patch::from_bytes(&patch.inverse().to_bytes().unwrap()).unwrap();
    let reopened = open_source_package(removed.to_bytes().unwrap());
    let restored_native = source_owner(&reopened)
        .apply(&inverse)
        .unwrap()
        .into_package();
    assert!(restored_native.source_position(Position::new(1)).is_err());
    let restored = reopened.adopt_source(restored_native).unwrap();

    let cold = restored.cold_loads();
    assert_eq!(
        restored.slide_catalog().unwrap(),
        original.slide_catalog().unwrap()
    );
    assert_eq!(
        restored.cold_loads(),
        cold,
        "catalog must not decode slide payloads"
    );
    assert_eq!(restored.slides().unwrap(), original.slides().unwrap());
    for position in 0..3 {
        assert_eq!(
            restored.slide_facts(position).unwrap(),
            original.slide_facts(position).unwrap()
        );
        assert_eq!(restored.slide_notes(position).unwrap(), None);
        assert!(restored.slide_comments(position).unwrap().is_none());
        assert!(restored.slide_modern_comments(position).unwrap().is_none());
    }
    assert!(restored.slide_facts(3).unwrap().is_none());
    assert!(restored.slide_notes(3).unwrap().is_none());
    assert!(restored.slide_comments(3).unwrap().is_none());
    assert!(restored.slide_modern_comments(3).unwrap().is_none());
    let file = crate::File::Pptx(restored);
    assert_eq!(
        node_text(&file.get("/slide[2]/shape[1]").unwrap()),
        Some("B".into())
    );
    let changed = file
        .set(
            "/slide[2]/shape[1]",
            crate::props! { "text" => "restored B" },
        )
        .unwrap();
    assert_eq!(
        node_text(&changed.get("/slide[2]/shape[1]").unwrap()),
        Some("restored B".into())
    );
    let removed_again = file.remove("/slide[2]").unwrap();
    assert_eq!(
        removed_again.to_bytes().unwrap(),
        source_owner(&reopened).to_bytes().unwrap()
    );
}
