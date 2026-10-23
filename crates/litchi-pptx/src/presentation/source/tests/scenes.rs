use super::*;
use crate::shape::owned::{self, Scene};

fn with_slide(xml: &[u8]) -> Vec<u8> {
    let mut package =
        litchi_opc::OpcPackage::from_bytes(&source_backed_pptx(), SourceLimits::default()).unwrap();
    package
        .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
        .unwrap()
        .set_blob(xml.to_vec());
    litchi_opc::PackageWriter::to_bytes(&package).unwrap()
}
fn shape_xml() -> Vec<u8> {
    br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="7" name="private title"/></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>"#.to_vec()
}
fn open(bytes: Vec<u8>) -> SourceBackedPresentation {
    SourceBackedPresentation::from_read_at(Arc::new(litchi_core::OwnedSource::new(bytes))).unwrap()
}
#[test]
fn selected_scene_keeps_unrelated_payloads_cold_and_hits_do_not_read() {
    let source = Arc::new(CountingSource::new(source_backed_pptx()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let first = slide.scene().unwrap();
    let reads = source.read_calls.load(Ordering::SeqCst);
    let second = presentation.clone().slide(0).unwrap().scene().unwrap();
    assert!(Scene::ptr_eq(&first, &second));
    assert_eq!(source.read_calls.load(Ordering::SeqCst), reads);
    assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);
    assert_eq!(presentation.scene_cache_diagnostics().hits, 1);
    assert_eq!(presentation.trim_scene_cache(), 0);
    drop((first, second));
    assert_eq!(presentation.trim_scene_cache(), 1);
    assert_eq!(presentation.scene_cache_diagnostics().retained_entries, 0);
    assert!(slide.scene().is_ok());
    assert_eq!(presentation.scene_cache_diagnostics().builds, 2);
}
#[test]
fn captured_scenes_do_not_hold_the_presentation_owner() {
    let presentation = open(with_slide(&shape_xml()));
    let scene = presentation.slide(0).unwrap().scene().unwrap();
    let handle = scene.shape("private title").unwrap().unwrap();
    assert!(!format!("{:?}", presentation.scene_cache_diagnostics()).contains("private title"));
    let editor = presentation.try_into_editor().unwrap();
    drop(editor);
    assert_eq!(handle.view().name(), Some("private title"));
    assert_eq!(scene.len(), 1);
}
#[test]
fn warm_and_cold_source_revision_changes_are_typed_refusals() {
    for warm in [false, true] {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        if warm {
            slide.scene().unwrap();
        }
        source.changed();
        assert!(matches!(
            slide.scene(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
        assert_eq!(presentation.scene_cache_diagnostics().hits, 0);
    }
}
#[test]
fn managed_scenes_hold_payload_charges_without_cache_retention() {
    let source = Arc::new(litchi_core::OwnedSource::new(with_slide(&shape_xml())));
    let (budget, cancel, context) = source_backed_context(16 * 1024 * 1024);
    let presentation = SourceBackedPresentation::from_read_at_with_options(
        source,
        crate::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let slide = presentation.slide(0).unwrap();
    let first = slide.scene().unwrap();
    let second = slide.scene().unwrap();
    assert!(!Scene::ptr_eq(&first, &second));
    assert_eq!(owned::raw::xml(&first), owned::raw::xml(&second));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.unretained_builds, 2);
    assert_eq!((diagnostics.max_bytes, diagnostics.max_entries), (0, 0));
    cancel.cancel();
    assert!(matches!(
        slide.scene(),
        Err(Error::Opc(litchi_opc::OpcError::Cancelled))
    ));
    drop(slide);
    drop(presentation);
    assert!(budget.used(Resource::Memory) >= shape_xml().len() as u64);
    drop(first);
    drop(second);
    assert_eq!(budget.used(Resource::Memory), 0);
}
#[test]
fn physical_xml_limits_apply_before_mce_and_failures_are_not_cached() {
    let nesting = "<x:n>".repeat(100) + &"</x:n>".repeat(100);
    let xml = format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="x">{nesting}</mc:Choice><mc:Fallback/></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
    );
    let limits = SourceLimits::builder()
        .max_depth(64)
        .unwrap()
        .build()
        .unwrap();
    let presentation = SourceBackedPresentation::from_read_at_with_options(
        Arc::new(litchi_core::OwnedSource::new(with_slide(xml.as_bytes()))),
        crate::OpenOptions::default().source_limits(limits),
    )
    .unwrap();
    let slide = presentation.slide(0).unwrap();
    for _ in 0..2 {
        assert!(matches!(
            slide.scene(),
            Err(Error::Opc(litchi_opc::OpcError::Limit {
                kind: SourceLimitKind::Depth,
                ..
            }))
        ));
    }
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.misses, 2);
    assert_eq!(diagnostics.builds, 0);
    assert_eq!(diagnostics.retained_entries, 0);
}
#[test]
fn source_scene_matches_eager_and_mce_rewrites_remain_detached() {
    for xml in [shape_xml(), br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unknown"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="7" name="private title"/></p:nvSpPr></p:sp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#.to_vec()] {
        let bytes = with_slide(&xml);
        let presentation = open(bytes.clone());
        let source_scene = presentation.slide(0).unwrap().scene().unwrap();
        let eager = crate::Package::from_bytes(&bytes).unwrap();
        let presentation = eager.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let eager_scene = slides[0].scene().unwrap();
        assert_eq!(owned::raw::xml(&source_scene), owned::raw::xml(&eager_scene));
        assert_eq!(source_scene.shape("private title").unwrap().unwrap().view().name(), Some("private title"));
    }
}

struct ArmedRevision {
    inner: CountingSource,
    checks: AtomicUsize,
    change_on: AtomicUsize,
}
impl ReadAt for ArmedRevision {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read_at(offset, output)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        let check = self.checks.fetch_add(1, Ordering::SeqCst) + 1;
        if check == self.change_on.load(Ordering::SeqCst) {
            self.inner.changed();
        }
        self.inner.version()
    }
}
#[test]
fn a_cache_hit_checks_the_revision_again_before_returning() {
    let source = Arc::new(ArmedRevision {
        inner: CountingSource::new(source_backed_pptx()),
        checks: AtomicUsize::new(0),
        change_on: AtomicUsize::new(usize::MAX),
    });
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let retained = slide.scene().unwrap();
    source
        .change_on
        .store(source.checks.load(Ordering::SeqCst) + 2, Ordering::SeqCst);
    assert!(matches!(
        slide.scene(),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
    ));
    assert_eq!(presentation.scene_cache_diagnostics().hits, 1);
    // Already captured immutable snapshots remain usable.
    assert!(retained.is_empty());
}

#[test]
fn a_failed_post_publication_fence_rolls_back_the_exact_scene() {
    let source = || {
        Arc::new(ArmedRevision {
            inner: CountingSource::new(source_backed_pptx()),
            checks: AtomicUsize::new(0),
            change_on: AtomicUsize::new(usize::MAX),
        })
    };
    let probe = source();
    let presentation = SourceBackedPresentation::from_read_at(probe.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    let before = probe.checks.load(Ordering::SeqCst);
    slide.scene().unwrap();
    let checks = probe.checks.load(Ordering::SeqCst) - before;
    assert!(checks > 2);
    let source = source();
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let slide = presentation.slide(0).unwrap();
    source.change_on.store(
        source.checks.load(Ordering::SeqCst) + checks,
        Ordering::SeqCst,
    );
    assert!(matches!(
        slide.scene(),
        Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
    ));
    let diagnostics = presentation.scene_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.fence_discards, 1);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
}
