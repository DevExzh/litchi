#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration assertions intentionally panic on an invalid fixture or result"
)]

use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    Resource,
};
use litchi_opc::{OpcPackage, PackURI, SourceLimits};
use litchi_pptx::{Error, builder::Builder, source::Package};

macro_rules! assert_refused_kind {
    ($operation:expr, $expected:expr) => {{
        match $operation {
            Err(Error::Refused(refusal)) => {
                assert_eq!(refusal.kind(), $expected);
            },
            Err(error) => panic!("unexpected error: {error:?}"),
            Ok(_) => panic!("operation unexpectedly succeeded"),
        }
    }};
}

const SIGNED_FIXTURE: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx");
const MODELED_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:p1710="http://schemas.microsoft.com/office/powerpoint/2017/10/main" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:showPr><p:present/><p:sldAll/><p:extLst><p:ext uri="{F99C55AA-B7CB-42B0-86F8-08522FDF87E8}"><p14:browseMode showStatus="0"/></p:ext><p:ext uri="{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}"><p14:laserClr><a:srgbClr val="FF0000"/></p14:laserClr></p:ext><p:ext uri="{2FDB2607-1784-4EEB-B798-7EB5836EED8A}"><p14:showMediaCtrls val="1"/></p:ext></p:extLst></p:showPr><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData val="0"/></p:ext><p:ext uri="{D31A062A-798A-4329-ABDD-BBA856620510}"><p14:defaultImageDpi val="220"/></p:ext><p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased val="1"/></p:ext><p:ext uri="{4599F94E-CEE6-441E-89CC-EB005ECD8F06}"><a14:m><m:mathPr><m:brkBin m:val="repeat"/></m:mathPr></a14:m></p:ext><p:ext uri="{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}"><p1710:readonlyRecommended val="1"/></p:ext></p:extLst></p:presentationPr>"#;

const UNKNOWN_EXTENSION_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:unknown"><p:extLst><p:ext uri="urn:unknown"><x:payload/></p:ext></p:extLst></p:presentationPr>"#;
const NESTED_EXTENSION_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData val="0"><p14:nested/></p14:discardImageEditData></p:ext></p:extLst></p:presentationPr>"#;
const MCE_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p"><mc:AlternateContent><mc:Choice Requires="p"><p:showPr/></mc:Choice><mc:Fallback><p:showPr/></mc:Fallback></mc:AlternateContent></p:presentationPr>"#;
const WRONG_URI_PAYLOAD_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:defaultImageDpi val="220"/></p:ext></p:extLst></p:presentationPr>"#;
const DUPLICATE_EXTENSION_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:extLst><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData val="0"/></p:ext><p:ext uri="{E76CE94A-603C-4142-B9EB-6D1370010A27}"><p14:discardImageEditData val="1"/></p:ext></p:extLst></p:presentationPr>"#;
const FOREIGN_LASER_CHILD_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:foreign"><p:showPr><p:extLst><p:ext uri="{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}"><p14:laserClr><a:srgbClr val="112233"><x:tint val="1000"/></a:srgbClr></p14:laserClr></p:ext></p:extLst></p:showPr></p:presentationPr>"#;
const RELATIONSHIP_LASER_ATTRIBUTE_PRES_PROPS: &[u8] = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:showPr><p:extLst><p:ext uri="{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}"><p14:laserClr><a:srgbClr val="112233" r:id="rId2"/></p14:laserClr></p:ext></p:extLst></p:showPr></p:presentationPr>"#;
const CUSTOM_SHOW_SUFFIX: &str = r#"<p:custShowLst><p:custShow name="Opening" id="7"><p:sldLst><p:sld r:id="rId2"/><p:sld r:id="rId4"/></p:sldLst></p:custShow></p:custShowLst>"#;

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        for text in ["A", "B", "C"] {
            presentation
                .add_slide()
                .unwrap()
                .add_text_box(text, 0, 0, 914_400, 914_400);
        }
    }
    let authored = builder.build().unwrap();
    authored.to_bytes().unwrap()
}

fn fresh(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn managed_context(memory_limit: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-source-primary-order-test",
        Limits::new(
            memory_limit,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory_limit.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn presentation_root_len(bytes: &[u8]) -> u64 {
    OpcPackage::from_bytes(bytes, SourceLimits::default())
        .unwrap()
        .get_part(&PackURI::new("/ppt/presentation.xml").unwrap())
        .unwrap()
        .blob()
        .len() as u64
}

fn with_part(bytes: &[u8], path: &str, xml: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    package
        .get_part_mut(&PackURI::new(path).unwrap())
        .unwrap()
        .set_blob(xml.to_vec());
    litchi_opc::PackageWriter::to_bytes(&package).unwrap()
}

fn with_pres_props(bytes: &[u8], xml: &[u8]) -> Vec<u8> {
    with_part(bytes, "/ppt/presProps.xml", xml)
}

fn with_presentation_root_suffix(bytes: &[u8], suffix: &str) -> Vec<u8> {
    let original = String::from_utf8(member_bytes(bytes, "/ppt/presentation.xml")).unwrap();
    let updated = original.replacen(
        "</p:presentation>",
        &format!("{suffix}</p:presentation>"),
        1,
    );
    with_part(bytes, "/ppt/presentation.xml", updated.as_bytes())
}

fn member_bytes(bytes: &[u8], path: &str) -> Vec<u8> {
    OpcPackage::from_bytes(bytes, SourceLimits::default())
        .unwrap()
        .get_part(&PackURI::new(path).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn reordered_bytes(bytes: &[u8]) -> Vec<u8> {
    let source = fresh(bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    edit.commit().unwrap().package().to_bytes().unwrap()
}

fn texts(package: &Package) -> Vec<String> {
    let presentation = package.presentation().unwrap();
    (0..presentation.slide_count().unwrap())
        .map(|position| presentation.slide(position).unwrap().text().unwrap())
        .collect()
}

fn source_positions(package: &Package) -> Vec<Position> {
    let count = package.presentation().unwrap().slide_count().unwrap();
    (0..count)
        .map(|position| package.source_position(Position::new(position)).unwrap())
        .collect()
}

#[test]
fn logical_reads_and_source_identity_follow_composed_mixed_edits() {
    let bytes = fixture();
    let source = fresh(&bytes);
    assert_eq!(texts(&source), ["A", "B", "C"]);
    assert_eq!(
        source_positions(&source),
        [Position::new(0), Position::new(1), Position::new(2)]
    );

    let mut edit = source.edit().unwrap();
    // This text edit addresses physical A before any order change.
    assert!(edit.set_shape_text(0, 0_usize, "A*").unwrap());
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    // The staged order is now [B, C, A*], so this edit addresses B.
    assert!(edit.set_shape_text(0, 0_usize, "B*").unwrap());
    assert!(
        edit.swap_slides(Position::new(0), Position::new(1))
            .unwrap()
    );

    let commit = edit.commit().unwrap();
    let changed = commit.package();
    assert!(commit.diagnostics().changed());
    assert_eq!(texts(changed), ["C", "B*", "A*"]);
    assert_eq!(
        source_positions(changed),
        [Position::new(2), Position::new(1), Position::new(0)]
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn cancelled_order_edit_is_a_noop_with_exact_zip_bytes() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(edit.move_slide(Position::new(2), Position::new(0)).unwrap());

    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), bytes);
    assert_eq!(texts(commit.package()), ["A", "B", "C"]);
}

#[test]
fn order_patch_replays_after_reopen_and_inverse_restores_original() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(edit.set_shape_text(0, 0_usize, "B*").unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&changed), ["B*", "C", "A"]);
    assert_eq!(
        source_positions(&changed),
        [Position::new(1), Position::new(2), Position::new(0)]
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);

    let durable = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let reopened = fresh(&bytes);
    let replayed = reopened.apply(&durable).unwrap().into_package();
    assert_eq!(replayed.to_bytes().unwrap(), changed.to_bytes().unwrap());
    assert_eq!(texts(&replayed), ["B*", "C", "A"]);
    assert_eq!(
        source_positions(&replayed),
        [Position::new(1), Position::new(2), Position::new(0)]
    );

    let inverse =
        litchi_pptx::source::Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    let reopened_changed = fresh(&replayed.to_bytes().unwrap());
    let restored = reopened_changed.apply(&inverse).unwrap().into_package();
    assert_eq!(restored.to_bytes().unwrap(), bytes);
    assert_eq!(texts(&restored), ["A", "B", "C"]);
    assert_eq!(
        source_positions(&restored),
        [Position::new(2), Position::new(0), Position::new(1)]
    );
}

#[test]
fn text_patch_authored_after_order_refuses_the_original_revision() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut order_edit = source.edit().unwrap();
    assert!(
        order_edit
            .move_slide(Position::new(0), Position::new(2))
            .unwrap()
    );
    let ordered = order_edit.commit().unwrap().into_package();

    let mut text_edit = ordered.edit().unwrap();
    assert!(text_edit.set_shape_text(0, 0_usize, "B*").unwrap());
    let text_patch = text_edit.commit().unwrap().patch().clone();
    let durable = litchi_pptx::source::Patch::from_bytes(&text_patch.to_bytes().unwrap()).unwrap();
    assert!(matches!(source.apply(&durable), Err(Error::PatchConflict)));

    let reopened_ordered = fresh(&ordered.to_bytes().unwrap());
    let accepted = reopened_ordered.apply(&durable).unwrap().into_package();
    assert_eq!(texts(&accepted), ["B*", "C", "A"]);
    assert_eq!(
        source_positions(&accepted),
        [Position::new(0), Position::new(1), Position::new(2)]
    );
}

#[test]
fn failed_later_order_step_keeps_the_prior_staged_state_atomic() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(
        edit.swap_slides(Position::new(3), Position::new(0))
            .is_err()
    );
    assert!(source.source_position(Position::new(3)).is_err());

    let changed = edit.commit().unwrap().into_package();
    assert_eq!(texts(&changed), ["B", "C", "A"]);
    assert_eq!(
        source_positions(&changed),
        [Position::new(1), Position::new(2), Position::new(0)]
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn signed_noop_order_is_exact_but_changed_primary_edits_are_refused() {
    let source = fresh(SIGNED_FIXTURE);
    let mut noop = source.edit().unwrap();
    assert!(!noop.move_slide(Position::new(0), Position::new(0)).unwrap());
    assert!(
        !noop
            .swap_slides(Position::new(0), Position::new(0))
            .unwrap()
    );
    let noop = noop.commit().unwrap();
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().to_bytes().unwrap(), SIGNED_FIXTURE);
    assert_eq!(
        source
            .apply(noop.patch())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        SIGNED_FIXTURE
    );

    let mut changed = source.edit().unwrap();
    assert!(matches!(
        changed.set_shape_text(0, 0_usize, "changed"),
        Err(Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert!(changed.commit().unwrap().patch().is_noop());
    assert_eq!(source.to_bytes().unwrap(), SIGNED_FIXTURE);
}

#[test]
fn order_rewrite_preserves_modeled_properties_and_custom_show_bytes() {
    let stock = fixture();
    let stock_pres_props = member_bytes(&stock, "/ppt/presProps.xml");
    let stock_reordered = reordered_bytes(&stock);
    assert_eq!(
        member_bytes(&stock_reordered, "/ppt/presProps.xml"),
        stock_pres_props
    );

    let with_custom_show = with_presentation_root_suffix(&stock, CUSTOM_SHOW_SUFFIX);
    let with_modeled_properties = with_pres_props(&with_custom_show, MODELED_PRES_PROPS);
    let before_pres_props = member_bytes(&with_modeled_properties, "/ppt/presProps.xml");
    let reordered = reordered_bytes(&with_modeled_properties);
    assert_eq!(
        member_bytes(&reordered, "/ppt/presProps.xml"),
        before_pres_props
    );

    let root = String::from_utf8(member_bytes(&reordered, "/ppt/presentation.xml")).unwrap();
    assert_eq!(root.matches(CUSTOM_SHOW_SUFFIX).count(), 1);
}

#[test]
fn standalone_source_order_publishes_modeled_properties_and_custom_show() {
    let stock = fixture();
    let with_custom_show = with_presentation_root_suffix(&stock, CUSTOM_SHOW_SUFFIX);
    let bytes = with_pres_props(&with_custom_show, MODELED_PRES_PROPS);
    let before_pres_props = member_bytes(&bytes, "/ppt/presProps.xml");
    let package = Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let mut edit = package.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.package().write_to(&mut output).unwrap();

    assert_eq!(
        member_bytes(&output, "/ppt/presProps.xml"),
        before_pres_props
    );
    let root = String::from_utf8(member_bytes(&output, "/ppt/presentation.xml")).unwrap();
    assert_eq!(root.matches(CUSTOM_SHOW_SUFFIX).count(), 1);
    assert_eq!(texts(&fresh(&output)), ["B", "C", "A"]);
}

#[test]
fn order_refuses_unsafe_presentation_properties_atomically() {
    let cases = [
        ("unknown extension", UNKNOWN_EXTENSION_PRES_PROPS),
        ("nested extension payload", NESTED_EXTENSION_PRES_PROPS),
        ("markup compatibility content", MCE_PRES_PROPS),
        ("wrong extension payload", WRONG_URI_PAYLOAD_PRES_PROPS),
        ("duplicate extension", DUPLICATE_EXTENSION_PRES_PROPS),
        ("foreign laser child", FOREIGN_LASER_CHILD_PRES_PROPS),
        (
            "relationship laser attribute",
            RELATIONSHIP_LASER_ATTRIBUTE_PRES_PROPS,
        ),
    ];

    for (label, properties) in cases {
        let bytes = with_pres_props(&fixture(), properties);
        let source = fresh(&bytes);
        let before = source.to_bytes().unwrap();
        let mut edit = source.edit().unwrap();
        assert_refused_kind!(
            edit.move_slide(Position::new(0), Position::new(2)),
            litchi_pptx::SlideOrderRefusal::PresentationProperties.kind()
        );
        let commit = edit.commit().unwrap();
        assert!(
            commit.patch().is_noop(),
            "failed order changed the patch: {label}"
        );
        assert_eq!(commit.package().to_bytes().unwrap(), before, "{label}");
        assert_eq!(source.to_bytes().unwrap(), before, "{label}");
    }
}

#[test]
fn order_edit_refuses_when_only_the_source_root_fits_memory() {
    let bytes = fixture();
    let root = presentation_root_len(&bytes);
    let (budget, _cancellation, context) = managed_context(root);
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    assert_eq!(budget.used(Resource::Memory), root);

    let mut edit = source.edit().unwrap();
    assert!(matches!(
        edit.move_slide(Position::new(0), Position::new(2)),
        Err(Error::Opc(litchi_opc::OpcError::Execution(
            litchi_core::ExecutionError::ResourceLimit(limit)
        ))) if limit.resource == Resource::Memory
    ));
    assert_eq!(budget.used(Resource::Memory), root);
    drop(edit);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancelled_order_edit_releases_its_root_reservation() {
    let bytes = fixture();
    let root = presentation_root_len(&bytes);
    let (budget, cancellation, context) = managed_context(16 * 1024 * 1024);
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    assert_eq!(budget.used(Resource::Memory), root);

    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    // Order admission also retains the inspected positional metadata parts.
    let with_order = budget.used(Resource::Memory);
    assert!(with_order >= root * 2);
    assert!(edit.move_slide(Position::new(2), Position::new(0)).unwrap());
    assert_eq!(budget.used(Resource::Memory), with_order - root);
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert_eq!(budget.used(Resource::Memory), with_order);
    cancellation.cancel();
    assert!(edit.commit().is_err());
    assert_eq!(budget.used(Resource::Memory), with_order - root);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn order_patch_and_inverse_clones_retain_root_until_the_last_drop() {
    let bytes = fixture();
    let root = presentation_root_len(&bytes);
    let (budget, _cancellation, context) = managed_context(16 * 1024 * 1024);
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    drop(commit);
    drop(source);

    let held = budget.used(Resource::Memory);
    assert!(held >= root * 2);
    drop(patch);
    assert_eq!(budget.used(Resource::Memory), held);
    drop(inverse);
    assert_eq!(budget.used(Resource::Memory), 0);
}
