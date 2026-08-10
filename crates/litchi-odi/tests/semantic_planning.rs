#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_odi::{
    Builder, ConflictKind, FlatImage, FrameProperty, History, Image, OperationKey,
    ProtectionDisposition, SecurityPolicy, SemanticValue, StyleDependencyState, SurfaceDisposition,
    SurfaceKind,
    frame::Frame,
    map::{Area, ImageMap},
    source::Source,
};

const FLAT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:mimetype="application/vnd.oasis.opendocument.image"><office:meta><dc:title>Before</dc:title><meta:user-defined meta:name="opaque">keep</meta:user-defined></office:meta><office:body><office:image>"#,
    r#"<draw:frame draw:name="Photo" draw:style-name="gr1" draw:layer="layout" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:image xlink:href="Pictures/photo.png"/></draw:frame>"#,
    r#"</office:image></office:body></office:document>"#,
);

const META: &str = r#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"><office:meta><dc:title>Before</dc:title><meta:user-defined meta:name="opaque">keep</meta:user-defined></office:meta></office:document-meta>"#;

const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:styles><style:style style:name="gr2" style:family="graphic"/></office:styles></office:document-styles>"#;

const TRANSITIVE_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="auto-child" style:family="graphic" style:parent-style-name="named-parent"/></office:automatic-styles><office:body><office:image><draw:frame draw:style-name="gr1"><draw:image draw:mime-type="image/png"><office:binary-data>iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X9p9WQAAAABJRU5ErkJggg==</office:binary-data></draw:image></draw:frame></office:image></office:body></office:document-content>"#;

const TRANSITIVE_STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:styles><style:style style:name="named-parent" style:family="graphic"/></office:styles></office:document-styles>"#;

const NORMATIVE_SYNTHETIC: &[u8] = include_bytes!("fixtures/odf-1.4-normative-synthetic.fodi");
const ODFDOM_ORIGINAL: &[u8] = include_bytes!("fixtures/odfdom-0.13.0-original.odi");
const ODFDOM_CHANGED: &[u8] = include_bytes!("fixtures/odfdom-0.13.0-changed.odi");

fn flat() -> FlatImage {
    FlatImage::from_bytes(FLAT.as_bytes().to_vec()).unwrap()
}

fn package() -> Image {
    let frame = Frame::new(Source::Linked("Pictures/photo.png".into()))
        .with_name("Photo")
        .with_style_name("gr1")
        .with_layer("layout")
        .with_geometry("1cm", "2cm", "3cm", "4cm");
    Image::from_bytes(
        Builder::new()
            .frame(&frame)
            .meta_xml(META)
            .resource("Pictures/photo.png", "image/png", b"image".to_vec())
            .build()
            .unwrap(),
    )
    .unwrap()
}

fn map() -> ImageMap {
    ImageMap::new(vec![
        Area::rectangle("0cm", "0cm", "1cm", "1cm")
            .with_href("https://example.test/one")
            .with_target_frame_name("_blank")
            .with_show("new")
            .with_actuate("onRequest")
            .with_title("One"),
        Area::circle("2cm", "2cm", "1cm").with_no_href(),
    ])
}

#[test]
fn retained_frame_image_and_style_semantics_share_the_durable_wire() {
    let source = flat();
    let mut edit = source.transaction();
    edit.set_frame_xml_id(0, Some("frame-xml-id".into()))
        .unwrap();
    edit.set_frame_title(0, Some("Accessible title".into()))
        .unwrap();
    edit.set_frame_description(0, Some("Accessible description".into()))
        .unwrap();
    edit.set_image_media_type(0, Some("image/png".into()))
        .unwrap();
    edit.set_image_xml_id(0, Some("image-xml-id".into()))
        .unwrap();
    edit.set_filter_name(0, Some("producer-filter".into()))
        .unwrap();
    edit.set_link_type(0, Some("simple".into())).unwrap();
    edit.set_show(0, Some("embed".into())).unwrap();
    edit.set_actuate(0, Some("onLoad".into())).unwrap();
    edit.set_copy_of(0, Some("OtherFrame".into())).unwrap();
    edit.set_image_map(0, Some(map())).unwrap();
    let commit = edit.commit().unwrap();
    let frame = commit.snapshot().frame().unwrap();
    assert_eq!(frame.xml_id(), Some("frame-xml-id"));
    assert_eq!(frame.title(), Some("Accessible title"));
    assert_eq!(frame.description(), Some("Accessible description"));
    assert_eq!(frame.media_type(), Some("image/png"));
    assert_eq!(frame.image_xml_id(), Some("image-xml-id"));
    assert_eq!(frame.filter_name(), Some("producer-filter"));
    assert_eq!(frame.link_type(), Some("simple"));
    assert_eq!(frame.show(), Some("embed"));
    assert_eq!(frame.actuate(), Some("onLoad"));
    assert_eq!(frame.copy_of(), Some("OtherFrame"));

    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();
    assert_eq!(patch.operations().len(), 11);
    let package_source = package();
    let transfer = patch.plan_package(&package_source);
    assert!(transfer.is_conflict_free());
    let transferred = transfer
        .commit_package(&package_source, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(transferred.image().frame().unwrap(), frame);

    let mut styles_edit = package_source.edit();
    styles_edit.set_styles_xml(Some(STYLES.into())).unwrap();
    let styles_commit = styles_edit.commit().unwrap();
    assert_eq!(styles_commit.image().styles_xml(), Some(STYLES));
    let styles_patch = styles_commit
        .semantic_patch(&SecurityPolicy::default())
        .unwrap();
    assert!(
        styles_patch
            .operations()
            .iter()
            .any(|operation| operation.key() == &OperationKey::Styles)
    );
    assert!(
        styles_commit
            .semantic_patch(&SecurityPolicy::default().with_xml_bytes(1))
            .is_err()
    );
    assert_eq!(
        styles_patch
            .inverse()
            .apply_package(styles_commit.image())
            .unwrap()
            .as_bytes(),
        package_source.as_bytes()
    );

    let mut remove = styles_commit.image().edit();
    remove.set_styles_xml(None).unwrap();
    assert_eq!(remove.commit().unwrap().image().styles_xml(), None);

    let mut noncompact = package_source.edit();
    assert!(
        noncompact
            .set_styles_xml(Some(STYLES.replace("><", ">\n<")))
            .is_err()
    );
}

#[test]
fn checked_in_normative_synthetic_fodi_is_explicitly_non_producer_evidence() {
    let image = FlatImage::from_bytes(NORMATIVE_SYNTHETIC.to_vec()).unwrap();
    assert_eq!(
        image.metadata().unwrap().title.as_deref(),
        Some("Normative synthetic ODI evidence")
    );
    assert_eq!(image.frame().unwrap().name(), Some("Synthetic"));
    assert_eq!(image.frame().unwrap().image_map().unwrap().areas().len(), 2);
    assert_eq!(image.as_bytes(), NORMATIVE_SYNTHETIC);
}

#[test]
fn checked_in_odfdom_producer_round_trip_is_genuine_odi_evidence() {
    let original = Image::from_bytes(ODFDOM_ORIGINAL.to_vec()).unwrap();
    let changed = Image::from_bytes(ODFDOM_CHANGED.to_vec()).unwrap();
    assert_eq!(
        original.frame().unwrap().name(),
        Some("ODFDOM-0.13.0-Original")
    );
    assert_eq!(
        changed.frame().unwrap().name(),
        Some("ODFDOM-0.13.0-Changed")
    );
    assert!(matches!(
        original.frame().unwrap().source(),
        Source::Embedded(_)
    ));
    assert!(matches!(
        changed.frame().unwrap().source(),
        Source::Embedded(_)
    ));
    assert_eq!(original.as_bytes(), ODFDOM_ORIGINAL);
    assert_eq!(changed.as_bytes(), ODFDOM_CHANGED);
}

#[test]
fn public_form_and_extension_inventory_is_inert_and_exact() {
    let xml = FLAT
        .replacen(
            " office:mimetype=",
            r#" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:vendor="urn:vendor:test" office:mimetype="#,
            1,
        )
        .replacen(
            "<office:body>",
            r#"<office:forms><form:form form:name="Search"><form:text form:id="query"/></form:form></office:forms><vendor:item vendor:flag="1"/><office:body>"#,
            1,
        );
    let image = FlatImage::from_bytes(xml.clone().into_bytes()).unwrap();
    assert_eq!(image.forms().len(), 3);
    assert_eq!(image.forms()[1].name(), Some("Search"));
    assert_eq!(image.forms()[2].control_id(), Some("query"));
    assert_eq!(image.extensions().len(), 2);
    assert_eq!(image.extensions()[0].kind(), SurfaceKind::ExtensionElement);
    assert_eq!(
        image.extensions()[1].preservation_disposition(),
        SurfaceDisposition::PreserveExact
    );
    assert_eq!(image.as_bytes(), xml.as_bytes());
}

#[test]
fn automatic_style_transfer_carries_transitive_exact_parent_closure() {
    let prepared = Image::from_bytes(
        Builder::new()
            .content_xml(TRANSITIVE_CONTENT)
            .styles_xml(TRANSITIVE_STYLES)
            .build()
            .unwrap(),
    )
    .unwrap();
    let mut edit = prepared.edit();
    edit.set_style_name(0, Some("auto-child".into())).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();
    let destination = package();
    let plan = patch.plan_package(&destination);
    assert!(plan.is_conflict_free(), "{:?}", plan.conflicts());
    let transferred = plan
        .commit_package(&destination, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(
        transferred
            .image()
            .style_dependency_state("auto-child", "graphic")
            .unwrap(),
        StyleDependencyState::Automatic
    );
    assert_eq!(
        transferred
            .image()
            .style_dependency_state("named-parent", "graphic")
            .unwrap(),
        StyleDependencyState::Named
    );
    let styles = transferred.image().styles_xml().unwrap();
    assert!(styles.contains(r#"style:name="auto-child""#));
    assert!(styles.contains(r#"style:name="named-parent""#));
}

#[test]
fn protection_inventory_exposes_exact_signed_rewrite_disposition() {
    let signed = Image::from_bytes(
        Builder::new()
            .resource(
                "META-INF/vendor-signatures-v2.xml",
                "text/xml",
                br"<signature/>".to_vec(),
            )
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(
        signed.protection().signature_members(),
        &["META-INF/vendor-signatures-v2.xml".to_owned()]
    );
    assert_eq!(
        signed.protection().disposition(),
        ProtectionDisposition::RefuseSignedRewrite
    );
    assert!(!signed.protection().is_encrypted());
}

#[test]
fn semantic_patch_is_deterministic_exact_reopenable_and_transferable() {
    let source = flat();
    let mut edit = source.transaction();
    edit.set_style_name(0, Some("gr2".into())).unwrap();
    edit.set_x(0, Some("5cm".into())).unwrap();
    edit.set_image_map(0, Some(map())).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();

    assert!(
        patch
            .operations()
            .windows(2)
            .all(|pair| pair[0].key() < pair[1].key())
    );
    assert!(patch.operations().iter().any(|operation| {
        operation.key()
            == &OperationKey::Frame {
                frame: 0,
                property: FrameProperty::ImageMap,
            }
    }));
    let exact = patch.apply_flat(&source).unwrap();
    assert_eq!(exact.as_bytes(), commit.snapshot().as_bytes());
    assert_eq!(
        patch.inverse().apply_flat(&exact).unwrap().as_bytes(),
        source.as_bytes()
    );
    let stale = FlatImage::from_bytes(FLAT.replace("gr1", "other").into_bytes()).unwrap();
    assert!(patch.apply_flat(&stale).is_err());

    let package_source = {
        let package_without_styles = package();
        let mut style_setup = package_without_styles.edit();
        style_setup.set_styles_xml(Some(STYLES.into())).unwrap();
        style_setup.commit().unwrap().image().clone()
    };
    let transfer = patch.plan_package(&package_source);
    assert!(transfer.is_conflict_free());
    let package_commit = transfer
        .commit_package(&package_source, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(
        package_commit.image().frame().unwrap().style_name(),
        Some("gr2")
    );
    assert_eq!(package_commit.image().frame().unwrap().x(), Some("5cm"));
    assert_eq!(
        package_commit
            .image()
            .frame()
            .unwrap()
            .image_map()
            .unwrap()
            .areas()
            .len(),
        2
    );
    let first_area = &package_commit
        .image()
        .frame()
        .unwrap()
        .image_map()
        .unwrap()
        .areas()[0];
    assert_eq!(first_area.target_frame_name(), Some("_blank"));
    assert_eq!(first_area.show(), Some("new"));
    assert_eq!(first_area.actuate(), Some("onRequest"));
}

#[test]
fn package_transfer_copies_complete_resource_and_style_dependencies() {
    let base = package();
    let mut prepare = base.edit();
    prepare.set_styles_xml(Some(STYLES.into())).unwrap();
    prepare
        .put_member(
            "Pictures/replacement.png".into(),
            "image/png".into(),
            b"replacement".to_vec(),
        )
        .unwrap();
    let prepared = prepare.commit().unwrap().image().clone();

    let mut edit = prepared.edit();
    edit.set_source(0, Source::Linked("Pictures/replacement.png".into()))
        .unwrap();
    edit.set_style_name(0, Some("gr2".into())).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();

    let destination = package();
    let plan = patch.plan_package(&destination);
    assert!(plan.is_conflict_free(), "{:?}", plan.conflicts());
    assert!(
        plan.operations()
            .iter()
            .any(|operation| operation.key() == &OperationKey::Styles)
    );
    assert!(plan.operations().iter().any(|operation| {
        operation.key() == &OperationKey::Resource("Pictures/replacement.png".into())
    }));
    let transferred = plan
        .commit_package(&destination, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(transferred.image().styles_xml(), Some(STYLES));
    assert_eq!(
        transferred
            .image()
            .member_bytes("Pictures/replacement.png")
            .unwrap(),
        Some(b"replacement".to_vec())
    );
    assert_eq!(
        transferred.image().frame().unwrap().source(),
        &Source::Linked("Pictures/replacement.png".into())
    );
    assert!(
        plan.commit_package(
            &destination,
            &SecurityPolicy::default().with_resource_bytes(1),
        )
        .is_err()
    );
    assert!(
        plan.commit_package(&destination, &SecurityPolicy::default().with_patch_bytes(1),)
            .is_err()
    );

    let mut seed_mismatch = destination.edit();
    seed_mismatch
        .put_member(
            "Pictures/replacement.png".into(),
            "image/png".into(),
            b"different".to_vec(),
        )
        .unwrap();
    let mismatched = seed_mismatch.commit().unwrap().image().clone();
    let refused = patch.plan_package(&mismatched);
    assert!(refused.conflicts().iter().any(|conflict| {
        conflict.kind() == ConflictKind::MissingDependency
            && conflict.key() == Some(&OperationKey::Resource("Pictures/replacement.png".into()))
    }));
    assert!(
        refused
            .commit_package(&mismatched, &SecurityPolicy::default())
            .is_err()
    );
    assert_eq!(
        mismatched.member_bytes("Pictures/replacement.png").unwrap(),
        Some(b"different".to_vec())
    );
}

#[test]
fn flat_transfer_reports_dependencies_it_cannot_supply() {
    let source = flat();
    let mut edit = source.transaction();
    edit.set_source(0, Source::Linked("Pictures/absent.png".into()))
        .unwrap();
    edit.set_style_name(0, Some("absent-style".into())).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();
    let destination = package();
    let plan = patch.plan_package(&destination);

    assert!(!plan.is_conflict_free());
    assert_eq!(
        plan.conflicts()
            .iter()
            .filter(|conflict| conflict.kind() == ConflictKind::MissingDependency)
            .count(),
        2
    );
    assert!(
        plan.commit_package(&destination, &SecurityPolicy::default())
            .is_err()
    );
    assert_eq!(destination.frame().unwrap().style_name(), Some("gr1"));
}

#[test]
fn compatible_join_and_divergent_three_way_plans_are_non_mutating() {
    let source = flat();
    let mut style_edit = source.transaction();
    style_edit
        .set_style_name(0, Some("joined-style".into()))
        .unwrap();
    let style = style_edit.commit().unwrap();
    let style_patch = style.semantic_patch(&SecurityPolicy::default()).unwrap();

    let mut layer_edit = source.transaction();
    layer_edit
        .set_layer(0, Some("joined-layer".into()))
        .unwrap();
    let layer = layer_edit.commit().unwrap();
    let layer_patch = layer.semantic_patch(&SecurityPolicy::default()).unwrap();

    let joined = style_patch.join(&layer_patch);
    assert!(joined.is_conflict_free());
    assert_eq!(joined.operations().len(), 2);
    assert_eq!(source.frame().unwrap().style_name(), Some("gr1"));
    let committed = joined
        .commit_flat(&source, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(
        committed.snapshot().frame().unwrap().style_name(),
        Some("joined-style")
    );
    assert_eq!(
        committed.snapshot().frame().unwrap().layer(),
        Some("joined-layer")
    );

    let mut competing_edit = source.transaction();
    competing_edit
        .set_style_name(0, Some("competing".into()))
        .unwrap();
    let competing = competing_edit.commit().unwrap();
    let competing_patch = competing
        .semantic_patch(&SecurityPolicy::default())
        .unwrap();
    let conflict = style_patch.join(&competing_patch);
    assert!(!conflict.is_conflict_free());
    assert_eq!(conflict.conflicts()[0].kind(), ConflictKind::Diverged);

    let three_way = style_patch.plan_flat(competing.snapshot());
    assert_eq!(three_way.conflicts()[0].kind(), ConflictKind::Diverged);
    assert!(matches!(
        three_way.conflicts()[0].actual(),
        Some(SemanticValue::Text(Some(value))) if value == "competing"
    ));
    assert!(
        joined
            .commit_flat(competing.snapshot(), &SecurityPolicy::default())
            .is_err()
    );
}

#[test]
fn package_metadata_and_resource_crud_have_transferable_inverse_patches() {
    let source = package();
    let mut edit = source.edit();
    edit.set_title(Some("After".into())).unwrap();
    edit.put_member(
        "Thumbnails/preview.webp".into(),
        "image/webp".into(),
        b"preview".to_vec(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();
    assert_eq!(patch.operations().len(), 2);
    assert_eq!(
        patch.apply_package(&source).unwrap().as_bytes(),
        commit.image().as_bytes()
    );
    assert_eq!(
        patch
            .inverse()
            .apply_package(commit.image())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
    assert!(
        commit
            .image()
            .meta_xml()
            .unwrap()
            .unwrap()
            .contains("opaque")
    );
    assert_eq!(
        commit
            .image()
            .member_bytes("Thumbnails/preview.webp")
            .unwrap(),
        Some(b"preview".to_vec())
    );

    let planned = patch.plan_package(&source);
    assert!(planned.is_conflict_free());
    let replay = planned
        .commit_package(&source, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(
        replay.image().metadata().unwrap().title.as_deref(),
        Some("After")
    );

    let mut removal = commit.image().edit();
    removal
        .remove_member("Thumbnails/preview.webp".into())
        .unwrap();
    let removed = removal.commit().unwrap();
    assert_eq!(
        removed
            .image()
            .member_bytes("Thumbnails/preview.webp")
            .unwrap(),
        None
    );
}

#[test]
fn flat_metadata_patch_preserves_opaque_nodes_and_transfers_to_package() {
    let source = flat();
    assert_eq!(source.metadata().unwrap().title.as_deref(), Some("Before"));
    let mut edit = source.transaction();
    edit.set_title(Some("Flat After".into())).unwrap();
    edit.set_author(Some("Flat Author".into())).unwrap();
    let commit = edit.commit().unwrap();
    let xml = std::str::from_utf8(commit.snapshot().as_bytes()).unwrap();
    assert!(xml.contains("opaque"));
    let patch = commit.semantic_patch(&SecurityPolicy::default()).unwrap();
    assert_eq!(patch.operations().len(), 2);

    let package = package();
    let transfer = patch.plan_package(&package);
    assert!(transfer.is_conflict_free());
    let transferred = transfer
        .commit_package(&package, &SecurityPolicy::default())
        .unwrap();
    assert_eq!(
        transferred.image().metadata().unwrap().title.as_deref(),
        Some("Flat After")
    );
    assert_eq!(
        transferred.image().metadata().unwrap().author.as_deref(),
        Some("Flat Author")
    );
}

#[test]
fn commit_coupled_history_is_byte_bounded_and_stale_safe() {
    let source = flat();
    let mut edit = source.transaction();
    edit.set_layer(0, Some("next".into())).unwrap();
    let commit = edit.commit().unwrap();
    let budget = source.as_bytes().len() + commit.snapshot().as_bytes().len();
    let mut history = History::with_byte_budget(source.clone(), 4, budget).unwrap();
    assert!(history.record_commit(&commit).unwrap());
    assert_eq!(history.stored_bytes(), budget);
    assert_eq!(history.undo().unwrap().as_bytes(), source.as_bytes());
    assert_eq!(
        history.redo().unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert!(history.record_commit(&commit).is_err());

    let mut too_small =
        History::with_byte_budget(source.clone(), 4, source.as_bytes().len()).unwrap();
    assert!(too_small.record_commit(&commit).is_err());
    assert_eq!(too_small.current().as_bytes(), source.as_bytes());
}

#[test]
fn hostile_policy_and_lossless_boundaries_refuse_before_publication() {
    let source = flat();
    let mut external_edit = source.transaction();
    external_edit
        .set_source(0, Source::Linked("javascript:alert(1)".into()))
        .unwrap();
    let external_commit = external_edit.commit().unwrap();
    assert!(
        external_commit
            .semantic_patch(&SecurityPolicy::strict())
            .is_err()
    );
    assert!(
        external_commit
            .semantic_patch(&SecurityPolicy::default().with_operations(0))
            .is_err()
    );

    let mut map_edit = source.transaction();
    map_edit.set_image_map(0, Some(map())).unwrap();
    let map_commit = map_edit.commit().unwrap();
    assert!(
        map_commit
            .semantic_patch(&SecurityPolicy::strict())
            .is_err()
    );
    assert!(
        map_commit
            .semantic_patch(&SecurityPolicy::default().with_map_areas(1))
            .is_err()
    );
    assert!(
        map_commit
            .semantic_patch(&SecurityPolicy::default().with_patch_bytes(1))
            .is_err()
    );

    let package = package();
    let mut unsafe_member = package.edit();
    assert!(
        unsafe_member
            .put_member("../escape".into(), "image/png".into(), vec![1])
            .is_err()
    );
    assert!(
        unsafe_member
            .remove_member("META-INF/documentsignatures.xml".into())
            .is_err()
    );

    let eventful_xml = FLAT.replace(
        "</draw:frame>",
        r#"<draw:image-map><draw:area-rectangle svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"><office:event-listeners/></draw:area-rectangle></draw:image-map></draw:frame>"#,
    );
    let eventful = FlatImage::from_bytes(eventful_xml.into_bytes()).unwrap();
    assert!(
        eventful
            .transaction()
            .set_image_map(0, Some(map()))
            .is_err()
    );
}
