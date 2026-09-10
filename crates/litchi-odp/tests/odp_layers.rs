#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::core::{OwnedPackage, PackageWriter};
use litchi_odp::{edit, layer};

const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SVG: &str = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

fn package(content: &str, styles: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn source_bytes(include_local_set: bool) -> Vec<u8> {
    let local = if include_local_set {
        r#"<d:layer-set><!-- retained layer comment --><d:layer d:name="base" draw:display="screen" d:protected="false"/></d:layer-set>"#
    } else {
        ""
    };
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:draw="{DRAW}" xmlns:style="{STYLE}" xmlns:vendor="urn:example:vendor"><office:body><office:presentation><d:page d:name="one"><vendor:opaque vendor:token="keep"/>{local}<d:frame d:name="shape" d:layer="base" vendor:untouched="yes"/></d:page><d:page d:name="two"/></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:style="{STYLE}"><office:master-styles><d:layer-set><d:layer d:name="global"/></d:layer-set></office:master-styles><style:master-page style:name="Master"/></office:document-styles>"#
    );
    package(&content, &styles)
}

#[test]
fn inventory_reads_alternate_prefixes_and_owners() {
    let snapshot = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let inventory = snapshot.layers().unwrap();
    assert_eq!(inventory.page(0).unwrap().layers()[0].name(), "base");
    assert_eq!(
        inventory.master_styles().unwrap().layers()[0].name(),
        "global"
    );
    assert!(inventory.master_page("Master").is_none());
    assert_eq!(
        inventory.page(0).unwrap().layers()[0].display(),
        Some("screen")
    );
    assert_eq!(
        inventory.page(0).unwrap().layers()[0].protected(),
        Some(false)
    );
}

#[test]
fn rename_updates_declaration_and_shape_reference_with_exact_patch() {
    let source = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let source_bytes = source.bytes().to_vec();
    let mut transaction = source.transaction().unwrap();
    transaction.rename_page_layer(0, "base", "renamed").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains(r#"d:name="renamed""#));
    assert!(content.contains(r#"d:layer="renamed""#));
    assert!(content.contains(r#"vendor:token="keep""#));
    assert!(content.contains("retained layer comment"));
    assert_eq!(
        commit.patch().apply(&source).unwrap().bytes(),
        commit.snapshot().bytes()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .bytes(),
        source_bytes
    );
    assert_eq!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .page(0)
            .unwrap()
            .layers()[0]
            .name(),
        "renamed"
    );
}

#[test]
fn remove_referenced_layer_is_refused_and_add_creates_missing_set() {
    let source = edit::Snapshot::from_bytes(source_bytes(false)).unwrap();
    let mut refused = source.transaction().unwrap();
    assert!(refused.remove_page_layer(0, "base").is_err());
    let noop = refused.commit().unwrap();
    assert!(!noop.changed());
    assert_eq!(noop.snapshot().bytes(), source.bytes());

    let mut transaction = source.transaction().unwrap();
    transaction
        .add_page_layer(0, layer::Layer::new("base").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .page(0)
            .unwrap()
            .layers()[0]
            .name(),
        "base"
    );
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("<d:layer-set"));
    assert!(content.contains(r#"d:name="base""#));
}

#[test]
fn stale_patch_is_rejected_for_layer_commit() {
    let source = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.rename_global_layer("global", "all").unwrap();
    let commit = transaction.commit().unwrap();
    let other = edit::Snapshot::from_bytes(source_bytes(false)).unwrap();
    assert!(commit.patch().apply(&other).is_err());
}

#[test]
fn reversing_layer_renames_retains_the_exact_source_package() {
    let source = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .rename_page_layer(0, "base", "temporary")
        .unwrap();
    transaction
        .rename_global_layer("global", "temporary-global")
        .unwrap();
    transaction
        .rename_page_layer(0, "temporary", "base")
        .unwrap();
    transaction
        .rename_global_layer("temporary-global", "global")
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(
        commit.patch().apply(&source).unwrap().bytes(),
        source.bytes()
    );
}

#[test]
fn global_and_master_page_additions_use_existing_style_owners() {
    let source = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_global_layer(layer::Layer::new("new-global").unwrap())
        .unwrap();
    transaction
        .add_master_page_layer("Master", layer::Layer::new("master-local").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    let inventory = commit.snapshot().layers().unwrap();
    assert_eq!(
        inventory
            .master_styles()
            .unwrap()
            .layers()
            .last()
            .unwrap()
            .name(),
        "new-global"
    );
    assert_eq!(
        inventory
            .master_page("Master")
            .unwrap()
            .layers()
            .last()
            .unwrap()
            .name(),
        "master-local"
    );
}

#[test]
fn master_page_owner_requires_style_name_ncname_and_ignores_draw_name() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"/></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:style="{STYLE}"><office:master-styles><style:master-page draw:name="Wrong" style:name="Right"/></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_master_page_layer("Right", layer::Layer::new("master-local").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .master_page("Right")
            .unwrap()
            .layers()[0]
            .name(),
        "master-local"
    );
    assert!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .master_page("Wrong")
            .is_none()
    );

    let missing_style_name = styles.replace(r#" style:name="Right""#, "");
    let missing = edit::Snapshot::from_bytes(package(&content, &missing_style_name)).unwrap();
    assert!(missing.layers().is_err());

    let invalid_style_name = styles.replace(r#"style:name="Right""#, r#"style:name="not valid""#);
    let invalid = edit::Snapshot::from_bytes(package(&content, &invalid_style_name)).unwrap();
    assert!(invalid.layers().is_err());

    let unicode_valid = styles.replace(r#"style:name="Right""#, r#"style:name="À.layer_1""#);
    let valid = edit::Snapshot::from_bytes(package(&content, &unicode_valid)).unwrap();
    assert!(valid.layers().is_ok());
    for invalid_name in ["µ", "ª", "º"] {
        let replacement = format!(r#"style:name="{invalid_name}""#);
        let invalid_style_name = styles.replace(r#"style:name="Right""#, &replacement);
        let invalid = edit::Snapshot::from_bytes(package(&content, &invalid_style_name)).unwrap();
        assert!(
            invalid.layers().is_err(),
            "accepted invalid NCName {invalid_name:?}"
        );
    }
}

#[test]
fn missing_empty_owner_elements_are_expanded_losslessly() {
    let source = edit::Snapshot::from_bytes(source_bytes(false)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_page_layer(1, layer::Layer::new("second").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .page(1)
            .unwrap()
            .layers()[0]
            .name(),
        "second"
    );
    let package = OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains(r#"<d:page d:name="two"><d:layer-set"#));
    assert!(content.contains(r#"</d:page>"#));
}

#[test]
fn display_values_follow_the_odf_enumeration() {
    assert!(
        layer::Layer::new("visible")
            .unwrap()
            .with_display("always")
            .is_ok()
    );
    assert!(
        layer::Layer::new("visible")
            .unwrap()
            .with_display("invalid-display")
            .is_err()
    );

    let invalid = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"><d:layer-set><d:layer d:name="base" d:display="invalid-display"/></d:layer-set></d:page></office:presentation></office:body></office:document-content>"#
    );
    let snapshot = edit::Snapshot::from_bytes(package(
        &invalid,
        &format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#),
    ))
    .unwrap();
    assert!(snapshot.layers().is_err());

    let invalid_protected =
        invalid.replace(r#"d:display="invalid-display""#, r#"d:protected="maybe""#);
    let snapshot = edit::Snapshot::from_bytes(package(
        &invalid_protected,
        &format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#),
    ))
    .unwrap();
    assert!(snapshot.layers().is_err());
}

#[test]
fn missing_page_layer_set_is_inserted_after_title_and_description() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:svg="{SVG}"><office:body><office:presentation><d:page d:name="one"><svg:title>title</svg:title><svg:desc>description</svg:desc><d:frame d:name="shape"/></d:page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"/></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_page_layer(0, layer::Layer::new("page-layer").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    let output = String::from_utf8(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("content.xml")
            .unwrap(),
    )
    .unwrap();
    let title = output.find("<svg:title>").unwrap();
    let desc = output.find("<svg:desc>").unwrap();
    let layers = output.find("<d:layer-set").unwrap();
    let shape = output.find("<d:frame").unwrap();
    assert!(title < desc && desc < layers && layers < shape);
}

#[test]
fn missing_master_page_layer_set_is_inserted_after_header_and_footer() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"/></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"><style:header/><style:footer/><d:frame d:name="shape"/></style:master-page></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_master_page_layer("Master", layer::Layer::new("master-layer").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    let output = String::from_utf8(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("styles.xml")
            .unwrap(),
    )
    .unwrap();
    let header = output.find("<style:header").unwrap();
    let footer = output.find("<style:footer").unwrap();
    let layers = output.find("<d:layer-set").unwrap();
    let shape = output.find("<d:frame").unwrap();
    assert!(header < footer && footer < layers && layers < shape);
}

#[test]
fn default_draw_namespace_addition_qualifies_attributes() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:presentation><page xmlns="{DRAW}" name="one"><layer-set/></page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"/></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    assert_eq!(source.layers().unwrap().page(0).unwrap().layers().len(), 0);
    let mut transaction = source.transaction().unwrap();
    transaction
        .add_page_layer(0, layer::Layer::new("default-layer").unwrap())
        .unwrap();
    let commit = transaction.commit().unwrap();
    let output = String::from_utf8(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("content.xml")
            .unwrap(),
    )
    .unwrap();
    assert!(output.contains("<draw:layer"));
    assert!(output.contains(r#"draw:name="default-layer""#));
    assert_eq!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .page(0)
            .unwrap()
            .layers()[0]
            .name(),
        "default-layer"
    );
}

#[test]
fn rename_preserves_opaque_layer_attributes_and_children() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:svg="{SVG}" xmlns:vendor="urn:example:vendor"><office:body><office:presentation><d:page d:name="one"><d:layer-set><d:layer d:name="base" vendor:flag="keep"><svg:title>opaque title</svg:title><vendor:extension vendor:value="untouched"/><opaque extension="keep"><nested/></opaque></d:layer></d:layer-set><d:frame d:name="shape" d:layer="base"/></d:page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"/></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    let mut transaction = source.transaction().unwrap();
    transaction.rename_page_layer(0, "base", "renamed").unwrap();
    let commit = transaction.commit().unwrap();
    let output = String::from_utf8(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("content.xml")
            .unwrap(),
    )
    .unwrap();
    assert!(output.contains(r#"d:name="renamed" vendor:flag="keep""#));
    assert!(output.contains("<svg:title>opaque title</svg:title>"));
    assert!(output.contains(r#"vendor:value="untouched""#));
    assert!(output.contains(r#"<opaque extension="keep"><nested/></opaque>"#));
}

#[test]
fn unresolved_prefixed_names_are_rejected() {
    let missing_prefix = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"><missing:extension/></d:page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#);
    let source = edit::Snapshot::from_bytes(package(&missing_prefix, &styles)).unwrap();
    assert!(source.layers().is_err());
}

#[test]
fn existing_layer_set_order_is_checked_before_layer_edits() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"><d:frame d:name="shape"/><d:layer-set/></d:page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(r#"<office:document-styles xmlns:office="{OFFICE}"/>"#);
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    assert!(source.layers().is_err());
}

#[test]
fn layer_set_cannot_be_followed_by_schema_prefix_children() {
    let page_content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:svg="{SVG}"><office:body><office:presentation><d:page d:name="one"><d:layer-set/><svg:title>late</svg:title></d:page></office:presentation></office:body></office:document-content>"#
    );
    let valid_styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"/></office:master-styles></office:document-styles>"#
    );
    let page_source = edit::Snapshot::from_bytes(package(&page_content, &valid_styles)).unwrap();
    assert!(page_source.layers().is_err());

    let master_content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"/></office:presentation></office:body></office:document-content>"#
    );
    let master_styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:style="{STYLE}"><office:master-styles><style:master-page style:name="Master"><d:layer-set/><style:header/></style:master-page></office:master-styles></office:document-styles>"#
    );
    let master_source =
        edit::Snapshot::from_bytes(package(&master_content, &master_styles)).unwrap();
    assert!(master_source.layers().is_err());
}

#[test]
fn global_layer_removal_closes_page_references_without_local_set() {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:d="{DRAW}"><office:body><office:presentation><d:page d:name="one"><d:frame d:name="shape" d:layer="global"/></d:page></office:presentation></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<office:document-styles xmlns:office="{OFFICE}" xmlns:d="{DRAW}" xmlns:style="{STYLE}"><office:master-styles><d:layer-set><d:layer d:name="global"/><d:layer d:name="replacement"/></d:layer-set><style:master-page style:name="Master"/></office:master-styles></office:document-styles>"#
    );
    let source = edit::Snapshot::from_bytes(package(&content, &styles)).unwrap();
    let mut refused = source.transaction().unwrap();
    assert!(refused.remove_global_layer("global").is_err());
    assert_eq!(refused.commit().unwrap().snapshot().bytes(), source.bytes());

    let mut transaction = source.transaction().unwrap();
    transaction
        .remove_layer_with_replacement(
            layer::LayerOwner::master_styles(),
            "global",
            Some("replacement"),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    let output = String::from_utf8(
        OwnedPackage::from_bytes(commit.snapshot().bytes().to_vec())
            .unwrap()
            .get_file("content.xml")
            .unwrap(),
    )
    .unwrap();
    assert!(output.contains(r#"d:layer="replacement""#));
    assert!(!output.contains(r#"d:layer="global""#));
    assert!(
        commit
            .snapshot()
            .layers()
            .unwrap()
            .master_styles()
            .unwrap()
            .get("global")
            .unwrap()
            .is_none()
    );
}

#[test]
fn same_name_rename_is_an_exact_noop() {
    let source = edit::Snapshot::from_bytes(source_bytes(true)).unwrap();
    let original = source.bytes().to_vec();
    let mut transaction = source.transaction().unwrap();
    transaction.rename_page_layer(0, "base", "base").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().bytes(), original);
}
