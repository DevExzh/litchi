#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests deliberately panic on unexpected failures"
)]

use std::io::Cursor;

use litchi_odf_common::constants::ODF_FORMULA;
use litchi_odf_common::core::{OwnedPackage, PackageWriter};
use litchi_odf_formula::authoring::{self, Display};
use litchi_odf_formula::codec;
use litchi_odf_formula::{Formula, Limits, NodePath, StarMathVersion};

const MATHML: &str = r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>x</mi></math>"#;

#[test]
fn root_transaction_is_atomic_source_checked_and_reversible() {
    let source = Formula::create(MATHML).expect("source");
    let source_bytes = source.to_bytes();
    let replacement = authoring::document_root(authoring::number("42"), Display::Block);

    let mut edit = source.edit();
    edit.set_root(replacement.clone()).expect("stage root");
    let commit = edit.commit().expect("commit root");

    assert!(commit.changed());
    assert!(commit.diagnostics().candidate_reopened());
    assert_eq!(
        commit.patch().change().expect("change").before(),
        source.root()
    );
    assert_eq!(
        commit.patch().change().expect("change").after(),
        &replacement
    );
    assert_eq!(source.as_bytes(), source_bytes);
    assert!(commit.patch().is_applicable_to(&source));

    let published = commit.patch().apply(&source).expect("apply");
    assert_eq!(published.root(), &replacement);
    let restored = commit.patch().inverse().apply(&published).expect("inverse");
    assert_eq!(restored.as_bytes(), source_bytes);

    let unrelated =
        Formula::create(r#"<math xmlns="http://www.w3.org/1998/Math/MathML"><mi>z</mi></math>"#)
            .expect("unrelated");
    assert!(!commit.patch().is_applicable_to(&unrelated));
    assert!(commit.patch().apply(&unrelated).is_err());
}

#[test]
fn semantic_noop_reuses_exact_package_and_skips_reopen() {
    let source = Formula::create(MATHML).expect("source");
    let mut edit = source.edit();
    edit.set_root(source.root().clone()).expect("stage no-op");
    let commit = edit.commit().expect("commit no-op");

    assert!(!commit.changed());
    assert!(!commit.diagnostics().candidate_reopened());
    assert!(commit.patch().change().is_none());
    assert_eq!(commit.formula().as_bytes(), source.as_bytes());
}

#[test]
fn transaction_preserves_auxiliary_members() {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODF_FORMULA).expect("mimetype");
    writer
        .add_file("content.xml", MATHML.as_bytes())
        .expect("content");
    writer
        .add_file("Configurations2/accelerator/current.xml", b"<config/>")
        .expect("auxiliary");
    let source = Formula::from_bytes(writer.finish_to_bytes().expect("package")).expect("open");

    let mut edit = source.edit();
    edit.set_root(authoring::document_root(
        authoring::identifier("changed"),
        Display::Inline,
    ))
    .expect("stage");
    let target = edit.commit().expect("commit").into_formula();
    let package = OwnedPackage::from_bytes(target.to_bytes()).expect("reopen raw package");

    assert_eq!(
        package
            .get_file("Configurations2/accelerator/current.xml")
            .expect("auxiliary member"),
        b"<config/>"
    );
}

#[test]
fn changed_libreoffice_formula_reopens_with_non_content_members_exact() {
    let bytes =
        include_bytes!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/tdf151842.odf");
    let source = Formula::from_bytes(bytes.to_vec()).expect("LibreOffice source");
    let original = OwnedPackage::from_bytes(bytes.to_vec()).expect("raw source package");
    let mut token_edit = source.edit();
    token_edit
        .set_text(&NodePath::new([0, 0]), "changed")
        .expect("stage token");
    let mut starmath_edit = source.edit();
    starmath_edit
        .set_starmath_source(StarMathVersion::V6, "changed + 1")
        .expect("stage StarMath");
    token_edit
        .join(&starmath_edit)
        .expect("join producer edits");
    let commit = token_edit.commit().expect("publish");
    assert!(commit.diagnostics().candidate_reopened());
    assert_eq!(commit.record().expect("commit record").changes().len(), 2);
    let published = commit.formula();
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(published)
            .expect("inverse")
            .as_bytes(),
        source.as_bytes()
    );
    let reopened = Formula::from_bytes(published.to_bytes()).expect("full Formula reopen");
    assert_eq!(reopened.starmath_source().as_deref(), Some("changed + 1"));

    let changed = OwnedPackage::from_bytes(reopened.to_bytes()).expect("raw changed package");
    let mut original_names = original.files().expect("source members");
    let mut changed_names = changed.files().expect("changed members");
    original_names.retain(|name| !name.ends_with('/'));
    changed_names.retain(|name| !name.ends_with('/'));
    original_names.sort_unstable();
    changed_names.sort_unstable();
    assert_eq!(changed_names, original_names);
    for name in original_names {
        if !matches!(name.as_str(), "content.xml" | "META-INF/manifest.xml") {
            assert_eq!(
                changed.get_file(&name).expect("changed member"),
                original.get_file(&name).expect("source member"),
                "member {name} changed"
            );
        }
    }
}

#[test]
fn selected_limits_apply_at_exact_boundaries_and_survive_edits() {
    let exact = Limits::new()
        .with_xml_bytes(MATHML.len())
        .expect("XML limit")
        .with_depth(2)
        .expect("depth limit")
        .with_nodes(2)
        .expect("node limit")
        .with_text_bytes(1)
        .expect("text limit");
    let source = Formula::create_with_limits(MATHML, exact).expect("exact-limit formula");
    assert_eq!(source.limits(), exact);

    assert!(
        Formula::create_with_limits(
            MATHML,
            exact
                .with_xml_bytes(MATHML.len() - 1)
                .expect("smaller XML limit")
        )
        .is_err()
    );
    assert!(Formula::create_with_limits(MATHML, exact.with_nodes(1).expect("one node")).is_err());

    let mut edit = source.edit();
    assert!(
        edit.set_root(authoring::document_root(
            authoring::literal_text("too long"),
            Display::Block
        ))
        .is_err()
    );
}

#[test]
fn stream_and_owned_package_byte_limits_are_enforced() {
    let package = Formula::create(MATHML).expect("package").to_bytes();
    let exact = Limits::new()
        .with_package_bytes(package.len())
        .expect("exact package limit");
    assert!(Formula::from_bytes_with_limits(package.clone(), exact).is_ok());
    assert!(Formula::from_reader_with_limits(Cursor::new(package.clone()), exact).is_ok());

    let smaller = exact
        .with_package_bytes(package.len() - 1)
        .expect("smaller package limit");
    assert!(Formula::from_bytes_with_limits(package.clone(), smaller).is_err());
    assert!(Formula::from_reader_with_limits(Cursor::new(package), smaller).is_err());
}

#[test]
fn serializer_assigns_foreign_prefixes_in_first_use_order() {
    let mut root = authoring::document_root(authoring::identifier("x"), Display::Inline);
    root.push_child(
        litchi_odf_formula::Element::with_namespace(Some("urn:vendor:first"), "one")
            .expect("first extension"),
    );
    root.push_child(
        litchi_odf_formula::Element::with_namespace(Some("urn:vendor:second"), "two")
            .expect("second extension"),
    );

    let expected = r#"<math xmlns="http://www.w3.org/1998/Math/MathML" xmlns:ns1="urn:vendor:first" xmlns:ns2="urn:vendor:second" display="inline"><mi>x</mi><ns1:one/><ns2:two/></math>"#;
    for _ in 0..16 {
        assert_eq!(codec::serialize(&root), expected);
    }
}

#[test]
fn checked_limit_builders_reject_zero_and_hard_ceiling_overflow() {
    assert!(Limits::new().with_depth(0).is_err());
    assert!(Limits::new().with_nodes(codec::HARD_MAX_NODES + 1).is_err());
    assert!(
        Limits::new()
            .with_package_bytes(codec::HARD_MAX_PACKAGE_BYTES)
            .is_ok()
    );
}

#[test]
fn libreoffice_formula_fixtures_open_and_remain_byte_exact() {
    let fixtures: [&[u8]; 2] = [
        include_bytes!(
            "../../../3rdparty/libreoffice-core/starmath/qa/cppunit/data/font-styles.odf"
        ),
        include_bytes!("../../../3rdparty/libreoffice-core/starmath/qa/extras/data/tdf151842.odf"),
    ];

    for bytes in fixtures {
        let formula = Formula::from_bytes(bytes.to_vec()).expect("LibreOffice formula fixture");
        assert_eq!(formula.as_bytes(), bytes);
        assert_eq!(formula.mimetype(), ODF_FORMULA);
        assert!(!formula.files().expect("files").is_empty());
    }
}
