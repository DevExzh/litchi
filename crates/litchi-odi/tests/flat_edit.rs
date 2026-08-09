#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error;
use litchi_odi::{FlatImage, source::Source};

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.image"><office:body><office:image><draw:frame draw:name="old" draw:style-name="keep"><draw:image xlink:href="Pictures/old.png"/></draw:frame></office:image><foreign:keep xmlns:foreign="urn:foreign"> exact </foreign:keep></office:body></office:document>"#;

#[test]
fn transaction_is_exact_for_noop_and_reversible_for_owned_metadata() {
    let source = FlatImage::from_bytes(XML.as_bytes().to_vec()).unwrap();
    let unchanged = source.transaction().commit().unwrap();
    assert_eq!(unchanged.snapshot().as_bytes(), XML.as_bytes());
    assert!(unchanged.patch().changes().is_empty());

    let mut transaction = source.transaction();
    transaction.set_name(Some("new & exact".into())).unwrap();
    transaction
        .set_image_source(Source::Linked("Pictures/new&x.png".into()))
        .unwrap();
    let commit = transaction.commit().unwrap();
    let target = commit.snapshot();
    assert_eq!(target.frames()[0].name(), Some("new & exact"));
    assert_eq!(
        target.frames()[0].source(),
        &Source::Linked("Pictures/new&x.png".into())
    );
    assert!(
        std::str::from_utf8(target.as_bytes())
            .unwrap()
            .contains("<foreign:keep xmlns:foreign=\"urn:foreign\"> exact </foreign:keep>")
    );
    assert_eq!(
        commit.patch().inverse().apply(target).unwrap().as_bytes(),
        XML.as_bytes()
    );
}

#[test]
fn transaction_replaces_inline_image_bytes() {
    let xml = XML.replace(
        r#"<draw:image xlink:href="Pictures/old.png"/>"#,
        "<draw:image><office:binary-data>aQ==</office:binary-data></draw:image>",
    );
    let source = FlatImage::from_bytes(xml.into_bytes()).unwrap();
    let mut transaction = source.transaction();
    transaction
        .set_source(0, Source::Embedded(b"updated".to_vec()))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.snapshot().frames()[0].source(),
        &Source::Embedded(b"updated".to_vec())
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn transaction_refuses_stale_or_lossy_source_representations() {
    let source = FlatImage::from_bytes(XML.as_bytes().to_vec()).unwrap();
    let other = FlatImage::from_bytes(XML.replacen("old", "different", 1).into_bytes()).unwrap();
    let mut transaction = source.transaction();
    transaction
        .set_source(0, Source::Linked("Pictures/new.png".into()))
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(matches!(
        commit.patch().apply(&other),
        Err(Error::InvalidFormat(_))
    ));
    let mut lossy = source.transaction();
    assert!(lossy.set_source(0, Source::Embedded(vec![1])).is_err());
}
