//! LENIENCY_AUDIT F9.11 at the ODP facade: a document type declaration in
//! the `content.xml` prologue is ignored without being processed, as
//! LibreOffice's SAX import ignores it, through every ODP open path, and
//! the repair is recorded on the document handle. Entity references that a
//! DTD would have declared stay hard errors, so entity safety is unchanged.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "../../litchi-odf-common/tests/support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{ContentDefect, Defect, ToleranceReport};
use litchi_odp::{Presentation, SourceBackedPresentation};

const OOO_DOCTYPE: &str = r#"<!DOCTYPE office:document-content PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "office.dtd">"#;

fn recorded(report: &ToleranceReport) -> Vec<(Defect, u16)> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn conforming() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/corpus/impress-master-layouts.odp"),
    )
    .unwrap()
}

/// Insert `doctype` right after the XML declaration of `content.xml`.
fn with_doctype(doctype: &str) -> Vec<u8> {
    odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        let end = xml.find("?>").expect("content.xml has an XML declaration") + 2;
        format!("{}\n{doctype}{}", &xml[..end], &xml[end..])
    })
}

fn assert_every_open_path(bytes: &[u8], defects: &[(Defect, u16)], label: &str) {
    let owned = Presentation::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(owned.slides().unwrap().len(), 5, "{label}: slides");
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Presentation::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(source.slides().unwrap().len(), 5, "{label}: source slides");
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
}

const DOCTYPE: (Defect, u16) = (Defect::Content(ContentDefect::DoctypeDeclaration), 0);

#[test]
fn a_prologue_doctype_is_ignored_and_recorded() {
    assert_every_open_path(&with_doctype(OOO_DOCTYPE), &[DOCTYPE], "public doctype");
    // An internal subset that declares nothing used is ignored the same way;
    // its `>` inside a quoted literal does not end the declaration early.
    assert_every_open_path(
        &with_doctype(r#"<!DOCTYPE office:document-content [ <!ENTITY unused "a > b"> ]>"#),
        &[DOCTYPE],
        "internal subset",
    );
}

#[test]
fn the_same_document_without_a_doctype_repairs_nothing() {
    assert_every_open_path(&conforming(), &[], "original");
    let rezipped = odf_rezip::rezip(&conforming(), odf_rezip::MimetypeLayout::CONFORMING);
    assert_every_open_path(&rezipped, &[], "re-zipped");
}

#[test]
fn a_dtd_never_declares_entities() {
    // The internal subset is not processed: `&payload;` stays undeclared, so
    // its replacement text can never appear in the projection.
    let bytes = odf_rezip::edit_member(&conforming(), "content.xml", |xml| {
        let end = xml.find("?>").unwrap() + 2;
        let declared = format!(
            "{}<!DOCTYPE office:document-content [ <!ENTITY payload \"EXPANDED-PAYLOAD\"> ]>{}",
            &xml[..end],
            &xml[end..]
        );
        let paragraph = declared.find("<text:p").expect("deck has a paragraph");
        let open_end = paragraph + declared[paragraph..].find('>').unwrap() + 1;
        format!(
            "{}&payload;{}",
            &declared[..open_end],
            &declared[open_end..]
        )
    });
    let never_expanded = |slides: litchi_core::Result<Vec<litchi_odp::model::Slide>>| match slides {
        Err(_) => true,
        Ok(slides) => slides
            .iter()
            .all(|slide| !format!("{slide:?}").contains("EXPANDED-PAYLOAD")),
    };
    assert!(never_expanded(
        Presentation::from_bytes(bytes.clone()).and_then(|presentation| presentation.slides())
    ));
    assert!(never_expanded(
        SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes)))
            .and_then(|presentation| presentation.slides())
    ));
}
