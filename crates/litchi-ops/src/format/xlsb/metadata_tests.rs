//! Source-backed metadata sets keep the handle source-backed (record 0261).

use std::path::PathBuf;

use crate::{Code, File, props};

fn fixture() -> Vec<u8> {
    std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/xlsb/Simple.xlsb"),
    )
    .expect("XLSB fixture")
}

fn xlsb(file: &File) -> &super::Xlsb {
    match file {
        File::Xlsb(xlsb) => xlsb,
        other => panic!("not XLSB: {other:?}"),
    }
}

fn title(file: &File) -> Option<String> {
    file.metadata()
        .expect("metadata")
        .and_then(|value| value.title)
}

#[test]
fn a_source_backed_set_publishes_over_the_source_without_materializing() {
    let bytes = fixture();
    let file = crate::open_bytes(bytes.clone()).expect("open");
    assert!(xlsb(&file).is_source_backed());
    let text = xlsb(&file).text().expect("text");

    let edited = file
        .set(
            "/metadata",
            props! { "title" => "Quarterly", "extended.company" => "Litchi" },
        )
        .expect("metadata set");
    assert!(
        xlsb(&edited).is_source_backed(),
        "the handle stays source-backed"
    );
    assert_eq!(title(&edited).as_deref(), Some("Quarterly"));
    assert_eq!(xlsb(&edited).text().expect("text"), text);
    assert_eq!(file.to_bytes().expect("source bytes"), bytes);

    let reopened = crate::open_bytes(edited.to_bytes().expect("bytes")).expect("reopen");
    assert_eq!(title(&reopened).as_deref(), Some("Quarterly"));
    assert_eq!(
        reopened
            .metadata()
            .expect("metadata")
            .and_then(|value| value.company)
            .as_deref(),
        Some("Litchi")
    );

    // A later workbook edit materializes the revised parts, not the source.
    let celled = edited
        .set("/Sheet1/A1", props! { "value" => "after" })
        .expect("cell set");
    let reopened = crate::open_bytes(celled.to_bytes().expect("bytes")).expect("reopen");
    assert_eq!(title(&reopened).as_deref(), Some("Quarterly"));
    assert_eq!(
        reopened
            .get("/Sheet1/A1")
            .expect("cell")
            .view()
            .text()
            .expect("text")
            .as_deref(),
        Some("after")
    );

    // A second metadata set composes over the source.
    let twice = edited
        .set("/metadata", props! { "subject" => "Plans" })
        .expect("second set");
    assert!(xlsb(&twice).is_source_backed());
    let reopened = crate::open_bytes(twice.to_bytes().expect("bytes")).expect("reopen");
    let value = reopened.metadata().expect("metadata").unwrap_or_default();
    assert_eq!(value.title.as_deref(), Some("Quarterly"));
    assert_eq!(value.subject.as_deref(), Some("Plans"));
}

#[test]
fn a_source_backed_no_op_returns_the_same_snapshot() {
    let file = crate::open_bytes(fixture()).expect("open");
    let current = title(&file);
    let same = file
        .set(
            "/metadata",
            props! { "title" => current.clone().unwrap_or_default() },
        )
        .expect("no-op set");
    if current.is_some() {
        assert!(xlsb(&same).same_snapshot(xlsb(&file)));
        assert_eq!(
            same.to_bytes().expect("bytes"),
            file.to_bytes().expect("bytes")
        );
    }
}

#[test]
fn source_backed_metadata_sets_replay_and_invert_in_transactions() {
    let file = crate::open_bytes(fixture()).expect("open");
    let commit = file
        .edit(|tx| {
            tx.set("/metadata", props! { "title" => "In transaction" })?;
            tx.set("/Sheet1/B2", props! { "value" => 7 })?;
            Ok(())
        })
        .expect("transaction");
    assert!(commit.diagnostics().changed());
    assert_eq!(title(commit.file()).as_deref(), Some("In transaction"));
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.file())
        .expect("inverse");
    assert_eq!(
        restored.to_bytes().expect("restored"),
        file.to_bytes().expect("source")
    );
    let replayed = commit.patch().apply(&file).expect("replay");
    assert_eq!(
        replayed.to_bytes().expect("replayed"),
        commit.file().to_bytes().expect("committed")
    );
}

#[test]
fn a_source_backed_set_rejects_unstorable_fields_and_changes_nothing() {
    let file = crate::open_bytes(fixture()).expect("open");
    let error = file
        .set("/metadata", props! { "contentType" => "text/plain" })
        .expect_err("unstorable");
    assert_eq!(error.code(), Code::UnsupportedProperty);
    assert_eq!(error.property(), Some("content_type"));
    assert!(xlsb(&file).is_source_backed());
}

#[test]
fn eager_and_source_owners_store_the_same_property_parts() {
    let bytes = fixture();
    let source = crate::open_bytes(bytes.clone()).expect("source open");
    let eager = crate::open_bytes_with(
        bytes,
        &crate::OpenOptions::new().source_mode(crate::SourceMode::Eager),
    )
    .expect("eager open");
    assert!(!xlsb(&eager).is_source_backed());
    let props = props! { "title" => "Same", "extended.manager" => "Grace" };
    let parts = |file: &File| {
        let bytes = file
            .set("/metadata", props.clone())
            .expect("set")
            .to_bytes()
            .expect("bytes");
        let package =
            litchi_opc::OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default())
                .expect("package");
        let mut parts: Vec<_> = package
            .iter_parts()
            .map(|part| part.partname().clone())
            .filter(|name| name.as_str().starts_with("/docProps/"))
            .map(|name| {
                let blob = package.get_part(&name).expect("part").blob().to_vec();
                (name.as_str().to_owned(), blob)
            })
            .collect();
        parts.sort();
        parts
    };
    assert_eq!(parts(&source), parts(&eager));
}
