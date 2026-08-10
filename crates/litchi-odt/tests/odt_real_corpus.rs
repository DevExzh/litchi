//! Real Writer corpus compatibility through the current ODT facade.

use litchi_core::Error;
use litchi_odt::{Document, ScriptResourceKind, core::OwnedPackage, transaction::Position};
use std::path::{Path, PathBuf};

fn corpus() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/corpus")
}

fn libreoffice_macro() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/xmlsecurity/qa/unit/signing/data/macro.odt")
}

fn assert_typed_or_ok<T>(result: Result<T, Error>, file: &str, api: &str) {
    if let Err(error) = result {
        assert!(
            matches!(error, Error::InvalidFormat(_)),
            "{file}: {api} produced a non-format error: {error:?}"
        );
    }
}

fn assert_invalid<T>(result: Result<T, Error>, file: &str, api: &str) {
    match result {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("{file}: {api} produced a non-format error: {error:?}"),
        Ok(_) => panic!("{file}: {api} unexpectedly parsed hostile XML"),
    }
}

#[test]
fn writer_packages_parse_and_extract_semantics() -> Result<(), Error> {
    let cases = [
        ("writer-table.odt", 5, 1),
        ("writer-header-footer.odt", 1, 0),
        ("writer-user-fields.odt", 4, 0),
        ("writer-images-frames.odt", 7, 0),
        ("writer-definition-lists.odt", 14, 0),
        ("writer-paragraph-styles.odt", 10, 0),
        ("writer-table-of-contents.odt", 9, 0),
    ];
    for (file, minimum_paragraphs, minimum_tables) in cases {
        let document = Document::open(corpus().join(file))?;
        assert!(!document.text()?.trim().is_empty(), "{file}: empty text");
        assert!(
            document.paragraphs()?.len() >= minimum_paragraphs,
            "{file}: too few paragraphs"
        );
        assert!(
            document.tables()?.len() >= minimum_tables,
            "{file}: too few tables"
        );
        assert_typed_or_ok(document.sections(), file, "sections");
        assert_typed_or_ok(document.forms(), file, "forms");
        assert_typed_or_ok(document.tracked_changes(), file, "tracked_changes");
        assert_typed_or_ok(document.dynamic_text_fields(), file, "dynamic_text_fields");
        assert_typed_or_ok(document.text_indexes(), file, "text_indexes");
        assert_typed_or_ok(document.master_pages(), file, "master_pages");
        assert_typed_or_ok(document.page_sequence(), file, "page_sequence");
        assert_typed_or_ok(document.metadata(), file, "metadata");
    }
    Ok(())
}

#[test]
fn writer_packages_expose_fixture_specific_semantics() -> Result<(), Error> {
    let table = Document::open(corpus().join("writer-table.odt"))?;
    assert_eq!(table.tables()?.len(), 1);

    let headers = Document::open(corpus().join("writer-header-footer.odt"))?;
    assert!(!headers.master_pages()?.is_empty());

    let fields = Document::open(corpus().join("writer-user-fields.odt"))?;
    assert!(!fields.dynamic_text_fields()?.is_empty());

    let images = Document::open(corpus().join("writer-images-frames.odt"))?;
    assert!(!images.images()?.is_empty());

    let toc = Document::open(corpus().join("writer-table-of-contents.odt"))?;
    assert_eq!(toc.text_indexes()?.len(), 1);
    Ok(())
}

#[test]
fn genuine_formatted_metadata_survives_a_public_paragraph_transaction_exactly() -> Result<(), Error>
{
    let path = corpus().join("writer-header-footer.odt");
    let source_bytes = std::fs::read(&path)?;
    let source_package = OwnedPackage::from_bytes(source_bytes.clone())?;
    let source_meta = source_package.get_file("meta.xml")?;
    let document = Document::from_bytes(source_bytes)?;
    let mut edit = document.edit()?;
    edit.replace_paragraph(Position::new(0), "provenance-safe paragraph")?;
    let commit = edit.commit()?;

    let reopened = commit.snapshot().document()?;
    assert!(reopened.text()?.contains("provenance-safe paragraph"));
    let changed = OwnedPackage::from_bytes(commit.snapshot().as_bytes().to_vec())?;
    assert_eq!(changed.get_file("meta.xml")?, source_meta);
    let manifest = String::from_utf8(changed.get_file("META-INF/manifest.xml")?)
        .map_err(|error| Error::InvalidFormat(error.to_string()))?;
    assert!(!manifest.contains("manifest:full-path=\"META-INF/\""));
    assert!(!manifest.contains("manifest:full-path=\"META-INF/manifest.xml\""));
    Ok(())
}

#[test]
fn real_xxe_package_is_rejected_without_resolution() -> Result<(), Error> {
    let file = "writer-minimal-nasty.odt";
    let document = Document::open(corpus().join(file))?;
    assert_invalid(document.text(), file, "text");
    assert_invalid(document.paragraphs(), file, "paragraphs");
    assert_invalid(document.forms(), file, "forms");
    drop(document.metadata()?);
    Ok(())
}

#[test]
fn real_macro_package_exposes_only_inert_stored_resources() -> Result<(), Error> {
    let document = Document::open(libreoffice_macro())?;
    let resources = document.script_resources()?;
    assert_eq!(resources.len(), 3);
    assert!(resources.iter().all(|resource| !resource.bytes.is_empty()));
    assert!(resources.iter().any(|resource| {
        resource.path == "Basic/Standard/Module1.xml"
            && resource.kind == ScriptResourceKind::BasicModule
    }));
    assert!(resources.iter().any(|resource| {
        resource.path == "Basic/Standard/script-lb.xml"
            && resource.kind == ScriptResourceKind::BasicLibrary
    }));
    assert!(resources.iter().any(|resource| {
        resource.path == "Basic/script-lc.xml" && resource.kind == ScriptResourceKind::BasicLibrary
    }));
    Ok(())
}
