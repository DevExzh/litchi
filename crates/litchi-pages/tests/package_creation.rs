//! Fresh documents, document properties and package directories (record
//! 0299): `Package::new`, `builder::{Builder, Identity}`,
//! `Package::edit_metadata`, and the directory route of `Package::open`.

use std::path::PathBuf;

use litchi_pages::builder::{Builder, Identity};
use litchi_pages::metadata::Error as MetadataError;
use litchi_pages::{Package, PackageError};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(relative)
}

#[test]
fn new_and_built_documents_open_edit_and_save() -> TestResult {
    let blank = Package::new()?;
    assert_eq!(blank.sections().len(), 1);
    assert_eq!(blank.text()?, "");
    let reopened = Package::from_bytes(&blank.to_bytes()?)?;
    assert_eq!(reopened.to_bytes()?, blank.to_bytes()?);

    let built = Builder::new()
        .identity(Identity::from_bytes([42; 16]))
        .body_text("First\nSecond")
        .build()?;
    assert_eq!(built.section_paragraphs(0_usize)?, ["First", "Second"]);
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("built.pages");
    built.save(&path)?;
    let saved = Package::open(&path)?;
    assert_eq!(saved.text()?, "First\nSecond");
    assert_eq!(
        saved.metadata().identifier.as_deref(),
        Some(Identity::from_bytes([42; 16]).document_uuid().as_str())
    );

    let generated = Builder::new().identity(Identity::generate()).build()?;
    assert_ne!(generated.metadata().identifier, saved.metadata().identifier);
    assert!(matches!(
        Builder::new().build(),
        Err(PackageError::MissingIdentity)
    ));
    Ok(())
}

#[test]
fn metadata_edits_round_trip_and_invert() -> TestResult {
    let package = Package::open(fixture("pages/basic.pages"))?;
    let original = package.to_bytes()?;
    let mut edit = package.edit_metadata();
    edit.set_title(Some("Quarterly report"))?
        .set_author(Some("Ryker"))?
        .set_keywords(&["finance".to_owned(), "q3".to_owned()])?
        .set_description(Some("Draft"))?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let metadata = commit.package().metadata();
    assert_eq!(metadata.title.as_deref(), Some("Quarterly report"));
    assert_eq!(metadata.author.as_deref(), Some("Ryker"));
    assert_eq!(metadata.keywords, ["finance", "q3"]);
    assert_eq!(metadata.description.as_deref(), Some("Draft"));
    // Only the properties member changed: the text still reads back.
    assert_eq!(commit.package().text()?, package.text()?);
    let reopened = Package::from_bytes(&commit.package().to_bytes()?)?;
    assert_eq!(reopened.metadata(), metadata);
    // A keyword the one `Keywords` string cannot keep is refused.
    let mut refused = package.edit_metadata();
    assert!(matches!(
        refused.set_keywords(&["a, b".to_owned()]),
        Err(MetadataError::Refused(
            litchi_pages::metadata::Refusal::InvalidKeyword
        ))
    ));

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_metadata(&inverse)?;
    assert!(!restored.patch().is_noop());
    assert!(restored.diagnostics().changed());
    assert_eq!(restored.patch(), &inverse);
    assert_eq!(restored.package().to_bytes()?, original);
    assert!(matches!(
        package.apply_metadata(&commit.patch().inverse()),
        Err(MetadataError::PatchConflict)
    ));

    let mut clear = commit.package().edit_metadata();
    clear.set_title(None)?;
    let cleared = clear.commit()?;
    assert_eq!(cleared.package().metadata().title, None);
    assert_eq!(
        cleared.package().metadata().author.as_deref(),
        Some("Ryker")
    );

    let mut noop = package.edit_metadata();
    noop.set_title(package.metadata().title.as_deref())?;
    let unchanged = noop.commit()?;
    assert!(!unchanged.diagnostics().changed());
    assert!(unchanged.patch().is_noop());
    assert_eq!(unchanged.package().to_bytes()?, original);

    let replayed_noop = package.apply_metadata(unchanged.patch())?;
    assert!(!replayed_noop.diagnostics().changed());
    assert!(replayed_noop.patch().is_noop());
    assert_eq!(replayed_noop.package().to_bytes()?, original);

    let mut invalid = package.edit_metadata();
    assert!(matches!(
        invalid.set_title(Some("a\0b")),
        Err(MetadataError::Refused(
            litchi_pages::metadata::Refusal::InvalidValue
        ))
    ));
    Ok(())
}

#[test]
fn package_directories_open_as_single_file_packages() -> TestResult {
    let directory = Package::open(fixture("directory/pages/basic.pages"))?;
    let single = Package::open(fixture("pages/basic.pages"))?;
    assert_eq!(directory.text()?, single.text()?);
    let temp = tempfile::tempdir()?;
    let path = temp.path().join("converted.pages");
    directory.save(&path)?;
    assert!(path.is_file());
    assert_eq!(Package::open(&path)?.text()?, single.text()?);
    Ok(())
}

fn package_bytes(package: &Package) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn packages_save_back_as_directories_byte_exactly() -> TestResult {
    let temp = tempfile::tempdir()?;
    // A directory-backed package saves back as a directory that opens to the
    // same single-file bytes.
    let directory = Package::open(fixture("directory/pages/basic.pages"))?;
    let path = temp.path().join("bundle.pages");
    directory.save_directory(&path)?;
    assert!(path.join("Index.zip").is_file());
    assert!(!path.join("Index").exists());
    let reopened = Package::open(&path)?;
    assert_eq!(package_bytes(&reopened)?, package_bytes(&directory)?);
    // Saving over the existing directory replaces it.
    let edited = Package::new()?;
    edited.save_directory(&path)?;
    assert_eq!(Package::open(&path)?.text()?, edited.text()?);
    assert_eq!(
        std::fs::read_dir(temp.path())?.count(),
        1,
        "no staging or retired directory is left behind"
    );
    // A single-file package converts too, and a file destination is refused.
    let single = Package::open(fixture("pages/basic.pages"))?;
    let other = temp.path().join("single.pages");
    single.save_directory(&other)?;
    assert_eq!(Package::open(&other)?.text()?, single.text()?);
    let file = temp.path().join("file.pages");
    single.save(&file)?;
    assert!(matches!(
        single.save_directory(&file),
        Err(litchi_pages::SaveError::Publication(_))
    ));
    Ok(())
}
