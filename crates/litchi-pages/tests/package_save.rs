use std::fs;
use std::io::{self, Write};
use std::path::PathBuf;

use litchi_pages::{Package, SaveError};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages/basic.pages")
}

fn round_trip_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn save_creates_new_destination_and_reopens_exact_source() -> TestResult {
    let source = fixture_path();
    let expected = fs::read(&source)?;
    let package = Package::open(&source)?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("new.pages");

    package.save(&destination)?;

    assert_eq!(fs::read(&destination)?, expected);
    let reopened = Package::open(&destination)?;
    assert_eq!(round_trip_bytes(&reopened)?, expected);
    Ok(())
}

#[test]
fn save_replaces_existing_destination_and_reopens() -> TestResult {
    let source = fixture_path();
    let expected = fs::read(&source)?;
    let package = Package::open(&source)?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("replacement.pages");
    fs::write(&destination, b"old destination bytes")?;

    package.save(&destination)?;

    assert_eq!(fs::read(&destination)?, expected);
    let reopened = Package::open(&destination)?;
    assert_eq!(round_trip_bytes(&reopened)?, expected);
    Ok(())
}

#[test]
fn save_supports_same_source_path_and_reopens() -> TestResult {
    let directory = tempfile::tempdir()?;
    let source = directory.path().join("same-source.pages");
    fs::copy(fixture_path(), &source)?;
    let package = Package::open(&source)?;
    let expected = fs::read(&source)?;

    package.save(&source)?;

    assert_eq!(fs::read(&source)?, expected);
    let reopened = Package::open(&source)?;
    assert_eq!(round_trip_bytes(&reopened)?, expected);
    Ok(())
}

#[test]
fn save_rejects_directory_destination_without_touching_it() -> TestResult {
    let package = Package::open(fixture_path())?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("destination-directory");
    fs::create_dir(&destination)?;

    let error = package
        .save(&destination)
        .expect_err("a directory cannot be a package destination");

    assert!(matches!(error, SaveError::Publication(_)));
    assert!(!error.was_committed());
    assert!(destination.is_dir());
    Ok(())
}

#[cfg(unix)]
#[test]
fn save_rejects_symbolic_link_destination_without_touching_target() -> TestResult {
    use std::os::unix::fs::symlink;

    let package = Package::open(fixture_path())?;
    let directory = tempfile::tempdir()?;
    let target = directory.path().join("target.pages");
    let link = directory.path().join("link.pages");
    fs::write(&target, b"target remains unchanged")?;
    symlink(&target, &link)?;

    let error = package
        .save(&link)
        .expect_err("a symbolic-link destination must be refused");

    assert!(matches!(error, SaveError::Publication(_)));
    assert!(!error.was_committed());
    assert_eq!(fs::read(&target)?, b"target remains unchanged");
    assert!(fs::symlink_metadata(&link)?.file_type().is_symlink());
    Ok(())
}

#[test]
fn save_errors_are_redacted_and_report_the_failure_family() -> TestResult {
    const SECRET: &str = "private-pages-save-secret";

    struct FailingWriter;

    impl Write for FailingWriter {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            Err(io::Error::other(SECRET))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    let package = Package::open(fixture_path())?;
    let write_error = package
        .write_to(&mut FailingWriter)
        .expect_err("the injected writer must fail");
    let write_error = SaveError::Write(write_error);
    assert!(write_error.write_error().is_some());
    assert!(!write_error.was_committed());
    assert!(!format!("{write_error:?}").contains(SECRET));
    assert!(!write_error.to_string().contains(SECRET));

    let directory = tempfile::tempdir()?;
    let destination = directory.path().join(SECRET).join("saved.pages");
    let publication_error = package
        .save(&destination)
        .expect_err("missing parent should fail publication");
    assert!(publication_error.write_error().is_none());
    assert!(publication_error.publication_error().is_some());
    assert!(!publication_error.was_committed());
    assert!(!format!("{publication_error:?}").contains(SECRET));
    assert!(!publication_error.to_string().contains(SECRET));
    Ok(())
}

#[test]
fn save_types_are_send_sync_and_debug() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<SaveError>();
}
