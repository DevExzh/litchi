//! Durable exact-artifact save coverage for native Numbers packages.

use std::error::Error as StdError;
use std::io::{self, Write};
use std::path::PathBuf;

use litchi_numbers::{Package, SaveError};

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

#[test]
fn save_to_new_path_is_exact_and_reopens() -> TestResult {
    let source = fixture_path();
    let expected = std::fs::read(&source)?;
    let package = Package::open(&source)?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("saved.numbers");

    package.save(&destination)?;

    assert_eq!(std::fs::read(&destination)?, expected);
    let reopened = Package::open(&destination)?;
    let mut round_trip = Vec::new();
    reopened.write_to(&mut round_trip)?;
    assert_eq!(round_trip, expected);
    Ok(())
}

#[test]
fn save_replaces_existing_destination_exactly_and_reopens() -> TestResult {
    let source = fixture_path();
    let expected = std::fs::read(&source)?;
    let package = Package::open(&source)?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("saved.numbers");
    std::fs::write(&destination, b"stale destination")?;

    package.save(&destination)?;

    assert_eq!(std::fs::read(&destination)?, expected);
    let reopened = Package::open(&destination)?;
    let mut round_trip = Vec::new();
    reopened.write_to(&mut round_trip)?;
    assert_eq!(round_trip, expected);
    Ok(())
}

#[test]
fn save_to_the_same_source_path_is_exact_and_reopens() -> TestResult {
    let source = fixture_path();
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("same-source.numbers");
    std::fs::copy(&source, &destination)?;
    let expected = std::fs::read(&destination)?;
    let package = Package::open(&destination)?;

    package.save(&destination)?;

    assert_eq!(std::fs::read(&destination)?, expected);
    let reopened = Package::open(&destination)?;
    let mut round_trip = Vec::new();
    reopened.write_to(&mut round_trip)?;
    assert_eq!(round_trip, expected);
    Ok(())
}

#[test]
fn save_rejects_a_directory_without_touching_it() -> TestResult {
    let package = Package::open(fixture_path())?;
    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("existing.numbers");
    std::fs::write(&destination, b"untouched destination")?;

    let error = package
        .save(directory.path())
        .expect_err("directory destination");

    assert!(matches!(error, SaveError::Publication(_)));
    assert!(error.write_error().is_none());
    assert!(error.publication_error().is_some());
    assert!(!error.was_committed());
    assert_eq!(std::fs::read(&destination)?, b"untouched destination");
    Ok(())
}

#[cfg(unix)]
#[test]
fn save_rejects_a_symbolic_link_destination() -> TestResult {
    use std::os::unix::fs::symlink;

    let package = Package::open(fixture_path())?;
    let directory = tempfile::tempdir()?;
    let target = directory.path().join("target.numbers");
    let destination = directory.path().join("linked.numbers");
    std::fs::write(&target, b"untouched target")?;
    symlink(&target, &destination)?;

    let error = package
        .save(&destination)
        .expect_err("symbolic-link destination");

    assert!(matches!(error, SaveError::Publication(_)));
    assert!(!error.was_committed());
    assert_eq!(std::fs::read(&target)?, b"untouched target");
    assert!(
        std::fs::symlink_metadata(&destination)?
            .file_type()
            .is_symlink()
    );
    Ok(())
}

#[test]
fn save_error_formatting_redacts_write_and_publication_details() -> TestResult {
    const SECRET: &str = "numbers-save-secret";

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
    let mut writer = FailingWriter;
    let write = package
        .write_to(&mut writer)
        .expect_err("the injected writer must fail");
    let write_error = SaveError::Write(write);
    assert!(!format!("{write_error:?}").contains(SECRET));
    assert!(!write_error.to_string().contains(SECRET));
    assert_eq!(
        write_error.write_error().map(|error| error.bytes_written()),
        Some(0)
    );
    assert!(!write_error.was_committed());

    let directory = tempfile::tempdir()?;
    let destination = directory.path().join(SECRET).join("saved.numbers");
    let publication_error = package
        .save(&destination)
        .expect_err("missing parent should fail publication");
    assert!(matches!(publication_error, SaveError::Publication(_)));
    assert!(!format!("{publication_error:?}").contains(SECRET));
    assert!(!publication_error.to_string().contains(SECRET));
    Ok(())
}

#[test]
fn package_and_save_error_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}

    assert_send_sync::<Package>();
    assert_send_sync::<SaveError>();
}
