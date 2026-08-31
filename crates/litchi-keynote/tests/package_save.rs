use std::io::{self, Write};
use std::path::PathBuf;

use litchi_keynote::{Package, SaveError};
use tempfile::tempdir;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
}

#[test]
fn saves_a_new_package_and_reopens_it() -> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let directory = tempdir()?;
    let destination = directory.path().join("saved.key");

    package.save(&destination)?;

    assert_eq!(std::fs::read(&destination)?, source);
    let reopened = Package::open(&destination)?;
    let mut reopened_bytes = Vec::new();
    reopened.write_to(&mut reopened_bytes)?;
    assert_eq!(reopened_bytes, source);
    Ok(())
}

#[test]
fn replaces_an_existing_package() -> Result<(), Box<dyn std::error::Error>> {
    let source = std::fs::read(fixture_path())?;
    let package = Package::open(fixture_path())?;
    let directory = tempdir()?;
    let destination = directory.path().join("saved.key");
    std::fs::write(&destination, b"old package bytes")?;

    package.save(&destination)?;

    assert_eq!(std::fs::read(&destination)?, source);
    let reopened = Package::open(&destination)?;
    let mut reopened_bytes = Vec::new();
    reopened.write_to(&mut reopened_bytes)?;
    assert_eq!(reopened_bytes, source);
    Ok(())
}

#[test]
fn saves_back_to_the_same_source_path() -> Result<(), Box<dyn std::error::Error>> {
    let directory = tempdir()?;
    let source_path = directory.path().join("source.key");
    let source = std::fs::read(fixture_path())?;
    std::fs::write(&source_path, &source)?;
    let package = Package::open(&source_path)?;

    package.save(&source_path)?;

    assert_eq!(std::fs::read(&source_path)?, source);
    let reopened = Package::open(&source_path)?;
    let mut reopened_bytes = Vec::new();
    reopened.write_to(&mut reopened_bytes)?;
    assert_eq!(reopened_bytes, source);
    Ok(())
}

#[test]
fn invalid_destinations_are_reported_without_replacement() -> Result<(), Box<dyn std::error::Error>>
{
    let package = Package::open(fixture_path())?;
    let directory = tempdir()?;
    let destination = directory.path().join("existing.key");
    std::fs::write(&destination, b"untouched")?;

    let error = package
        .save(directory.path())
        .expect_err("a directory cannot be a package destination");
    assert!(error.publication_error().is_some());
    assert!(!error.was_committed());
    assert_eq!(std::fs::read(&destination)?, b"untouched");
    Ok(())
}

#[cfg(unix)]
#[test]
fn symbolic_link_destinations_are_rejected_without_touching_the_target()
-> Result<(), Box<dyn std::error::Error>> {
    use std::os::unix::fs::symlink;

    let package = Package::open(fixture_path())?;
    let directory = tempdir()?;
    let target = directory.path().join("target.key");
    let link = directory.path().join("saved.key");
    std::fs::write(&target, b"untouched")?;
    symlink(&target, &link)?;

    let error = package
        .save(&link)
        .expect_err("a symbolic link cannot be a package destination");
    assert!(error.publication_error().is_some());
    assert!(!error.was_committed());
    assert_eq!(std::fs::read(&target)?, b"untouched");
    assert!(std::fs::symlink_metadata(&link)?.file_type().is_symlink());
    Ok(())
}

#[test]
fn save_errors_redact_underlying_sink_messages() -> Result<(), Box<dyn std::error::Error>> {
    const SECRET: &str = "private-keynote-save-secret";

    struct SecretWriter;

    impl Write for SecretWriter {
        fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
            Err(io::Error::other(SECRET))
        }

        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    let package = Package::open(fixture_path())?;
    let write_error = package
        .write_to(&mut SecretWriter)
        .expect_err("the test sink must reject package bytes");
    let error = SaveError::Write(write_error);

    assert!(error.write_error().is_some());
    assert!(!error.was_committed());
    assert!(!format!("{error:?}").contains(SECRET));
    assert!(!error.to_string().contains(SECRET));

    let directory = tempdir()?;
    let destination = directory.path().join(SECRET).join("saved.key");
    let publication_error = package
        .save(&destination)
        .expect_err("a missing parent must fail publication");
    assert!(publication_error.write_error().is_none());
    assert!(publication_error.publication_error().is_some());
    assert!(!publication_error.was_committed());
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
