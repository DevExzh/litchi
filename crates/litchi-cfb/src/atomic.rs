//! Atomic publication for artifacts produced by the CFB and legacy Office
//! formats.
//!
//! The publisher stages the complete artifact in the CFB writer's sibling
//! temporary file and delegates flushing, synchronization, identity checking,
//! replacement, cleanup, and committed-state reporting to
//! `crate::writer::publish_staged`.  This module only supplies the public
//! overwrite policy and the destination-specific checks needed by callers
//! that do not own a CFB writer.

use std::error::Error;
use std::fs::{self, File, Permissions};
use std::io::{self, BufWriter, Write};
use std::path::Path;

use litchi_core::Durability;

use crate::file::OleError;
use crate::writer::{PublishFailure, PublishSteps, SystemSteps, TemporaryIdentity, publish_staged};

/// Whether publication may replace an existing regular-file destination.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Overwrite {
    /// Atomically replace an existing regular file.
    Replace,
    /// Create the destination only if it is still absent at publication time.
    Refuse,
}

/// Stage and atomically publish one artifact.
///
/// `stage` receives the sibling temporary file's buffered writer and returns
/// the caller's typed result after writing the complete artifact. `verify` is
/// called after staging has flushed and, when requested by `durability`, the
/// temporary file has been synchronized, but before publication. Filesystem
/// failures are converted through `E::from(OleError)`; values returned by the
/// two callbacks are preserved.
///
/// Existing regular-file permissions are copied to a replacement temporary
/// file. Symbolic links and non-regular destinations are rejected before
/// staging. `Refuse` uses an atomic hard-link publication, so a destination
/// created after staging still wins the race and is reported as an I/O error
/// with `AlreadyExists` preserved by the operating system.
pub fn save<T, E, P, Stage, Verify>(
    destination: P,
    durability: Durability,
    overwrite: Overwrite,
    stage: Stage,
    verify: Verify,
) -> Result<T, E>
where
    P: AsRef<Path>,
    Stage: FnOnce(&mut BufWriter<File>) -> Result<T, E>,
    Verify: FnOnce() -> Result<(), E>,
    E: From<OleError>,
{
    let destination = destination.as_ref();
    let permissions = destination_permissions(destination).map_err(E::from)?;
    let stage = move |staging: &mut BufWriter<File>| {
        let staged = stage(staging)?;
        if let Some(permissions) = permissions {
            staging
                .flush()
                .map_err(|source| E::from(OleError::Io(source)))?;
            staging
                .get_ref()
                .set_permissions(permissions)
                .map_err(|source| E::from(OleError::Io(source)))?;
        }
        Ok(staged)
    };

    match overwrite {
        Overwrite::Replace => {
            let mut steps = SystemSteps;
            publish(destination, durability, &mut steps, stage, verify)
        },
        Overwrite::Refuse => {
            let mut steps = RefuseSteps {
                system: SystemSteps,
            };
            publish(destination, durability, &mut steps, stage, verify)
        },
    }
}

fn publish<T, E, S, Stage, Verify>(
    destination: &Path,
    durability: Durability,
    steps: &mut S,
    stage: Stage,
    verify: Verify,
) -> Result<T, E>
where
    S: PublishSteps + ?Sized,
    Stage: FnOnce(&mut BufWriter<File>) -> Result<T, E>,
    Verify: FnOnce() -> Result<(), E>,
    E: From<OleError>,
{
    publish_staged(
        destination,
        durability,
        TemporaryIdentity::Checked,
        steps,
        stage,
        move |_staged_file, _staged| verify(),
    )
    .map_err(map_failure)
}

fn map_failure<T, E>(failure: PublishFailure<T, E>) -> E
where
    E: From<OleError>,
{
    match failure {
        PublishFailure::Create(error) => E::from(error),
        PublishFailure::Caller(error) => error,
        PublishFailure::Flush { source, .. }
        | PublishFailure::SyncFile { source, .. }
        | PublishFailure::Identity { source, .. } => E::from(OleError::Io(source)),
        PublishFailure::Replace { source, .. } => E::from(replacement_error(source)),
        PublishFailure::Committed { source, .. } => E::from(OleError::Committed { source }),
    }
}

fn destination_permissions(path: &Path) -> Result<Option<Permissions>, OleError> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(OleError::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "CFB destination must not be a symbolic link",
        ))),
        Ok(metadata) if !metadata.is_file() => Err(OleError::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "CFB destination is not a regular file",
        ))),
        Ok(metadata) => Ok(Some(metadata.permissions())),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(OleError::Io(error)),
    }
}

/// A no-clobber publisher implemented on top of the shared staged pipeline.
///
/// `hard_link` is the publication operation: it succeeds only when the
/// destination name is absent, and therefore closes the check-to-publish race
/// that an existence preflight would leave open.
struct RefuseSteps {
    system: SystemSteps,
}

impl PublishSteps for RefuseSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()> {
        self.system.sync_file(file)
    }

    fn replace(&mut self, temporary: &Path, destination: &Path) -> io::Result<()> {
        fs::hard_link(temporary, destination)?;
        if let Err(source) = fs::remove_file(temporary) {
            return Err(io::Error::other(RefuseCleanupCommitted { source }));
        }
        Ok(())
    }

    fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
        self.system.sync_parent(parent)
    }
}

fn replacement_error(source: io::Error) -> OleError {
    if source
        .get_ref()
        .and_then(|cause| cause.downcast_ref::<RefuseCleanupCommitted>())
        .is_some()
    {
        OleError::Committed { source }
    } else {
        OleError::Io(source)
    }
}

/// Marker retained inside the typed I/O error when no-clobber linked the
/// destination but could not remove the now-unneeded temporary name.
#[derive(Debug)]
struct RefuseCleanupCommitted {
    source: io::Error,
}

impl std::fmt::Display for RefuseCleanupCommitted {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "temporary cleanup failed after publication: {}",
            self.source
        )
    }
}

impl Error for RefuseCleanupCommitted {
    fn source(&self) -> Option<&(dyn Error + 'static)> {
        Some(&self.source)
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "focused atomic publication tests use panic-on-failure assertions"
    )]

    use super::*;
    use std::fs;
    use std::io::{self, Write};
    use std::path::PathBuf;
    use std::sync::atomic::{AtomicU64, Ordering};

    struct Scratch(PathBuf);

    impl Scratch {
        fn new(label: &str) -> Self {
            static NEXT: AtomicU64 = AtomicU64::new(0);
            let path = std::env::temp_dir().join(format!(
                "litchi-cfb-atomic-{label}-{}-{}",
                std::process::id(),
                NEXT.fetch_add(1, Ordering::Relaxed)
            ));
            fs::create_dir(&path).unwrap();
            Self(path)
        }

        fn join(&self, name: &str) -> PathBuf {
            self.0.join(name)
        }

        fn entry_count(&self) -> usize {
            fs::read_dir(&self.0).unwrap().count()
        }
    }

    impl Drop for Scratch {
        fn drop(&mut self) {
            drop(fs::remove_dir_all(&self.0));
        }
    }

    #[derive(Debug)]
    enum CallbackError {
        Stage,
        Verify,
        Ole,
    }

    impl From<OleError> for CallbackError {
        fn from(_: OleError) -> Self {
            Self::Ole
        }
    }

    #[test]
    fn replace_publishes_and_preserves_existing_permissions() {
        let scratch = Scratch::new("replace");
        let destination = scratch.join("document.ole");
        fs::write(&destination, b"old").unwrap();

        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;

            fs::set_permissions(&destination, Permissions::from_mode(0o640)).unwrap();
        }

        let result: Result<u64, OleError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            |staging| {
                staging.write_all(b"new").unwrap();
                Ok(3)
            },
            || Ok(()),
        );

        assert_eq!(result.unwrap(), 3);
        assert_eq!(fs::read(&destination).unwrap(), b"new");
        assert_eq!(scratch.entry_count(), 1);

        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;

            assert_eq!(
                fs::metadata(&destination).unwrap().permissions().mode() & 0o777,
                0o640
            );
        }
    }

    #[test]
    fn refuse_existing_destination_leaves_it_and_the_temporary_clean() {
        let scratch = Scratch::new("refuse-existing");
        let destination = scratch.join("document.ole");
        fs::write(&destination, b"old").unwrap();

        let result: Result<(), OleError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Refuse,
            |staging| {
                staging.write_all(b"new").unwrap();
                Ok(())
            },
            || Ok(()),
        );

        assert!(matches!(
            result,
            Err(OleError::Io(error)) if error.kind() == io::ErrorKind::AlreadyExists
        ));
        assert_eq!(fs::read(&destination).unwrap(), b"old");
        assert_eq!(scratch.entry_count(), 1);
    }

    #[test]
    fn refuse_race_rejects_destination_created_during_staging() {
        let scratch = Scratch::new("refuse-race");
        let destination = scratch.join("document.ole");

        let result: Result<(), OleError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Refuse,
            |staging| {
                staging.write_all(b"candidate").unwrap();
                fs::write(&destination, b"raced").unwrap();
                Ok(())
            },
            || Ok(()),
        );

        assert!(matches!(
            result,
            Err(OleError::Io(error)) if error.kind() == io::ErrorKind::AlreadyExists
        ));
        assert_eq!(fs::read(&destination).unwrap(), b"raced");
        assert_eq!(scratch.entry_count(), 1);
    }

    #[test]
    fn callback_failures_preserve_their_types_and_clean_staging() {
        let scratch = Scratch::new("callbacks");
        let destination = scratch.join("document.ole");
        fs::write(&destination, b"old").unwrap();

        let stage_result: Result<(), CallbackError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            |_staging| Err(CallbackError::Stage),
            || Ok(()),
        );
        assert!(matches!(stage_result, Err(CallbackError::Stage)));
        assert_eq!(fs::read(&destination).unwrap(), b"old");
        assert_eq!(scratch.entry_count(), 1);

        let verify_result: Result<(), CallbackError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            |staging| {
                staging.write_all(b"candidate").unwrap();
                Ok(())
            },
            || Err(CallbackError::Verify),
        );
        assert!(matches!(verify_result, Err(CallbackError::Verify)));
        assert_eq!(fs::read(&destination).unwrap(), b"old");
        assert_eq!(scratch.entry_count(), 1);
    }

    #[cfg(unix)]
    #[test]
    fn symbolic_link_destination_is_rejected_before_staging() {
        use std::os::unix::fs::symlink;

        let scratch = Scratch::new("symlink");
        let target = scratch.join("target.ole");
        let destination = scratch.join("link.ole");
        fs::write(&target, b"old").unwrap();
        symlink(&target, &destination).unwrap();

        let result: Result<(), OleError> = save(
            &destination,
            Durability::NoSync,
            Overwrite::Replace,
            |staging| {
                staging.write_all(b"new").unwrap();
                Ok(())
            },
            || Ok(()),
        );

        assert!(matches!(
            result,
            Err(OleError::Io(error)) if error.kind() == io::ErrorKind::InvalidInput
        ));
        assert_eq!(fs::read(&target).unwrap(), b"old");
        assert_eq!(scratch.entry_count(), 2);
    }
}
