//! Atomic filesystem replacement for finalized package artifacts.

use std::fs::{self, File, Permissions};
use std::io::{self, BufWriter, Write};
use std::path::Path;

use litchi_core::Durability;
use tempfile::{NamedTempFile, PersistError};

use crate::error::{OpcError, Result};

/// Bytes staged before the temporary artifact sees a write syscall.
///
/// The preservation writer emits a local header, a payload chunk, an optional
/// descriptor and a central record per member, so an unbuffered temporary file
/// took one syscall per framing record. A page-sized staging buffer collapses
/// those into whole-buffer writes without changing a single published byte.
const STAGING_BUFFER_BYTES: usize = 64 * 1024;

/// The temporary artifact handed to an atomic replacement closure.
///
/// Writes are staged in a fixed buffer and flushed before the artifact is
/// synchronized and persisted, so finalization reaches the filesystem in
/// page-sized writes. The buffer is private to this module: a caller-owned
/// sink passed to a streaming writer is never wrapped, so its incomplete-output
/// accounting still counts bytes the caller's own sink accepted. Buffered bytes
/// that a failed finalization leaves unflushed belong to a temporary artifact
/// that is discarded without replacing the destination.
pub struct AtomicSink<'file> {
    inner: BufWriter<&'file mut File>,
}

impl<'file> AtomicSink<'file> {
    fn new(file: &'file mut File) -> Self {
        Self {
            inner: BufWriter::with_capacity(STAGING_BUFFER_BYTES, file),
        }
    }

    /// Flush every staged byte to the temporary artifact.
    fn finish(self) -> io::Result<()> {
        self.inner
            .into_inner()
            .map_err(io::Error::from)
            .map(|_file| ())
    }
}

impl Write for AtomicSink<'_> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.inner.write(bytes)
    }

    fn write_all(&mut self, bytes: &[u8]) -> io::Result<()> {
        self.inner.write_all(bytes)
    }

    fn write_vectored(&mut self, buffers: &[io::IoSlice<'_>]) -> io::Result<usize> {
        self.inner.write_vectored(buffers)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl std::fmt::Debug for AtomicSink<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("AtomicSink").finish_non_exhaustive()
    }
}

/// Finalize an artifact in a sibling temporary file, then replace `path`.
///
/// The destination is left untouched when `write` fails. Existing regular-file
/// permissions are preserved on Unix, and symbolic-link or non-file
/// destinations are rejected before any replacement is attempted. Windows
/// permission preservation is not currently promised. This is
/// [`replace_with_durability`] at [`Durability::Full`].
///
/// # Errors
///
/// Returns an error when `path` is not a usable file destination, when the
/// temporary file cannot be written, synchronized, or persisted, or when the
/// `write` callback fails. If the destination was already replaced but the
/// parent directory could not be synchronized, the error is
/// [`OpcError::Committed`].
pub fn replace(path: &Path, write: impl FnOnce(&mut AtomicSink<'_>) -> Result<()>) -> Result<()> {
    replace_with(path, write)
}

/// Atomically replace `path` while preserving a caller-owned typed error.
///
/// Filesystem validation and persistence failures are converted from
/// [`OpcError`], while an error returned by `write` is passed through exactly.
/// This is [`replace_with_durability`] at [`Durability::Full`].
///
/// # Errors
///
/// Returns an error when `path` is not a usable file destination, when the
/// temporary file cannot be written, synchronized, or persisted, or when the
/// `write` callback fails. Filesystem failures are converted into `E` from
/// [`OpcError`]; a `write` failure is returned unchanged.
pub fn replace_with<E>(
    path: &Path,
    write: impl FnOnce(&mut AtomicSink<'_>) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
{
    replace_with_durability(path, Durability::Full, write)
}

/// Atomically replace `path` at a caller-chosen [`Durability`].
///
/// Every level validates the destination, stages the complete artifact in a
/// sibling temporary file, preserves existing Unix permissions and replaces
/// the destination with one rename, exactly as [`replace_with`] does, and a
/// failure before the rename leaves the destination untouched and removes the
/// temporary file. The level decides only the synchronizations:
/// [`Durability::Full`] synchronizes the temporary file before the rename and
/// the parent directory after it, [`Durability::FileOnly`] only the temporary
/// file, and [`Durability::NoSync`] neither. [`Durability`] states what each
/// level guarantees after a crash. A skipped step is never attempted, so it
/// cannot fail.
///
/// # Errors
///
/// As [`replace_with`], except that only [`Durability::Full`] synchronizes
/// the parent directory and can therefore return [`OpcError::Committed`], and
/// [`Durability::NoSync`] cannot fail to synchronize the temporary file.
pub fn replace_with_durability<E>(
    path: &Path,
    durability: Durability,
    write: impl FnOnce(&mut AtomicSink<'_>) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
{
    replace_with_steps(path, durability, &mut SystemSteps, write)
}

/// The filesystem operations of one replacement whose calls a test double
/// can observe or fail. Production uses [`SystemSteps`].
trait PublishSteps {
    /// Synchronizes the complete staged temporary file. Called only by a level
    /// whose [`Durability::syncs_file`] is true.
    fn sync_file(&mut self, file: &File) -> io::Result<()>;

    /// Renames the staged temporary file over `destination`.
    fn persist(
        &mut self,
        temporary: NamedTempFile,
        destination: &Path,
    ) -> std::result::Result<File, PersistError>;

    /// Synchronizes the destination's parent directory after the rename.
    /// Called only by a level whose [`Durability::syncs_directory`] is true.
    fn sync_parent(&mut self, parent: &Path) -> io::Result<()>;
}

/// The real filesystem operations.
struct SystemSteps;

impl PublishSteps for SystemSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()> {
        file.sync_all()
    }

    fn persist(
        &mut self,
        temporary: NamedTempFile,
        destination: &Path,
    ) -> std::result::Result<File, PersistError> {
        temporary.persist(destination)
    }

    fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
        sync_parent(parent)
    }
}

// Only the unix-gated parent-sync failure test uses this entry point; gating it
// the same way keeps the Windows test build free of dead code.
#[cfg(all(test, unix))]
fn replace_with_impl<E, S>(
    path: &Path,
    write: impl FnOnce(&mut AtomicSink<'_>) -> std::result::Result<(), E>,
    sync: S,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
    S: FnOnce(&Path) -> io::Result<()>,
{
    replace_with_steps(
        path,
        Durability::Full,
        &mut testing::ParentSyncHook::new(sync),
        write,
    )
}

fn replace_with_steps<E, S>(
    path: &Path,
    durability: Durability,
    steps: &mut S,
    write: impl FnOnce(&mut AtomicSink<'_>) -> std::result::Result<(), E>,
) -> std::result::Result<(), E>
where
    E: From<OpcError>,
    S: PublishSteps,
{
    if path.file_name().is_none() {
        return Err(E::from(invalid_path(
            "package destination must name a file",
        )));
    }

    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let permissions = destination_permissions(path).map_err(E::from)?;
    let mut temporary = tempfile::Builder::new()
        .prefix(".litchi-")
        .suffix(".tmp")
        .tempfile_in(parent)
        .map_err(OpcError::from)
        .map_err(E::from)?;

    let mut sink = AtomicSink::new(temporary.as_file_mut());
    write(&mut sink)?;
    sink.finish().map_err(OpcError::from).map_err(E::from)?;
    if let Some(existing_permissions) = permissions {
        temporary
            .as_file()
            .set_permissions(existing_permissions)
            .map_err(OpcError::from)
            .map_err(E::from)?;
    }
    if durability.syncs_file() {
        steps
            .sync_file(temporary.as_file())
            .map_err(OpcError::from)
            .map_err(E::from)?;
    }
    let _persisted = steps
        .persist(temporary, path)
        .map_err(|error| OpcError::IoError(error.error))
        .map_err(E::from)?;
    if durability.syncs_directory() {
        steps
            .sync_parent(parent)
            .map_err(|source| E::from(OpcError::Committed { source }))?;
    }
    Ok(())
}

fn destination_permissions(path: &Path) -> Result<Option<Permissions>> {
    match fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(invalid_path(
            "refusing to replace a package destination through a symbolic link",
        )),
        Ok(metadata) if !metadata.is_file() => {
            Err(invalid_path("package destination is not a regular file"))
        },
        Ok(metadata) => Ok(Some(metadata.permissions())),
        Err(error) if error.kind() == io::ErrorKind::NotFound => Ok(None),
        Err(error) => Err(error.into()),
    }
}

#[cfg(unix)]
fn sync_parent(parent: &Path) -> io::Result<()> {
    match File::open(parent).and_then(|directory| directory.sync_all()) {
        Ok(()) => Ok(()),
        Err(error)
            if matches!(
                error.kind(),
                io::ErrorKind::InvalidInput | io::ErrorKind::Unsupported
            ) =>
        {
            Ok(())
        },
        Err(error) => Err(error),
    }
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> io::Result<()> {
    Ok(())
}

fn invalid_path(message: &'static str) -> OpcError {
    OpcError::IoError(io::Error::new(io::ErrorKind::InvalidInput, message))
}

/// Test doubles for the replacement steps.
#[cfg(test)]
mod testing {
    use std::fs::File;
    use std::io;
    use std::path::Path;

    use tempfile::{NamedTempFile, PersistError};

    use super::{PublishSteps, sync_parent};

    /// Real steps whose parent synchronization is a caller hook; the shape of
    /// the pre-0761 private test entry point, whose only test is unix-only.
    #[cfg(unix)]
    pub(super) struct ParentSyncHook<S> {
        sync: Option<S>,
    }

    #[cfg(unix)]
    impl<S> ParentSyncHook<S> {
        pub(super) fn new(sync: S) -> Self {
            Self { sync: Some(sync) }
        }
    }

    #[cfg(unix)]
    impl<S> PublishSteps for ParentSyncHook<S>
    where
        S: FnOnce(&Path) -> io::Result<()>,
    {
        fn sync_file(&mut self, file: &File) -> io::Result<()> {
            file.sync_all()
        }

        fn persist(
            &mut self,
            temporary: NamedTempFile,
            destination: &Path,
        ) -> Result<File, PersistError> {
            temporary.persist(destination)
        }

        fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
            match self.sync.take() {
                Some(sync) => sync(parent),
                None => Err(io::Error::other("parent-sync hook called twice")),
            }
        }
    }

    /// Real steps that count every call and can fail a synchronization.
    #[derive(Debug, Default)]
    pub(super) struct CountingSteps {
        pub(super) file_syncs: usize,
        pub(super) persists: usize,
        pub(super) parent_syncs: usize,
        pub(super) fail_file_sync: bool,
        pub(super) fail_parent_sync: bool,
    }

    impl CountingSteps {
        /// `(file syncs, parent syncs, renames)`.
        pub(super) fn counts(&self) -> (usize, usize, usize) {
            (self.file_syncs, self.parent_syncs, self.persists)
        }
    }

    impl PublishSteps for CountingSteps {
        fn sync_file(&mut self, file: &File) -> io::Result<()> {
            self.file_syncs += 1;
            if self.fail_file_sync {
                return Err(io::Error::other("injected file sync failure"));
            }
            file.sync_all()
        }

        fn persist(
            &mut self,
            temporary: NamedTempFile,
            destination: &Path,
        ) -> Result<File, PersistError> {
            self.persists += 1;
            temporary.persist(destination)
        }

        fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
            self.parent_syncs += 1;
            if self.fail_parent_sync {
                return Err(io::Error::other("injected parent sync failure"));
            }
            sync_parent(parent)
        }
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use std::io::Write;

    use super::*;

    #[derive(Debug)]
    enum TypedError {
        Opc(OpcError),
        Write,
    }

    impl From<OpcError> for TypedError {
        fn from(error: OpcError) -> Self {
            Self::Opc(error)
        }
    }

    impl std::fmt::Display for TypedError {
        fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
            match self {
                Self::Opc(error) => write!(formatter, "{error}"),
                Self::Write => formatter.write_str("typed write failure"),
            }
        }
    }

    impl std::error::Error for TypedError {}

    #[test]
    fn failed_finalization_leaves_the_destination_untouched() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace(&destination, |temporary| {
            temporary.write_all(b"partial")?;
            Err(OpcError::InvalidRelationship(
                "injected finalization failure".to_owned(),
            ))
        });

        assert!(matches!(result, Err(OpcError::InvalidRelationship(_))));
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"original"
        );
        assert_eq!(
            fs::read_dir(directory.path())
                .expect("list temporary directory")
                .count(),
            1
        );
    }

    #[test]
    fn successful_finalization_replaces_the_destination() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");

        replace(&destination, |temporary| {
            temporary.write_all(b"new")?;
            Ok(())
        })
        .expect("atomic replacement");

        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[test]
    fn staged_writes_reach_the_destination_in_order() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        let payload: Vec<u8> = (0..(STAGING_BUFFER_BYTES * 3 + 7))
            .map(|index| u8::try_from(index % 251).expect("payload byte"))
            .collect();

        let staged = payload.clone();
        replace(&destination, move |temporary| {
            // One byte at a time is the shape the preservation writer used to
            // hand the raw file: framing records far smaller than a page.
            for byte in &staged {
                temporary.write_all(std::slice::from_ref(byte))?;
            }
            Ok(())
        })
        .expect("atomic replacement");

        assert_eq!(fs::read(destination).expect("read destination"), payload);
    }

    #[test]
    fn a_failed_write_after_staging_leaves_the_destination_untouched() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace(&destination, |temporary| {
            temporary.write_all(&vec![7_u8; STAGING_BUFFER_BYTES / 2])?;
            Err(OpcError::InvalidRelationship(
                "injected staged failure".to_owned(),
            ))
        });

        assert!(matches!(result, Err(OpcError::InvalidRelationship(_))));
        assert_eq!(
            fs::read(&destination).expect("read destination"),
            b"original"
        );
        assert_eq!(
            fs::read_dir(directory.path())
                .expect("list temporary directory")
                .count(),
            1
        );
    }

    #[cfg(unix)]
    #[test]
    fn directory_sync_failure_reports_that_replacement_already_committed() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");

        let result = replace_with_impl::<TypedError, _>(
            &destination,
            |temporary| {
                temporary
                    .write_all(b"new")
                    .map_err(|error| TypedError::Opc(OpcError::IoError(error)))?;
                Ok(())
            },
            |_parent| {
                Err(io::Error::new(
                    io::ErrorKind::PermissionDenied,
                    "injected directory sync failure",
                ))
            },
        );

        assert!(matches!(
            result,
            Err(TypedError::Opc(OpcError::Committed { source }))
                if source.kind() == io::ErrorKind::PermissionDenied
        ));
        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[test]
    fn caller_owned_write_error_is_not_erased() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.docx");
        fs::write(&destination, b"original").expect("seed destination");

        let result = replace_with::<TypedError>(&destination, |_temporary| Err(TypedError::Write));

        assert!(matches!(result, Err(TypedError::Write)));
        assert_eq!(
            fs::read(destination).expect("read destination"),
            b"original"
        );
    }

    #[cfg(unix)]
    #[test]
    fn replacement_preserves_existing_file_permissions() {
        use std::os::unix::fs::PermissionsExt;

        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.xlsx");
        fs::write(&destination, b"old").expect("seed destination");
        fs::set_permissions(&destination, Permissions::from_mode(0o640))
            .expect("set destination permissions");

        replace(&destination, |temporary| {
            temporary.write_all(b"new")?;
            Ok(())
        })
        .expect("atomic replacement");

        let mode = fs::metadata(destination)
            .expect("destination metadata")
            .permissions()
            .mode()
            & 0o777;
        assert_eq!(mode, 0o640);
    }

    #[cfg(unix)]
    #[test]
    fn symbolic_link_destinations_are_refused_without_touching_the_target() {
        use std::os::unix::fs::symlink;

        let directory = tempfile::tempdir().expect("temporary directory");
        let target = directory.path().join("target.xlsx");
        let link = directory.path().join("report.xlsx");
        fs::write(&target, b"original").expect("seed target");
        symlink(&target, &link).expect("create symbolic link");

        let result = replace(&link, |temporary| {
            temporary.write_all(b"replacement")?;
            Ok(())
        });

        assert!(matches!(
            result,
            Err(OpcError::IoError(error)) if error.kind() == io::ErrorKind::InvalidInput
        ));
        assert_eq!(fs::read(target).expect("read target"), b"original");
        assert!(
            fs::symlink_metadata(link)
                .expect("link metadata")
                .file_type()
                .is_symlink()
        );
    }

    const LEVELS: [Durability; 3] = [Durability::Full, Durability::FileOnly, Durability::NoSync];

    fn directory_entries(directory: &Path) -> Vec<std::ffi::OsString> {
        let mut names: Vec<_> = fs::read_dir(directory)
            .expect("list directory")
            .map(|entry| entry.expect("directory entry").file_name())
            .collect();
        names.sort();
        names
    }

    #[test]
    fn each_level_makes_exactly_its_synchronizations() {
        // (file syncs, parent syncs, renames) per level, as change 0761 fixes.
        let expected = [(1, 1, 1), (1, 0, 1), (0, 0, 1)];
        for (durability, expected) in LEVELS.into_iter().zip(expected) {
            let directory = tempfile::tempdir().expect("temporary directory");
            let destination = directory.path().join("report.docx");
            fs::write(&destination, b"old").expect("seed destination");
            let mut steps = testing::CountingSteps::default();

            replace_with_steps::<OpcError, _>(&destination, durability, &mut steps, |temporary| {
                temporary.write_all(b"new")?;
                Ok(())
            })
            .expect("atomic replacement");

            assert_eq!(steps.counts(), expected, "{durability:?}");
            assert_eq!(fs::read(&destination).expect("read destination"), b"new");
            assert_eq!(
                directory_entries(directory.path()),
                vec![std::ffi::OsString::from("report.docx")],
                "{durability:?} left a temporary file behind"
            );
        }
    }

    #[test]
    fn file_sync_failure_leaves_the_destination_and_removes_the_temporary() {
        for durability in [Durability::Full, Durability::FileOnly] {
            let directory = tempfile::tempdir().expect("temporary directory");
            let destination = directory.path().join("report.docx");
            fs::write(&destination, b"original").expect("seed destination");
            let mut steps = testing::CountingSteps {
                fail_file_sync: true,
                ..testing::CountingSteps::default()
            };

            let result = replace_with_steps::<OpcError, _>(
                &destination,
                durability,
                &mut steps,
                |temporary| {
                    temporary.write_all(b"replacement")?;
                    Ok(())
                },
            );

            assert!(
                matches!(result, Err(OpcError::IoError(ref error)) if error.to_string() == "injected file sync failure"),
                "{durability:?}: {result:?}"
            );
            assert_eq!(steps.counts(), (1, 0, 0), "{durability:?}");
            assert_eq!(
                fs::read(&destination).expect("read destination"),
                b"original"
            );
            assert_eq!(
                directory_entries(directory.path()),
                vec![std::ffi::OsString::from("report.docx")],
                "{durability:?} left a temporary file behind"
            );
        }
    }

    #[test]
    fn no_sync_never_attempts_the_file_sync_it_skips() {
        let directory = tempfile::tempdir().expect("temporary directory");
        let destination = directory.path().join("report.docx");
        fs::write(&destination, b"old").expect("seed destination");
        let mut steps = testing::CountingSteps {
            fail_file_sync: true,
            fail_parent_sync: true,
            ..testing::CountingSteps::default()
        };

        replace_with_steps::<OpcError, _>(
            &destination,
            Durability::NoSync,
            &mut steps,
            |temporary| {
                temporary.write_all(b"new")?;
                Ok(())
            },
        )
        .expect("a skipped synchronization cannot fail");

        assert_eq!(steps.counts(), (0, 0, 1));
        assert_eq!(fs::read(destination).expect("read destination"), b"new");
    }

    #[test]
    fn parent_sync_failure_is_committed_only_under_full() {
        for durability in LEVELS {
            let directory = tempfile::tempdir().expect("temporary directory");
            let destination = directory.path().join("report.docx");
            fs::write(&destination, b"old").expect("seed destination");
            let mut steps = testing::CountingSteps {
                fail_parent_sync: true,
                ..testing::CountingSteps::default()
            };

            let result = replace_with_steps::<OpcError, _>(
                &destination,
                durability,
                &mut steps,
                |temporary| {
                    temporary.write_all(b"new")?;
                    Ok(())
                },
            );

            if durability == Durability::Full {
                assert!(
                    matches!(result, Err(OpcError::Committed { .. })),
                    "{result:?}"
                );
                assert_eq!(steps.parent_syncs, 1);
            } else {
                assert!(result.is_ok(), "{durability:?}: {result:?}");
                assert_eq!(steps.parent_syncs, 0, "{durability:?}");
            }
            // The rename happened at every level before any directory sync.
            assert_eq!(fs::read(&destination).expect("read destination"), b"new");
            assert_eq!(
                directory_entries(directory.path()),
                vec![std::ffi::OsString::from("report.docx")]
            );
        }
    }

    #[test]
    fn every_level_publishes_identical_bytes() {
        let payload: Vec<u8> = (0..(STAGING_BUFFER_BYTES * 2 + 13))
            .map(|index| u8::try_from(index % 241).expect("payload byte"))
            .collect();
        let mut published = Vec::new();
        for durability in LEVELS {
            let directory = tempfile::tempdir().expect("temporary directory");
            let destination = directory.path().join("report.docx");
            let staged = payload.clone();
            replace_with_durability::<OpcError>(&destination, durability, move |temporary| {
                temporary.write_all(&staged)?;
                Ok(())
            })
            .expect("atomic replacement");
            published.push(fs::read(&destination).expect("read destination"));
        }
        assert!(published.iter().all(|bytes| *bytes == payload));
    }

    #[test]
    fn weaker_levels_keep_every_destination_check_and_failure_rule() {
        for durability in [Durability::FileOnly, Durability::NoSync] {
            let directory = tempfile::tempdir().expect("temporary directory");

            // A failed write leaves the destination and removes the temporary.
            let destination = directory.path().join("report.docx");
            fs::write(&destination, b"original").expect("seed destination");
            let result = replace_with_durability::<TypedError>(&destination, durability, |_| {
                Err(TypedError::Write)
            });
            assert!(matches!(result, Err(TypedError::Write)), "{durability:?}");
            assert_eq!(
                fs::read(&destination).expect("read destination"),
                b"original"
            );
            assert_eq!(
                directory_entries(directory.path()),
                vec![std::ffi::OsString::from("report.docx")]
            );

            // A destination that does not name a file is refused.
            let refused =
                replace_with_durability::<OpcError>(directory.path(), durability, |_| Ok(()));
            assert!(
                matches!(refused, Err(OpcError::IoError(ref error)) if error.kind() == io::ErrorKind::InvalidInput),
                "{durability:?}: {refused:?}"
            );
        }
    }

    #[cfg(unix)]
    #[test]
    fn weaker_levels_preserve_permissions_and_refuse_symbolic_links() {
        use std::os::unix::fs::{PermissionsExt, symlink};

        for durability in [Durability::FileOnly, Durability::NoSync] {
            let directory = tempfile::tempdir().expect("temporary directory");
            let destination = directory.path().join("report.xlsx");
            fs::write(&destination, b"old").expect("seed destination");
            fs::set_permissions(&destination, Permissions::from_mode(0o640))
                .expect("set destination permissions");
            replace_with_durability::<OpcError>(&destination, durability, |temporary| {
                temporary.write_all(b"new")?;
                Ok(())
            })
            .expect("atomic replacement");
            let mode = fs::metadata(&destination)
                .expect("destination metadata")
                .permissions()
                .mode()
                & 0o777;
            assert_eq!(mode, 0o640, "{durability:?}");

            let target = directory.path().join("target.xlsx");
            let link = directory.path().join("link.xlsx");
            fs::write(&target, b"original").expect("seed target");
            symlink(&target, &link).expect("create symbolic link");
            let result = replace_with_durability::<OpcError>(&link, durability, |temporary| {
                temporary.write_all(b"replacement")?;
                Ok(())
            });
            assert!(
                matches!(result, Err(OpcError::IoError(ref error)) if error.kind() == io::ErrorKind::InvalidInput),
                "{durability:?}"
            );
            assert_eq!(fs::read(&target).expect("read target"), b"original");
        }
    }

    #[test]
    fn committed_converts_into_the_typed_core_variant() {
        let error: litchi_core::Error = OpcError::Committed {
            source: io::Error::other("directory sync failed"),
        }
        .into();
        assert!(
            matches!(&error, litchi_core::Error::Committed(source) if source.to_string() == "directory sync failed"),
            "{error:?}"
        );
    }
}
