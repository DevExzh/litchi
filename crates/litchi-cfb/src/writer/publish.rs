//! The one atomic sibling-temporary publication shared by every CFB save.
//!
//! [`OleWriter::save`](super::OleWriter::save),
//! [`SequentialOleWriter::save`](super::SequentialOleWriter::save) and
//! [`ValidatedOverlayPlan::save`](crate::ValidatedOverlayPlan::save) each stage
//! a complete artifact in a sibling temporary file, optionally verify it, and
//! replace the destination with it. Until change 0761 each route spelled that
//! flush, file-sync, rename and parent-sync sequence itself; [`publish_staged`]
//! now owns it once, so the caller's [`Durability`] decides the same
//! synchronizations on every route and one test double observes all three.
//!
//! What stays per route is passed in: the staging closure, the verification
//! that runs after the staged file is complete (and synchronized, when the
//! level synchronizes it) but before the rename, and whether the temporary
//! file's identity is checked ([`TemporaryIdentity`]). Each route maps a
//! [`PublishFailure`] back onto its own typed error exactly as it did before.

use super::{atomic_replace, create_sibling_temp_file, parent_directory, sync_parent};
use crate::file::OleError;
use litchi_core::Durability;
use std::fs::{self, File};
use std::io::{self, BufWriter, ErrorKind, Write};
use std::path::Path;

/// The filesystem operations of one publication whose calls a test double
/// can observe or fail. Production uses [`SystemSteps`].
pub(crate) trait PublishSteps {
    /// Synchronizes the complete staged temporary file. Called only by a level
    /// whose [`Durability::syncs_file`] is true.
    fn sync_file(&mut self, file: &File) -> io::Result<()>;

    /// Replaces `destination` with the staged `temporary` file.
    fn replace(&mut self, temporary: &Path, destination: &Path) -> io::Result<()>;

    /// Synchronizes the destination's parent directory after the replacement.
    /// Called only by a level whose [`Durability::syncs_directory`] is true.
    fn sync_parent(&mut self, parent: &Path) -> io::Result<()>;
}

/// The real filesystem operations.
pub(crate) struct SystemSteps;

impl PublishSteps for SystemSteps {
    fn sync_file(&mut self, file: &File) -> io::Result<()> {
        file.sync_all()
    }

    fn replace(&mut self, temporary: &Path, destination: &Path) -> io::Result<()> {
        atomic_replace(temporary, destination)
    }

    fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
        sync_parent(parent)
    }
}

/// Whether a publication checks that its temporary path still names the file
/// it created.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TemporaryIdentity {
    /// Remove the temporary path after a failure (the `OleWriter` and overlay
    /// routes).
    Unchecked,
    /// Capture the temporary file's native identity when it is created,
    /// refuse to replace the destination when the path no longer names that
    /// file, and after a failure remove the path only while it still names it
    /// (the sequential route). This is a best-effort defence that still
    /// requires a trusted, private parent directory.
    Checked,
}

/// Where a publication stopped. Every variant before [`Self::Committed`]
/// leaves the destination untouched.
pub(crate) enum PublishFailure<T, E> {
    /// The sibling temporary file could not be created or identified; nothing
    /// was staged.
    Create(OleError),
    /// The caller's staging or verification closure failed.
    Caller(E),
    /// Flushing staged bytes into the temporary file failed.
    Flush { source: io::Error, staged: T },
    /// Synchronizing the temporary file failed.
    SyncFile { source: io::Error, staged: T },
    /// The temporary path no longer names the staged file.
    Identity { source: io::Error, staged: T },
    /// Replacing the destination failed.
    Replace { source: io::Error, staged: T },
    /// The destination was replaced, but synchronizing its parent directory
    /// failed. Only a level that synchronizes the directory reaches this.
    Committed { source: io::Error, staged: T },
}

/// Stages an artifact in a sibling temporary file and atomically replaces
/// `destination` with it, synchronizing as `durability` asks.
///
/// The sequence is fixed: create the temporary file (and capture its identity
/// when checked), run `stage` into a buffered writer, flush, synchronize the
/// file when [`Durability::syncs_file`], run `verify` on the still-open file,
/// close it, re-check its identity when checked, replace the destination,
/// then synchronize the parent directory when [`Durability::syncs_directory`].
/// A level never calls a step it skips, so it can never fail at one. Every
/// failure before the replacement removes the temporary file (subject to the
/// identity check) after closing every handle to it; a failure after the
/// replacement removes nothing, because the temporary name is gone.
pub(crate) fn publish_staged<T, E, S>(
    destination: &Path,
    durability: Durability,
    identity: TemporaryIdentity,
    steps: &mut S,
    stage: impl FnOnce(&mut BufWriter<File>) -> Result<T, E>,
    verify: impl FnOnce(&File, &T) -> Result<(), E>,
) -> Result<T, PublishFailure<T, E>>
where
    S: PublishSteps + ?Sized,
{
    let parent = parent_directory(destination);
    let (temporary_path, created) =
        create_sibling_temp_file(destination).map_err(PublishFailure::Create)?;
    // Declared before every handle to the temporary file, so each handle is
    // closed before this guard removes the path (Windows requires that).
    let mut cleanup = TemporaryCleanupGuard::new(&temporary_path);
    let expected = match identity {
        TemporaryIdentity::Unchecked => None,
        TemporaryIdentity::Checked => match file_identity(&created) {
            Ok(observed) => {
                cleanup.set_identity(observed);
                Some(observed)
            },
            Err(source) => {
                // Close the handle before the guard attempts exact-name
                // cleanup.
                drop(created);
                return Err(PublishFailure::Create(OleError::Io(source)));
            },
        },
    };

    let mut buffered = BufWriter::new(created);
    let staged = stage(&mut buffered).map_err(PublishFailure::Caller)?;
    if let Err(source) = buffered.flush() {
        return Err(PublishFailure::Flush { source, staged });
    }
    let file = match buffered.into_inner() {
        Ok(file) => file,
        Err(error) => {
            return Err(PublishFailure::Flush {
                source: error.into_error(),
                staged,
            });
        },
    };
    if durability.syncs_file() {
        if let Err(source) = steps.sync_file(&file) {
            return Err(PublishFailure::SyncFile { source, staged });
        }
    }
    verify(&file, &staged).map_err(PublishFailure::Caller)?;
    // Windows requires every handle to be closed before the temporary name can
    // replace the destination.
    drop(file);

    if let Some(expected) = expected {
        // Best-effort identity checking rejects a known substitution and lets
        // cleanup preserve it; portable path APIs cannot close every
        // hostile-directory race, so the parent must be trusted and private.
        if let Err(source) = ensure_temp_identity(&temporary_path, expected) {
            return Err(PublishFailure::Identity { source, staged });
        }
    }
    if let Err(source) = steps.replace(&temporary_path, destination) {
        return Err(PublishFailure::Replace { source, staged });
    }
    cleanup.mark_published();
    if durability.syncs_directory() {
        if let Err(source) = steps.sync_parent(parent) {
            return Err(PublishFailure::Committed { source, staged });
        }
    }
    Ok(staged)
}

/// Native identity of a staged temporary file.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum TempFileIdentity {
    #[cfg(unix)]
    Unix { device: u64, inode: u64 },
    #[cfg(windows)]
    Windows { volume: u32, index: u64 },
    // Targets without a native file identity get only a length discriminator.
    // The public save contract therefore requires a trusted/private parent;
    // this fallback is not protection against a hostile shared directory.
    #[cfg(not(any(unix, windows)))]
    Fallback { length: u64 },
}

fn file_identity(file: &File) -> io::Result<TempFileIdentity> {
    identity_from_metadata(&file.metadata()?)
}

pub(crate) fn path_identity(path: &Path) -> io::Result<TempFileIdentity> {
    let metadata = fs::symlink_metadata(path)?;
    if !metadata.file_type().is_file() {
        return Err(io::Error::new(
            ErrorKind::InvalidData,
            "CFB temporary path is not a regular file",
        ));
    }
    identity_from_metadata(&metadata)
}

fn identity_from_metadata(metadata: &fs::Metadata) -> io::Result<TempFileIdentity> {
    #[cfg(unix)]
    {
        use std::os::unix::fs::MetadataExt;

        return Ok(TempFileIdentity::Unix {
            device: metadata.dev(),
            inode: metadata.ino(),
        });
    }

    #[cfg(windows)]
    {
        use std::os::windows::fs::MetadataExt;

        let volume = metadata.volume_serial_number().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "CFB temporary file has no volume identity",
            )
        })?;
        let index = metadata.file_index().ok_or_else(|| {
            io::Error::new(
                ErrorKind::Unsupported,
                "CFB temporary file has no file identity",
            )
        })?;
        return Ok(TempFileIdentity::Windows { volume, index });
    }

    #[cfg(not(any(unix, windows)))]
    {
        Ok(TempFileIdentity::Fallback {
            length: metadata.len(),
        })
    }
}

fn ensure_temp_identity(temporary_path: &Path, expected: TempFileIdentity) -> io::Result<()> {
    let observed = path_identity(temporary_path)?;
    if observed == expected {
        Ok(())
    } else {
        Err(io::Error::new(
            ErrorKind::AlreadyExists,
            "CFB temporary path changed while staged output was open",
        ))
    }
}

fn cleanup_owned_temp(temporary_path: &Path, expected: TempFileIdentity) {
    // Best-effort identity checking preserves a known replacement.  The save
    // contract still requires a trusted/private parent because portable,
    // path-based cleanup and replacement cannot close a hostile-directory
    // race between this check and the filesystem operation.
    if path_identity(temporary_path).is_ok_and(|observed| observed == expected) {
        drop(fs::remove_file(temporary_path));
    }
}

struct TemporaryCleanupGuard<'a> {
    path: &'a Path,
    identity: Option<TempFileIdentity>,
    published: bool,
}

impl<'a> TemporaryCleanupGuard<'a> {
    fn new(path: &'a Path) -> Self {
        Self {
            path,
            identity: None,
            published: false,
        }
    }

    fn set_identity(&mut self, identity: TempFileIdentity) {
        self.identity = Some(identity);
    }

    fn mark_published(&mut self) {
        self.published = true;
    }
}

impl Drop for TemporaryCleanupGuard<'_> {
    fn drop(&mut self) {
        if self.published {
            return;
        }

        if let Some(identity) = self.identity {
            cleanup_owned_temp(self.path, identity);
        } else {
            // Either the route does not check identity, or identity
            // acquisition failed before the path could be compared. This
            // exact-name cleanup is correct only under save's trusted/private
            // parent contract, but prevents an orphaned temporary file.
            drop(fs::remove_file(self.path));
        }
    }
}

/// Test doubles for the publication steps.
#[cfg(test)]
pub(crate) mod testing {
    use super::{PublishSteps, atomic_replace, sync_parent};
    use std::fs::File;
    use std::io;
    use std::path::Path;

    /// Real steps whose replacement and parent sync are caller hooks; the
    /// shape of the pre-0761 private test entry points.
    pub(crate) struct HookSteps<R, S> {
        replace: Option<R>,
        sync_parent: Option<S>,
    }

    impl<R, S> HookSteps<R, S> {
        pub(crate) fn new(replace: R, sync_parent: S) -> Self {
            Self {
                replace: Some(replace),
                sync_parent: Some(sync_parent),
            }
        }
    }

    impl<R, S> PublishSteps for HookSteps<R, S>
    where
        R: FnOnce(&Path, &Path) -> io::Result<()>,
        S: FnOnce(&Path) -> io::Result<()>,
    {
        fn sync_file(&mut self, file: &File) -> io::Result<()> {
            file.sync_all()
        }

        fn replace(&mut self, temporary: &Path, destination: &Path) -> io::Result<()> {
            match self.replace.take() {
                Some(replace) => replace(temporary, destination),
                None => Err(io::Error::other("replacement hook called twice")),
            }
        }

        fn sync_parent(&mut self, parent: &Path) -> io::Result<()> {
            match self.sync_parent.take() {
                Some(sync) => sync(parent),
                None => Err(io::Error::other("parent-sync hook called twice")),
            }
        }
    }

    /// Real steps that count every call and can fail a synchronization.
    #[derive(Debug, Default)]
    pub(crate) struct CountingSteps {
        pub(crate) file_syncs: usize,
        pub(crate) replaces: usize,
        pub(crate) parent_syncs: usize,
        pub(crate) fail_file_sync: bool,
        pub(crate) fail_parent_sync: bool,
    }

    impl CountingSteps {
        /// `(file syncs, parent syncs, replacements)`.
        pub(crate) fn counts(&self) -> (usize, usize, usize) {
            (self.file_syncs, self.parent_syncs, self.replaces)
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

        fn replace(&mut self, temporary: &Path, destination: &Path) -> io::Result<()> {
            self.replaces += 1;
            atomic_replace(temporary, destination)
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
