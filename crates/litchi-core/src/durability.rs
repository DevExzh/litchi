//! Caller-selected durability of atomic filesystem saves.
//!
//! Every Litchi save to a filesystem path publishes through a sibling
//! temporary file that is renamed over the destination. [`Durability`] is the
//! explicit, per-call choice of which synchronizations that publication
//! performs. It is never read from the environment, a global, or a package's
//! stored preferences: the ordinary `save` methods always use
//! [`Durability::Full`], and a weaker level is available only through the
//! `*_with_durability` methods that take it as an argument (change 0761,
//! authorized by change 0758 decision 5).

/// How much an atomic filesystem save synchronizes before it returns.
///
/// # What every level keeps
///
/// The level changes only which synchronization system calls a save makes.
/// At every level the save:
///
/// - writes byte-identical output;
/// - makes every check its route makes before publishing (for example the
///   OOXML routes' refusal of symbolic-link and non-file destinations, the
///   sequential CFB writer's candidate validation and temporary-file identity
///   check, and a generic overlay source's fingerprint recheck);
/// - stages the complete artifact in a sibling temporary file in the
///   destination's own directory;
/// - replaces the destination with one same-directory rename, so a failure or
///   crash **before** the rename leaves the old destination untouched, and,
///   where the filesystem renames atomically, a concurrent reader of the
///   destination path sees either the complete old file or the complete new
///   file, never a partial one;
/// - removes its temporary file when it fails before the rename;
/// - reports failures with the same typed errors, and never an error for a
///   step the level skipped: only [`Full`](Self::Full) can report that the
///   destination was replaced but its directory could not be synchronized.
///
/// # What each level adds
///
/// | Level | Temporary-file sync | Rename | Parent-directory sync |
/// | --- | --- | --- | --- |
/// | [`Full`](Self::Full) (default) | yes | yes | yes, where supported |
/// | [`FileOnly`](Self::FileOnly) | yes | yes | no |
/// | [`NoSync`](Self::NoSync) | no | yes | no |
///
/// The synchronizations decide what survives an operating-system crash or
/// power loss *after* the save has returned. Exact crash behaviour is
/// filesystem-specific; on an ordinary journaling local filesystem:
///
/// - [`Full`](Self::Full): the destination names the complete new file.
/// - [`FileOnly`](Self::FileOnly): the destination names either what it
///   named before the save (the complete old file, or nothing if it did not
///   exist) or the complete new file; the replacement itself may be lost, and
///   a leftover temporary file may remain beside it.
/// - [`NoSync`](Self::NoSync): nothing is promised. The destination may name
///   what it named before the save, the new file, or, on a filesystem that
///   does not write file data before the rename that exposes it, a truncated
///   or zero-filled file.
///
/// A process crash or kill, as opposed to an operating-system crash, needs no
/// synchronization: once the save has returned, the new file is visible and
/// complete at every level.
///
/// A platform may synchronize more than a level asks for. On Windows, where
/// Litchi performs no separate directory synchronization, the OLE2 writers
/// request `MOVEFILE_WRITE_THROUGH` for the replacement at every level.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Durability {
    /// Synchronize the temporary file, rename it over the destination, then
    /// synchronize the parent directory where the platform supports it. This
    /// is what every ordinary `save` does.
    #[default]
    Full,
    /// Synchronize the temporary file and rename it over the destination;
    /// skip the parent-directory synchronization.
    FileOnly,
    /// Rename the flushed temporary file over the destination without any
    /// synchronization.
    NoSync,
}

impl Durability {
    /// Whether the staged temporary file is synchronized before the rename.
    #[must_use]
    pub const fn syncs_file(self) -> bool {
        matches!(self, Self::Full | Self::FileOnly)
    }

    /// Whether the destination's parent directory is synchronized after the
    /// rename. Only a level that does this can report a failure after the
    /// destination was already replaced.
    #[must_use]
    pub const fn syncs_directory(self) -> bool {
        matches!(self, Self::Full)
    }

    /// Stable, content-free label: `full`, `file-only` or `no-sync`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Full => "full",
            Self::FileOnly => "file-only",
            Self::NoSync => "no-sync",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::Durability;

    #[test]
    fn full_is_the_default_and_syncs_both_file_and_directory() {
        assert_eq!(Durability::default(), Durability::Full);
        assert!(Durability::Full.syncs_file());
        assert!(Durability::Full.syncs_directory());
    }

    #[test]
    fn weaker_levels_drop_only_their_named_synchronizations() {
        assert!(Durability::FileOnly.syncs_file());
        assert!(!Durability::FileOnly.syncs_directory());
        assert!(!Durability::NoSync.syncs_file());
        assert!(!Durability::NoSync.syncs_directory());
    }

    #[test]
    fn labels_are_stable_and_distinct() {
        assert_eq!(Durability::Full.as_str(), "full");
        assert_eq!(Durability::FileOnly.as_str(), "file-only");
        assert_eq!(Durability::NoSync.as_str(), "no-sync");
    }
}
