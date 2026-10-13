//! Atomic filesystem publication of complete ODF packages.
//!
//! Every packaged ODF owner's `save` and `save_with_durability` route through
//! [`save_bytes`], which stages the artifact in a uniquely named sibling file
//! and publishes it with one replacement through the format-neutral
//! [`litchi_core::atomic`] publisher (record 0226).  A failure before
//! publication leaves an existing destination byte-identical and removes the
//! staging file; a failure after publication is reported as the typed
//! [`litchi_core::Error::Committed`] payload so callers never retry it as an
//! untouched destination.

use std::io::Write;
use std::path::Path;

use litchi_core::Durability;
use litchi_core::atomic::{self, Overwrite};

use crate::error::Error;

/// Atomically replace `path` with `bytes`.
///
/// Symbolic-link and non-file destinations are refused before staging.
/// Existing regular-file permissions are copied to the published file.
///
/// # Errors
///
/// Returns [`Error::Core`] carrying the I/O failure; an `Io` payload means the
/// destination is unchanged, a `Committed` payload means it was replaced but a
/// later durability step failed.
pub fn save_bytes(path: &Path, bytes: &[u8], durability: Durability) -> Result<(), Error> {
    atomic::save(
        path,
        durability,
        Overwrite::Replace,
        |writer| writer.write_all(bytes).map_err(Error::from),
        || Ok(()),
    )
}

#[cfg(all(test, unix))]
mod tests {
    use super::save_bytes;
    use litchi_core::Durability;

    #[test]
    fn replaces_the_destination_and_leaves_no_staging_file() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("out.odt");
        std::fs::write(&path, b"old").unwrap();
        save_bytes(&path, b"new bytes", Durability::Full).unwrap();
        assert_eq!(std::fs::read(&path).unwrap(), b"new bytes");
        assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 1);
    }

    #[test]
    fn a_directory_destination_is_refused_without_staging() {
        let directory = tempfile::tempdir().unwrap();
        let error = save_bytes(directory.path(), b"x", Durability::NoSync).unwrap_err();
        assert!(matches!(
            error,
            crate::Error::Core(litchi_core::Error::Io(_))
        ));
        assert_eq!(std::fs::read_dir(directory.path()).unwrap().count(), 0);
    }
}
