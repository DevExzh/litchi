//! The rooted body story read (record 0371; the migration host's
//! `body_storage` and `body_text`).

use super::{Package, PackageError, PackageResult, storage_rewrite_limits};
use crate::body::Storage;

fn invalid() -> PackageError {
    PackageError::InvalidFormat("the Pages body storage cannot be decoded".to_owned())
}

impl Package {
    /// Read the rooted body story: its complete native text (section breaks,
    /// object-replacement and footnote markers included) and its style and
    /// attribute runs. [`Storage::text`] is the body's native text.
    ///
    /// # Errors
    ///
    /// [`PackageError::InvalidFormat`] when the body storage cannot be
    /// resolved or decoded within the package's limits.
    pub fn body_storage(&self) -> PackageResult<Storage> {
        let (storage, _) =
            super::body_storage_objects::body_storage(self).map_err(|_error| invalid())?;
        let limits =
            storage_rewrite_limits(self.state.source.limits()).map_err(|_error| invalid())?;
        let limits = litchi_iwa_text_wire::Limits::new(
            limits.max_message_bytes().max(storage.payload.len()).max(1),
            limits.max_fields(),
            limits.max_fragments(),
            limits.max_text_bytes(),
        )
        .map_err(|_error| invalid())?;
        litchi_iwa_text_wire::from_bytes(&storage.payload, limits).map_err(|_error| invalid())
    }
}
