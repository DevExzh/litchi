//! Document properties of `meta.xml` as a scoped transaction (ADR 0034 §7).
//!
//! [`Package::edit_metadata`](crate::Package::edit_metadata) stages an
//! [`Edit`]; [`Edit::commit`] publishes a new [`Package`](crate::Package)
//! and leaves the source untouched;
//! [`Package::apply_metadata`](crate::Package::apply_metadata) replays the
//! [`Patch`] onto the exact source snapshot. The engine is shared by every
//! ODF leaf ([`litchi_odf_common::metadata`], record 0288): a change splices
//! only the touched `office:meta` children of `meta.xml` (or adds a
//! `meta.xml` to a package without one); generator, dates and statistics
//! are never rewritten, so a commit consults no clock. A changed edit of a
//! signed or encrypted package is refused.
//!
//! ```
//! use litchi_odp::Package;
//! use litchi_odp::metadata::Field;
//!
//! # fn main() -> Result<(), litchi_odp::metadata::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_metadata();
//! edit.set(Field::Title, Some("Quarterly report"))?;
//! let commit = edit.commit()?;
//! assert!(commit.diagnostics().changed());
//! let restored = commit.package().apply_metadata(&commit.patch().inverse())?.into_package();
//! assert_eq!(restored.as_bytes(), package.as_bytes());
//! # Ok(())
//! # }
//! ```

/// The family error: typed limits plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::metadata::{Diagnostics, Field};

use litchi_odf_common::OwnedPackage;

/// A detached metadata edit of one source [`Package`](crate::Package).
pub type Edit = litchi_odf_common::metadata::Edit<crate::Package>;
/// The result of one metadata commit.
pub type Commit = litchi_odf_common::metadata::Commit<crate::Package>;
/// An exact-byte metadata patch, applicable only to its source snapshot.
pub type Patch = litchi_odf_common::metadata::Patch<crate::Package>;

impl litchi_odf_common::metadata::Owner for crate::Package {
    fn metadata_source(&self) -> Result<OwnedPackage, Error> {
        Ok(self.owned_package().clone())
    }

    fn adopt_metadata(&self, bytes: Vec<u8>) -> Result<Self, Error> {
        Self::from_vec(bytes).map_err(Error::Core)
    }

    fn metadata_bytes(&self) -> &[u8] {
        self.as_bytes()
    }
}

impl crate::Package {
    /// Start a metadata transaction of this package (ADR 0034 §7).
    #[must_use]
    pub fn edit_metadata(&self) -> Edit {
        Edit::new(self)
    }

    /// Replay a metadata patch onto this exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Commit, Error> {
        patch.apply(self)
    }
}
