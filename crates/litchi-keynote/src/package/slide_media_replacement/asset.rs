//! Exact-source replacement of one embedded media asset by its handle
//! (record 0417; the migration host's Keynote `replace_media` and the
//! format-neutral `media::Draft::replace`).
//!
//! The asset is any [`MediaAsset`] that [`Package::media_assets`] lists:
//! image data, movie or audio content, a poster, a slide or master
//! background image. The replacement keeps the asset's data identifier, so
//! every drawable that shares it shows the new bytes. The `Data/` member is
//! rewritten and its `DataInfo` digest and materialized length are patched,
//! as the host does; the three root previews are invalidated, as every
//! Keynote media transaction does.

use std::fmt;
use std::num::NonZeroUsize;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::media::Type as MediaType;
use litchi_iwa_common::media::embedded::AssetId;
use sha1::{Digest, Sha1};

use super::{
    Error, MAX_REPLACEMENT_BYTES, MediaBudget, Refusal, SHA1_BYTES, physical_catalog,
    rewrite_member, verify_zip_locality,
};
use crate::package::Package;
use crate::package::media_assets::{Kind, MediaAsset};

/// Options of one asset replacement.
///
/// The replacement ceiling defaults to 1 GiB (the host media editor's
/// default) and never exceeds the package's per-member byte limit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[must_use]
pub struct EditOptions {
    max_replacement_bytes: NonZeroUsize,
}

impl Default for EditOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl EditOptions {
    /// The default options.
    pub const fn new() -> Self {
        Self {
            max_replacement_bytes: match NonZeroUsize::new(MAX_REPLACEMENT_BYTES) {
                Some(value) => value,
                None => NonZeroUsize::MIN,
            },
        }
    }

    /// Set the largest replacement accepted, in bytes (the host's
    /// `media::Draft::set_max_replacement_bytes`). A larger replacement is
    /// refused with [`Refusal::ReplacementTooLarge`].
    pub const fn max_replacement_bytes(mut self, maximum: NonZeroUsize) -> Self {
        self.max_replacement_bytes = maximum;
        self
    }

    /// The configured replacement ceiling.
    #[must_use]
    pub const fn max_replacement_bytes_value(self) -> NonZeroUsize {
        self.max_replacement_bytes
    }
}

/// One media-asset replacement staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    asset: MediaAsset,
    options: EditOptions,
    before: &'a [u8],
    after: Option<Arc<[u8]>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.as_ref().map(|bytes| bytes.len()))
            .finish_non_exhaustive()
    }
}

fn media_type(kind: Kind) -> MediaType {
    match kind {
        Kind::Image => MediaType::Image,
        Kind::Video => MediaType::Video,
        Kind::Audio => MediaType::Audio,
        Kind::Pdf => MediaType::Pdf,
        Kind::Unknown => MediaType::Unknown,
    }
}

/// The reachable asset `asset` names in `package`, with its current bytes.
fn resolve(package: &Package, asset: AssetId) -> Result<(MediaAsset, &[u8]), Error> {
    let found = package
        .media_assets()
        .map_err(|_| Error::InvalidSource)?
        .into_iter()
        .find(|candidate| candidate.data_identifier() == asset)
        .ok_or(Refusal::AssetNotReachable)?;
    let path = found.package_path().ok_or(Refusal::AssetNotMaterialized)?;
    let bytes = package
        .structure_member_bytes(path)
        .ok_or(Refusal::AssetNotMaterialized)?;
    Ok((found, bytes))
}

impl<'a> Edit<'a> {
    fn new(source: &'a Package, asset: &MediaAsset, options: EditOptions) -> Result<Self, Error> {
        let (asset, before) = resolve(source, asset.data_identifier())?;
        Ok(Self {
            source,
            asset,
            options,
            before,
            after: None,
        })
    }

    /// The selected asset as the source package lists it.
    #[must_use]
    pub const fn asset(&self) -> &MediaAsset {
        &self.asset
    }

    /// The asset's current bytes.
    #[must_use]
    pub const fn before(&self) -> &[u8] {
        self.before
    }

    /// The staged replacement, when one is staged.
    #[must_use]
    pub fn after(&self) -> Option<&[u8]> {
        self.after.as_deref()
    }

    /// Stage new bytes for the asset.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyReplacement`]; [`Refusal::ReplacementTooLarge`] above
    /// the options' ceiling or the package's per-member byte limit;
    /// [`Refusal::ReplacementType`] when the bytes' signature names another
    /// media family than the asset's file name (the host's rule);
    /// [`Error::Allocation`].
    pub fn set(&mut self, replacement: &[u8]) -> Result<&mut Self, Error> {
        if replacement.is_empty() {
            return Err(Refusal::EmptyReplacement.into());
        }
        let length = u64::try_from(replacement.len()).unwrap_or(u64::MAX);
        if replacement.len() > self.options.max_replacement_bytes.get()
            || length > self.source.state.options.archive().max_entry_bytes()
        {
            return Err(Refusal::ReplacementTooLarge.into());
        }
        let declared = media_type(self.asset.kind());
        let detected = MediaType::from_bytes(replacement);
        if declared != MediaType::Unknown && detected != MediaType::Unknown && declared != detected
        {
            return Err(Refusal::ReplacementType.into());
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(replacement.len())
            .map_err(|_error| Error::Allocation {
                amount: replacement.len(),
            })?;
        owned.extend_from_slice(replacement);
        self.after = Some(Arc::from(owned));
        Ok(self)
    }

    /// Publish the replacement: rewrite the member and its `DataInfo`,
    /// reopen the candidate and verify that the asset reads back the new
    /// bytes and that every other member is untouched. Replacing an asset
    /// with its current bytes is an exact no-op.
    ///
    /// # Errors
    ///
    /// [`Refusal::EmptyReplacement`] when nothing is staged,
    /// [`Refusal::UnsupportedSource`] for a source without exact ZIP bytes,
    /// [`Error::Limit`], [`Error::InvalidSource`], [`Error::Verification`].
    pub fn commit(self) -> Result<Commit, Error> {
        let replacement = self
            .after
            .clone()
            .ok_or(Error::Refused(Refusal::EmptyReplacement))?;
        let identifier = self.asset.data_identifier();
        let before_digest: [u8; SHA1_BYTES] = Sha1::digest(self.before).into();
        let after_digest: [u8; SHA1_BYTES] = Sha1::digest(&replacement).into();
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        if self.before == &*replacement {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    asset: identifier,
                    before_digest,
                    after_digest,
                },
                diagnostics: Diagnostics::default(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let digest: [u8; SHA1_BYTES] = self
            .asset
            .digest()
            .try_into()
            .map_err(|_error| Error::InvalidSource)?;
        let path = self
            .asset
            .package_path()
            .ok_or(Refusal::AssetNotMaterialized)?;
        let mut budget = MediaBudget::for_package(self.source)?;
        budget.charge_catalog(catalog)?;
        budget.media_bytes(replacement.len())?;
        let preview_plan =
            super::super::rendering_invalidation::root_preview_deletions(catalog.package())
                .map_err(|_error| Error::InvalidSource)?;
        let (package, deleted_previews) = rewrite_member(
            self.source,
            identifier.get(),
            &digest,
            self.before.len(),
            path,
            &replacement,
            &mut budget,
        )?;
        let (_, written) = resolve(&package, identifier).map_err(|_error| Error::Verification)?;
        if written != &*replacement {
            return Err(Error::Verification);
        }
        let candidate = physical_catalog(&package)?;
        verify_zip_locality(catalog, candidate, path, preview_plan.names())?;
        let target_bytes = candidate.shared_source();
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, target_bytes),
                asset: identifier,
                before_digest,
                after_digest,
            },
            diagnostics: Diagnostics {
                changed: true,
                deleted_previews,
            },
        })
    }
}

/// An exact-source reversible media-asset replacement.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    asset: AssetId,
    before_digest: [u8; SHA1_BYTES],
    after_digest: [u8; SHA1_BYTES],
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("noop", &self.is_noop())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch keeps the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            asset: self.asset,
            before_digest: self.after_digest,
            after_digest: self.before_digest,
        }
    }
}

/// A published media-asset replacement.
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// What one media-asset replacement changed.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    deleted_previews: usize,
}

impl Diagnostics {
    /// Whether the package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many root previews the replacement invalidated.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
}

impl Package {
    /// Start replacing the bytes of one reachable, materialized media asset
    /// that [`Package::media_assets`] lists, with default [`EditOptions`]
    /// (host `replace_media`).
    ///
    /// # Errors
    ///
    /// [`Refusal::AssetNotReachable`], [`Refusal::AssetNotMaterialized`],
    /// [`Error::InvalidSource`].
    pub fn edit_media_asset_data(&self, asset: &MediaAsset) -> Result<Edit<'_>, Error> {
        Edit::new(self, asset, EditOptions::new())
    }

    /// [`Package::edit_media_asset_data`] with explicit options.
    ///
    /// # Errors
    ///
    /// As [`Package::edit_media_asset_data`].
    pub fn edit_media_asset_data_with_options(
        &self,
        asset: &MediaAsset,
        options: EditOptions,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, asset, options)
    }

    /// Replay a media-asset replacement on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on any other package, [`Error::Verification`]
    /// when the target does not read back the patched asset.
    pub fn apply_media_asset_data(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let (_, before) = resolve(self, patch.asset).map_err(|_error| Error::PatchConflict)?;
        if <[u8; SHA1_BYTES]>::from(Sha1::digest(before)) != patch.before_digest {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(|_error| Error::PatchConflict)?;
        candidate.validate().map_err(|_error| Error::Verification)?;
        let (_, after) = resolve(&candidate, patch.asset).map_err(|_error| Error::Verification)?;
        if <[u8; SHA1_BYTES]>::from(Sha1::digest(after)) != patch.after_digest {
            return Err(Error::Verification);
        }
        let deleted_previews =
            super::super::rendering_invalidation::root_preview_deletions(catalog.package())
                .map_err(|_error| Error::InvalidSource)?
                .len()
                .saturating_sub(
                    super::super::rendering_invalidation::root_preview_deletions(
                        physical_catalog(&candidate)?.package(),
                    )
                    .map_err(|_error| Error::Verification)?
                    .len(),
                );
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                deleted_previews,
            },
        })
    }
}
