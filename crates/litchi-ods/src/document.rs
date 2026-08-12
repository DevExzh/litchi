//! Unified, source-bound ODS package transactions and durable patches.

#![deny(
    clippy::cast_possible_truncation,
    clippy::expect_used,
    clippy::panic,
    clippy::unreachable,
    clippy::unwrap_used
)]

use std::{
    collections::{BTreeMap, BTreeSet},
    fmt,
    sync::Arc,
};

use litchi_core::{
    BlobBundle, BlobLimits, CompositionLimits, ConflictSet, DiagnosticFingerprint, Error,
    History as CoreHistory, HistoryLimits, JoinedSubEdits, MergeChoice, Patch as CorePatch,
    PatchError, PatchLimits, PatchOperation, Result, Reversible, ReversibleOperation, SubEdit,
    SubEditConflict, SubEditJoinFailure, ThreeWayMergePlan,
};
use litchi_odf_common::{
    core::{PackageWriter, Profile, XmlSourcePart, XmlSplicePublication},
    package::{Addition, raw_identical_members, rebuild_package},
};
use serde_json::{Value, json};

use crate::package::Package;

pub use crate::advanced::{
    BoundFormControl, CellProperties, CellStyle, CellStyleNode, ChartObject, DataStyleNode,
    Drawing, DrawingFrame, DrawingGeometry, DrawingGeometryKind, DrawingGroup, DrawingPoint,
    DrawingPolygon, DrawingTextBox, EffectiveCellStyle, FormControl, FormControlKind, FormEvent,
    NumberStyleNode, RichFormControl, RichRun, RichText, StyleGraph, TextProperties, TextStyleNode,
};

const FORMAT: &str = "litchi.ods.document";
const MAX_PATH_BYTES: usize = 4_096;
const DOCUMENT_SIGNATURE_PATH: &str = "META-INF/documentsignatures.xml";
const MACRO_SIGNATURE_PATH: &str = "META-INF/macrosignatures.xml";
const CORE_PATHS: [&str; 7] = [
    "mimetype",
    "content.xml",
    "styles.xml",
    "meta.xml",
    "settings.xml",
    "manifest.rdf",
    "META-INF/manifest.xml",
];

/// Finite bounds for one unified package transaction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    package_bytes: usize,
    resources: usize,
    resource_bytes: usize,
    patch: PatchLimits,
    composition: CompositionLimits,
    history: HistoryLimits,
}

impl Limits {
    /// Construct explicit package, transfer, patch, composition, and history bounds.
    #[must_use]
    pub const fn new(
        max_package_bytes: usize,
        max_resources: usize,
        max_resource_bytes: usize,
        patch: PatchLimits,
        composition: CompositionLimits,
        history: HistoryLimits,
    ) -> Self {
        Self {
            package_bytes: max_package_bytes,
            resources: max_resources,
            resource_bytes: max_resource_bytes,
            patch,
            composition,
            history,
        }
    }

    /// Maximum accepted exact package byte length.
    #[must_use]
    pub const fn max_package_bytes(self) -> usize {
        self.package_bytes
    }

    /// Maximum retained auxiliary package members.
    #[must_use]
    pub const fn max_resources(self) -> usize {
        self.resources
    }

    /// Maximum byte length of one transferred resource.
    #[must_use]
    pub const fn max_resource_bytes(self) -> usize {
        self.resource_bytes
    }

    /// Durable semantic-patch bounds.
    #[must_use]
    pub const fn patch(self) -> PatchLimits {
        self.patch
    }

    /// Independent-sub-edit composition bounds.
    #[must_use]
    pub const fn composition(self) -> CompositionLimits {
        self.composition
    }

    /// Undo/redo retention bounds.
    #[must_use]
    pub const fn history(self) -> HistoryLimits {
        self.history
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            package_bytes: 64 * 1024 * 1024,
            resources: 4_096,
            resource_bytes: 32 * 1024 * 1024,
            patch: PatchLimits::new(
                BlobLimits::new(1, 64 * 1024 * 1024, 64 * 1024 * 1024),
                180 * 1024 * 1024,
                256,
                8,
                MAX_PATH_BYTES,
                2 * 1024 * 1024,
            ),
            composition: CompositionLimits::new(64, 512, 8_192, 512),
            history: HistoryLimits::new(32, 512 * 1024 * 1024),
        }
    }
}

/// An immutable ODS package snapshot with exact source lineage.
#[derive(Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    limits: Limits,
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("source_bytes", &self.source.len())
            .field("limits", &self.limits)
            .finish()
    }
}

impl Snapshot {
    /// Parse an exact ODS package with default finite bounds.
    ///
    /// # Errors
    ///
    /// Returns an error when the package exceeds a bound or fails complete facade readback.
    pub fn from_bytes(source: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with(source, Limits::default())
    }

    /// Parse an exact ODS package under explicit finite bounds.
    ///
    /// # Errors
    ///
    /// Returns an error when the package exceeds a bound or fails complete facade readback.
    pub fn from_bytes_with(source: Vec<u8>, limits: Limits) -> Result<Self> {
        Self::from_arc(Arc::from(source), limits)
    }

    fn from_arc(source: Arc<[u8]>, limits: Limits) -> Result<Self> {
        validate_package_size(source.len(), limits)?;
        let bytes = source.as_ref().to_vec();
        let package = Package::from_bytes(bytes)?;
        validate_resource_count(&package, limits)?;
        let _facade = crate::Spreadsheet::from_package(package)?;
        Ok(Self { source, limits })
    }

    /// Borrow the exact package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Bounds retained by this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Content-free diagnostic fingerprint of the exact package.
    #[must_use]
    pub fn fingerprint(&self) -> DiagnosticFingerprint {
        DiagnosticFingerprint::of(&self.source)
    }

    /// Read one bounded auxiliary resource.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is reserved, missing package metadata is malformed, or the
    /// resource exceeds the retained transfer bound.
    pub fn resource(&self, path: &str) -> Result<Option<Resource>> {
        validate_resource_path(path)?;
        let package = Package::from_bytes(self.source.as_ref().to_vec())?;
        if !package.package().has_file(path)? {
            return Ok(None);
        }
        let bytes = package.package().get_file(path)?;
        validate_resource_size(bytes.len(), self.limits)?;
        let media_type = package
            .package()
            .package()?
            .manifest()
            .get_media_type(path)
            .unwrap_or("application/octet-stream")
            .to_string();
        Ok(Some(Resource {
            path: path.to_string(),
            media_type,
            bytes: Arc::from(bytes),
        }))
    }

    /// Report explicit cryptographic write capability/refusal states for this source.
    ///
    /// # Errors
    ///
    /// Returns an error when package security metadata is malformed.
    pub fn security_capabilities(&self) -> Result<SecurityCapabilities> {
        let package = Package::from_bytes(self.source.as_ref().to_vec())?;
        let reader = package.package().package()?;
        let signed =
            reader.has_file(DOCUMENT_SIGNATURE_PATH) || reader.has_file(MACRO_SIGNATURE_PATH);
        let encrypted = reader.manifest().has_encrypted_entries();
        Ok(SecurityCapabilities {
            encrypt: if encrypted {
                SecurityCapability::RefusedUnsupported
            } else {
                SecurityCapability::Available
            },
            sign: SecurityCapability::RefusedUnsupported,
            resign: if signed {
                SecurityCapability::RefusedUnsupported
            } else {
                SecurityCapability::NotApplicable
            },
            decrypt: if encrypted {
                SecurityCapability::RefusedCredentialsRequired
            } else {
                SecurityCapability::NotApplicable
            },
            reencrypt: if encrypted {
                SecurityCapability::RefusedUnsupported
            } else {
                SecurityCapability::NotApplicable
            },
        })
    }

    /// Resolve inherited cell-style properties across `styles.xml` and content automatic styles.
    ///
    /// Content automatic styles take precedence over same-named common styles. Unsupported
    /// property vocabulary remains inert while the supported cell/text projection is resolved.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/cyclic style or malformed package XML.
    pub fn effective_cell_style(&self, name: &str) -> Result<EffectiveCellStyle> {
        crate::advanced::resolve_package_cell_style(self.source.as_ref(), name)
    }

    /// Report whether terminal password encryption can publish this exact snapshot.
    ///
    /// The report is machine-readable and accounts for encrypted input plus the caller's stale
    /// signature policy.
    ///
    /// # Errors
    ///
    /// Returns an error when package security metadata is malformed.
    pub fn encryption_write_capability(
        &self,
        signatures: SignatureWritePolicy,
    ) -> Result<EncryptionWriteCapability> {
        let package = Package::from_bytes(self.source.as_ref().to_vec())?;
        let reader = package.package().package()?;
        if reader.manifest().has_encrypted_entries() {
            return Ok(EncryptionWriteCapability::Refused(
                EncryptionWriteRefusal::EncryptedSourceRequiresCredentialSnapshot,
            ));
        }
        if (reader.has_file(DOCUMENT_SIGNATURE_PATH) || reader.has_file(MACRO_SIGNATURE_PATH))
            && matches!(signatures, SignatureWritePolicy::Refuse)
        {
            return Ok(EncryptionWriteCapability::Refused(
                EncryptionWriteRefusal::SignedSourceRequiresStripPolicy,
            ));
        }
        Ok(EncryptionWriteCapability::Available)
    }

    /// Publish this exact plaintext snapshot as a bounded password-encrypted ODS archive.
    ///
    /// This terminal operation does not create an editable encrypted `Snapshot`. Existing
    /// signatures are refused by default because encryption invalidates their package bytes.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty/oversized password, encrypted or signed input, package
    /// bounds, or encrypted archive publication failure.
    pub fn encrypt_to_bytes(&self, password: &str, profile: EncryptionProfile) -> Result<Vec<u8>> {
        self.encrypt_to_bytes_with_signatures(password, profile, SignatureWritePolicy::Refuse)
    }

    /// Publish terminal encrypted bytes with an explicit stale-signature policy.
    ///
    /// `StripInvalidated` deliberately omits existing document and macro signature containers.
    ///
    /// # Errors
    ///
    /// Returns an error for an empty/oversized password, a machine-readable capability refusal,
    /// package bounds, or encrypted archive publication failure.
    pub fn encrypt_to_bytes_with_signatures(
        &self,
        password: &str,
        profile: EncryptionProfile,
        signatures: SignatureWritePolicy,
    ) -> Result<Vec<u8>> {
        if password.is_empty() || password.len() > 4_096 {
            return invalid("ODS encryption password must contain 1 through 4096 UTF-8 bytes");
        }
        match self.encryption_write_capability(signatures)? {
            EncryptionWriteCapability::Available => {},
            EncryptionWriteCapability::Refused(
                EncryptionWriteRefusal::EncryptedSourceRequiresCredentialSnapshot,
            ) => {
                return invalid(
                    "ODS terminal encryption refuses an already-encrypted source without a credential-retaining snapshot",
                );
            },
            EncryptionWriteCapability::Refused(
                EncryptionWriteRefusal::SignedSourceRequiresStripPolicy,
            ) => {
                return invalid(
                    "ODS terminal encryption requires explicit stale-signature stripping",
                );
            },
        }
        let package = Package::from_bytes(self.source.as_ref().to_vec())?;
        let mut writer = PackageWriter::new_bounded(self.limits.package_bytes);
        writer.set_mimetype(&package.package().mimetype()?)?;
        writer.set_encryption(password.to_string(), profile.shared())?;
        for path in ["content.xml", "styles.xml", "meta.xml"] {
            if package.package().has_file(path)? {
                let part = XmlSourcePart::load(package.package(), path)?;
                writer.add_spliced_xml(XmlSplicePublication::new(part))?;
            }
        }
        writer.copy_auxiliary_files_from(package.package())?;
        writer.finish_to_bounded_bytes()
    }

    /// Begin one clone-staged unified package transaction.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            before: self.clone(),
            candidate: self.source.as_ref().to_vec(),
            steps: Vec::new(),
            spliced_parts: BTreeSet::new(),
        }
    }
}

/// One bounded inert auxiliary package resource.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    path: String,
    media_type: String,
    bytes: Arc<[u8]>,
}

impl Resource {
    /// Construct one detached resource value.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe/reserved path, empty media type, or noncompact XML payload.
    pub fn new(
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Result<Self> {
        let path = path.into();
        let media_type = media_type.into();
        validate_resource_path(&path)?;
        validate_media_type(&media_type)?;
        let bytes = bytes.into();
        validate_xml_resource(&media_type, &bytes)?;
        Ok(Self {
            path,
            media_type,
            bytes: Arc::from(bytes),
        })
    }

    /// Package-relative resource path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Declared manifest media type.
    #[must_use]
    pub fn media_type(&self) -> &str {
        &self.media_type
    }

    /// Exact inert resource bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }
}

/// Explicit resource-collision policy for package and cross-document transfer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Collision {
    /// Refuse any existing destination path.
    Reject,
    /// Reuse an existing byte-and-media equivalent resource, otherwise refuse.
    ReuseEquivalent,
    /// Explicitly replace an existing resource at the same safe path.
    Replace,
}

/// Outcome of one bounded resource transfer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum TransferDisposition {
    /// New bytes were staged in the destination package.
    Added,
    /// Existing exactly equivalent bytes were reused.
    Reused,
    /// Existing bytes were explicitly replaced.
    Replaced,
}

/// Explicit disposition for signatures invalidated by a changed publication.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum SignatureWritePolicy {
    /// Reject changes to a signed source.
    #[default]
    Refuse,
    /// Deliberately omit stale document and macro signature containers.
    StripInvalidated,
}

/// Explicit encrypted-source write policy.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionWritePolicy {
    /// Reject changes to encrypted source members; no password is inferred or retained.
    #[default]
    Refuse,
    /// Require password-backed decrypt/re-encrypt publication.
    ///
    /// The current unified ODS root exposes this intent explicitly but refuses it because it has
    /// no credential-bearing writer capability.
    DecryptAndReencrypt,
}

/// Bounded password-encryption profile for terminal ODS publication.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionProfile {
    /// AES-256-CBC, SHA-256, and PBKDF2 for broad ODF compatibility.
    #[default]
    Compatible,
    /// AES-256-GCM, SHA-256, and Argon2id for authenticated encryption.
    Authenticated,
}

impl EncryptionProfile {
    fn shared(self) -> Profile {
        match self {
            Self::Compatible => Profile::compatible(),
            Self::Authenticated => Profile::authenticated(),
        }
    }
}

/// Machine-readable reason terminal password encryption cannot proceed.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionWriteRefusal {
    /// The source already has encrypted members and requires a credential-retaining edit owner.
    EncryptedSourceRequiresCredentialSnapshot,
    /// Existing signatures require explicit stale-signature stripping.
    SignedSourceRequiresStripPolicy,
}

/// Availability of terminal password encryption for one exact snapshot and signature policy.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum EncryptionWriteCapability {
    Available,
    Refused(EncryptionWriteRefusal),
}

/// Caller-selected security disposition for changed package publication.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct SecurityWritePolicy {
    pub signatures: SignatureWritePolicy,
    pub encryption: EncryptionWritePolicy,
}

/// Explicit availability state for a cryptographic package operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum SecurityCapability {
    /// The operation is available under the documented bounded API.
    Available,
    /// The operation is irrelevant to the current source state.
    NotApplicable,
    /// The ODS unified root deliberately does not implement the operation.
    RefusedUnsupported,
    /// The operation would require caller credentials that this root never infers.
    RefusedCredentialsRequired,
}

/// Encrypt/sign/resign/decrypt/re-encrypt capability states for one exact snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SecurityCapabilities {
    pub encrypt: SecurityCapability,
    pub sign: SecurityCapability,
    pub resign: SecurityCapability,
    pub decrypt: SecurityCapability,
    pub reencrypt: SecurityCapability,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct Step {
    op: String,
    target: String,
    effects: Vec<String>,
}

/// A unified clone-staged ODS package edit.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    candidate: Vec<u8>,
    steps: Vec<Step>,
    spliced_parts: BTreeSet<String>,
}

impl Edit {
    /// Borrow the exact current package candidate.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.candidate
    }

    /// Stage worksheet structure, cell value, formula, and direct-style changes.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or worksheet publication fails.
    pub fn worksheets<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::worksheet::Edit) -> Result<()>,
    {
        let source = Arc::new(std::mem::take(&mut self.candidate));
        let outcome: Result<Option<(Arc<Vec<u8>>, bool)>> = (|| {
            let snapshot = crate::worksheet::Snapshot::from_shared_bytes(Arc::clone(&source))?;
            let mut edit = snapshot.edit();
            update(&mut edit)?;
            let commit = edit.commit()?;
            if !commit.changed() {
                return Ok(None);
            }
            let provenance_spliced = commit.content_provenance_spliced();
            Ok(Some((
                commit.into_snapshot().into_shared_bytes(),
                provenance_spliced,
            )))
        })();
        self.candidate = take_shared_bytes(source);
        let Some((candidate, provenance_spliced)) = outcome? else {
            return Ok(());
        };
        if provenance_spliced {
            self.stage_spliced_shared("worksheet.edit", "worksheets", candidate)
        } else {
            self.stage_shared("worksheet.edit", "worksheets", candidate)
        }
    }

    /// Stage ordered named-range and named-expression changes.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or definition publication fails.
    pub fn definitions<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::definitions::Edit) -> Result<()>,
    {
        let snapshot = crate::definitions::Snapshot::from_bytes(self.candidate.clone())?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        self.stage(
            "definition.edit",
            "definitions",
            commit.snapshot().as_bytes().to_vec(),
        )
    }

    /// Stage cell-annotation CRUD.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, annotation publication, or package rebuild fails.
    pub fn annotations<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::annotations::Transaction) -> Result<()>,
    {
        let package = Package::from_bytes(self.candidate.clone())?;
        let snapshot = crate::annotations::Snapshot::parse(package.content_xml())?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        let bytes = if commit.changed() {
            package
                .replace_content_xml(commit.content_xml())?
                .into_bytes()
        } else {
            self.candidate.clone()
        };
        self.stage("annotation.edit", "annotations", bytes)
    }

    /// Stage bounded content-validation definition CRUD.
    ///
    /// Cell bindings remain compact source-owned runs. Removal and clear operations are refused
    /// while any such run retains a definition, and renames are refused because this API does not
    /// perform an atomic binding rewrite. Changed publication is security-checked only by
    /// [`Self::commit`] or [`Self::commit_with_security`].
    ///
    /// # Errors
    ///
    /// Returns an error for opaque ownership, broken binding closure, limits, allocation,
    /// provenance-splice failure, or typed package readback failure.
    pub fn content_validations<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(
            &mut crate::content_validation::Transaction<'_, '_>,
        ) -> crate::content_validation::Result<()>,
    {
        let source = Arc::new(std::mem::take(&mut self.candidate));
        let outcome: Result<Option<Vec<u8>>> = (|| {
            let package = Package::from_shared_bytes(Arc::clone(&source))?;
            let snapshot = crate::content_validation::Snapshot::parse(package.content_xml())
                .map_err(content_validation_error)?;
            let mut edit = snapshot.edit().map_err(content_validation_error)?;
            update(&mut edit).map_err(content_validation_error)?;
            let commit = edit.commit().map_err(content_validation_error)?;
            if !commit.changed() {
                return Ok(None);
            }
            crate::content_validation::publish_package_commit(&package, &commit)
                .map(Package::into_bytes)
                .map(Some)
        })();
        self.candidate = take_shared_bytes(source);
        let Some(candidate) = outcome? else {
            return Ok(());
        };
        self.stage_spliced("content-validation.edit", "content-validations", candidate)
    }

    /// Stage a reversible content-validation patch against its exact XML source.
    ///
    /// # Errors
    ///
    /// Returns an error for stale/foreign lineage, invalid closure, splice publication, or
    /// complete package readback failure.
    pub fn apply_content_validation_patch(
        &mut self,
        patch: &crate::content_validation::Patch,
    ) -> Result<()> {
        let source = Arc::new(std::mem::take(&mut self.candidate));
        let outcome: Result<Option<Vec<u8>>> = (|| {
            let package = Package::from_shared_bytes(Arc::clone(&source))?;
            let snapshot = crate::content_validation::Snapshot::parse(package.content_xml())
                .map_err(content_validation_error)?;
            let commit = patch.apply(&snapshot).map_err(content_validation_error)?;
            if !commit.changed() {
                return Ok(None);
            }
            crate::content_validation::publish_package_commit(&package, &commit)
                .map(Package::into_bytes)
                .map(Some)
        })();
        self.candidate = take_shared_bytes(source);
        let Some(candidate) = outcome? else {
            return Ok(());
        };
        self.stage_spliced("content-validation.patch", "content-validations", candidate)
    }

    /// Stage inert RDF graph and triple CRUD.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or RDF package publication fails.
    pub fn rdf<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::metadata_graphs::Edit) -> Result<()>,
    {
        let snapshot = crate::metadata_graphs::Snapshot::from_bytes(self.candidate.clone())?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit();
        self.stage(
            "rdf.edit",
            "metadata-graphs",
            commit.snapshot().as_bytes().to_vec(),
        )
    }

    /// Stage document, sheet, and direct cell-style protection metadata.
    ///
    /// This does not grant an unlock capability. A changed transaction whose accepted source is
    /// already protected is refused at the unified commit boundary.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, protection publication, or package rebuild fails.
    pub fn protection<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::protection::Transaction) -> Result<()>,
    {
        let package = Package::from_bytes(self.candidate.clone())?;
        let snapshot =
            crate::protection::Snapshot::parse(package.content_xml(), package.styles_xml())?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        let bytes = if commit.changed() {
            package
                .replace_content_xml(commit.content_xml())?
                .into_bytes()
        } else {
            self.candidate.clone()
        };
        self.stage("protection.edit", "protection", bytes)
    }

    /// Stage `DataPilot` table CRUD.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure or `DataPilot` package publication fails.
    pub fn data_pilot<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::data_pilot::Edit) -> Result<()>,
    {
        let snapshot = crate::data_pilot::Snapshot::from_bytes(self.candidate.clone())?;
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        let commit = edit.commit()?;
        self.stage(
            "data-pilot.edit",
            "data-pilot",
            commit.snapshot().as_bytes().to_vec(),
        )
    }

    /// Stage tracked-change graph CRUD and acceptance metadata.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, tracked-change publication, or package rebuild fails.
    pub fn tracked_changes<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::tracked_changes::Transaction) -> Result<()>,
    {
        let package = Package::from_bytes(self.candidate.clone())?;
        let snapshot = crate::tracked_changes::Snapshot::parse(package.content_xml())?;
        let mut edit = crate::tracked_changes::Transaction::new(snapshot)?;
        update(&mut edit)?;
        let commit = edit.commit()?;
        let bytes = if commit.changed() {
            package
                .replace_content_xml(commit.content_xml())?
                .into_bytes()
        } else {
            self.candidate.clone()
        };
        self.stage("tracked-change.edit", "tracked-changes", bytes)
    }

    /// Stage embedded-chart part replacement.
    ///
    /// # Errors
    ///
    /// Returns an error when the closure, compact chart XML, or chart publication is invalid.
    pub fn charts<F>(&mut self, update: F) -> Result<()>
    where
        F: FnOnce(&mut crate::charts::Edit) -> Result<()>,
    {
        let snapshot = crate::charts::Snapshot::from_bytes(self.candidate.clone())?;
        let before = snapshot.charts().to_vec();
        let mut edit = snapshot.edit();
        update(&mut edit)?;
        for (index, chart) in edit.charts().iter().enumerate() {
            if before.get(index) != Some(chart) {
                litchi_odf_common::compact_xml::validate(chart.content_xml().as_bytes())
                    .map_err(Error::from)?;
            }
        }
        let commit = edit.commit()?;
        self.stage(
            "chart.edit",
            "charts",
            commit.snapshot().as_bytes().to_vec(),
        )
    }

    /// Replace one existing non-repeated cell body with checked rich paragraphs and inline runs.
    ///
    /// The splice is provenance-bound to the exact `content.xml`; untouched producer formatting
    /// stays byte-exact while the replacement cell is compact authored XML.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/repeated cell, invalid rich text, unsafe hyperlink, package
    /// bound, compactness failure, or complete readback failure.
    pub fn set_rich_cell_text(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        rich: &RichText,
    ) -> Result<()> {
        let bytes = crate::advanced::set_rich_text(
            &self.candidate,
            sheet,
            row,
            column,
            rich,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("cell.rich-text", &format!("{sheet}!R{row}C{column}"), bytes)
    }

    /// Set one inert formula through a fine-grained provenance splice.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/repeated cell, invalid formula, package bound, compactness,
    /// or complete readback failure.
    pub fn set_cell_formula(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        formula: &str,
    ) -> Result<()> {
        let bytes = crate::advanced::set_cell_formula(
            &self.candidate,
            sheet,
            row,
            column,
            formula,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("cell.formula", &format!("{sheet}!R{row}C{column}"), bytes)
    }

    /// Set a direct cell style reference through a fine-grained provenance splice.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/repeated cell, invalid style name, package bound,
    /// compactness, or complete readback failure.
    pub fn set_cell_style(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        style_name: &str,
    ) -> Result<()> {
        let bytes = crate::advanced::set_cell_style(
            &self.candidate,
            sheet,
            row,
            column,
            style_name,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("cell.style", &format!("{sheet}!R{row}C{column}"), bytes)
    }

    /// Append one compact row without rewriting existing Calc XML.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown sheet, invalid row, package bound, or splice failure.
    pub fn append_row(&mut self, sheet: &str, row: &crate::Row) -> Result<()> {
        let bytes = crate::advanced::append_row(
            &self.candidate,
            sheet,
            row,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("row.insert", sheet, bytes)
    }

    /// Remove one checked physical non-repeated row.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown sheet, out-of-range/repeated row, or splice failure.
    pub fn remove_row(&mut self, sheet: &str, physical_position: usize) -> Result<()> {
        let bytes = crate::advanced::remove_row(
            &self.candidate,
            sheet,
            physical_position,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("row.remove", &format!("{sheet}#{physical_position}"), bytes)
    }

    /// Append one compact structural column declaration.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown sheet, invalid column, package bound, or splice failure.
    pub fn append_column(
        &mut self,
        sheet: &str,
        column: &crate::model::structure::Column,
    ) -> Result<()> {
        let bytes = crate::advanced::append_column(
            &self.candidate,
            sheet,
            column,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("column.insert", sheet, bytes)
    }

    /// Remove one checked physical non-repeated column declaration.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown sheet, out-of-range/repeated column, or splice failure.
    pub fn remove_column(&mut self, sheet: &str, physical_position: usize) -> Result<()> {
        let bytes = crate::advanced::remove_column(
            &self.candidate,
            sheet,
            physical_position,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced(
            "column.remove",
            &format!("{sheet}#{physical_position}"),
            bytes,
        )
    }

    /// Append one checked compact worksheet.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid sheet, package bound, or provenance splice failure.
    pub fn append_sheet(&mut self, sheet: &crate::Sheet) -> Result<()> {
        let bytes = crate::advanced::append_sheet(
            &self.candidate,
            sheet,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("sheet.insert", &sheet.name, bytes)
    }

    /// Remove one worksheet by exact name.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/ambiguous sheet or provenance splice failure.
    pub fn remove_sheet(&mut self, sheet: &str) -> Result<()> {
        let bytes = crate::advanced::remove_sheet(
            &self.candidate,
            sheet,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("sheet.remove", sheet, bytes)
    }

    /// Add one compact automatic table-cell style to `content.xml`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid/duplicate style, color, package bound, or splice failure.
    pub fn put_cell_style(&mut self, style: &CellStyle) -> Result<()> {
        let bytes = crate::advanced::put_cell_style(
            &self.candidate,
            style,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("style.put", &style.name, bytes)
    }

    /// Add a dependency-checked automatic number/text/cell style graph.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate names, unresolved parent/data-style references, invalid
    /// properties, existing-name collisions, bounds, or splice failure.
    pub fn put_style_graph(&mut self, graph: &StyleGraph) -> Result<()> {
        let bytes = crate::advanced::put_style_graph(
            &self.candidate,
            graph,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("style-graph.put", "automatic-styles", bytes)
    }

    /// Replace every same-family automatic style named by a closed graph.
    ///
    /// Unrelated automatic styles and formatting outside checked splice ranges remain exact.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid graph, absent/duplicate/mismatched nodes, bounds, or splice
    /// failure.
    pub fn replace_style_graph(&mut self, graph: &StyleGraph) -> Result<()> {
        let bytes = crate::advanced::replace_style_graph(
            &self.candidate,
            graph,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("style-graph.replace", "automatic-styles", bytes)
    }

    /// Remove exact automatic-style names when no retained XML node references them.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicates, absent styles, retained references, bounds, or splice
    /// failure.
    pub fn remove_automatic_styles(&mut self, names: &[String]) -> Result<()> {
        let bytes = crate::advanced::remove_automatic_styles(
            &self.candidate,
            names,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("style-graph.remove", "automatic-styles", bytes)
    }

    /// Atomically add a style dependency graph and apply rich text to one cell.
    ///
    /// # Errors
    ///
    /// Returns an error when either the graph or rich cell publication fails; no partial style
    /// nodes are retained.
    pub fn set_rich_cell_text_with_styles(
        &mut self,
        sheet: &str,
        row: usize,
        column: usize,
        rich: &RichText,
        styles: &StyleGraph,
    ) -> Result<()> {
        let declared = styles
            .text_styles
            .iter()
            .map(|style| style.name.as_str())
            .collect::<BTreeSet<_>>();
        if rich.paragraphs().iter().flatten().any(|run| {
            matches!(
                run,
                RichRun::Span { style_name, .. } if !declared.contains(style_name.as_str())
            )
        }) {
            return invalid("ODS rich cell style dependency is unresolved");
        }
        let mut candidate = self.clone();
        candidate.put_style_graph(styles)?;
        candidate.set_rich_cell_text(sheet, row, column, rich)?;
        *self = candidate;
        Ok(())
    }

    /// Replace or remove the typed conditional-format catalog of one sheet.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid rules/ranges, an unknown sheet, bounds, or splice failure.
    pub fn set_conditional_formats(
        &mut self,
        sheet: &str,
        formats: &[crate::model::conditional_format::Format],
    ) -> Result<()> {
        let bytes = crate::advanced::set_conditional_formats(
            &self.candidate,
            sheet,
            formats,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("conditional-format.edit", sheet, bytes)
    }

    /// Replace or remove the typed sparkline-group catalog of one sheet.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid groups/ranges, an unknown sheet, bounds, or splice failure.
    pub fn set_sparkline_groups(
        &mut self,
        sheet: &str,
        groups: &[crate::model::sparkline::Group],
    ) -> Result<()> {
        let bytes = crate::advanced::set_sparkline_groups(
            &self.candidate,
            sheet,
            groups,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("sparkline.edit", sheet, bytes)
    }

    /// Replace or remove the inert spreadsheet form-button catalog.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid controls, duplicate containers, bounds, or splice failure.
    pub fn set_form_controls(&mut self, controls: &[FormControl]) -> Result<()> {
        let bytes = crate::advanced::set_forms(
            &self.candidate,
            controls,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("form.edit", "forms", bytes)
    }

    /// Replace bound form controls and close every declared image dependency atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for duplicate controls, invalid bindings, an absent/unprovided image,
    /// collision, bounds, or splice failure.
    pub fn set_bound_form_controls_with_resources(
        &mut self,
        controls: &[BoundFormControl],
        resources: Vec<Resource>,
        collision: Collision,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let supplied = resources
            .iter()
            .map(Resource::path)
            .collect::<BTreeSet<_>>();
        let current = Package::from_bytes(candidate.candidate.clone())?;
        for control in controls {
            if let Some(path) = &control.image_path
                && !supplied.contains(path.as_str())
                && !current.package().has_file(path)?
            {
                return invalid(format!("ODS form image dependency '{path}' is unresolved"));
            }
        }
        for resource in resources {
            let _disposition = candidate.put_resource(resource, collision)?;
        }
        let bytes = crate::advanced::set_bound_forms(
            &candidate.candidate,
            controls,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("form.bindings", "forms", bytes)?;
        *self = candidate;
        Ok(())
    }

    /// Replace typed inert controls, event bindings, and image dependencies atomically.
    ///
    /// Events and macro targets are retained as data and are never executed.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid controls/events/bindings, unresolved images, collisions,
    /// bounds, or splice failure.
    pub fn set_rich_form_controls_with_resources(
        &mut self,
        controls: &[RichFormControl],
        resources: Vec<Resource>,
        collision: Collision,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let supplied = resources
            .iter()
            .map(Resource::path)
            .collect::<BTreeSet<_>>();
        let current = Package::from_bytes(candidate.candidate.clone())?;
        for control in controls {
            if let Some(path) = &control.image_path
                && !supplied.contains(path.as_str())
                && !current.package().has_file(path)?
            {
                return invalid(format!("ODS form image dependency '{path}' is unresolved"));
            }
        }
        for resource in resources {
            let _disposition = candidate.put_resource(resource, collision)?;
        }
        let bytes = crate::advanced::set_rich_forms(
            &candidate.candidate,
            controls,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("form.rich", "forms", bytes)?;
        *self = candidate;
        Ok(())
    }

    /// Atomically add a drawing frame and its exact package dependency.
    ///
    /// The detached resource path must equal the drawing reference. Both staged changes roll back
    /// if either collision policy or XML publication fails.
    ///
    /// # Errors
    ///
    /// Returns an error for dependency mismatch, collision, invalid drawing, bounds, or splice
    /// failure.
    pub fn put_drawing_with_resource(
        &mut self,
        sheet: &str,
        drawing: &Drawing,
        resource: Resource,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        if drawing.resource_path != resource.path {
            return invalid("ODS drawing dependency path differs from its resource");
        }
        let mut candidate = self.clone();
        let disposition = candidate.put_resource(resource, collision)?;
        let bytes = crate::advanced::put_drawing(
            &candidate.candidate,
            sheet,
            drawing,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("drawing.put", &drawing.name, bytes)?;
        *self = candidate;
        Ok(disposition)
    }

    /// Atomically add a positioned drawing frame and its package dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid geometry/anchor, dependency mismatch, collision, bounds, or
    /// splice failure.
    pub fn put_drawing_frame_with_resource(
        &mut self,
        sheet: &str,
        frame: &DrawingFrame,
        resource: Resource,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        if frame.resource_path != resource.path {
            return invalid("ODS drawing-frame dependency path differs from its resource");
        }
        let mut candidate = self.clone();
        let disposition = candidate.put_resource(resource, collision)?;
        let bytes = crate::advanced::put_drawing_frame(
            &candidate.candidate,
            sheet,
            frame,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("drawing-frame.put", &frame.name, bytes)?;
        *self = candidate;
        Ok(disposition)
    }

    /// Add a bounded drawing group containing rich text frames.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid identity, duplicate names, geometry, rich text, bounds, or
    /// splice failure.
    pub fn put_drawing_group(&mut self, sheet: &str, group: &DrawingGroup) -> Result<()> {
        let bytes = crate::advanced::put_drawing_group(
            &self.candidate,
            sheet,
            group,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("drawing-group.put", &group.name, bytes)
    }

    /// Add a finite positioned rectangle, ellipse, or line with optional rich text.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid identity, duplicate names, geometry, rich text, bounds, or
    /// splice failure.
    pub fn put_drawing_geometry(&mut self, sheet: &str, geometry: &DrawingGeometry) -> Result<()> {
        let bytes = crate::advanced::put_drawing_geometry(
            &self.candidate,
            sheet,
            geometry,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("drawing-geometry.put", &geometry.name, bytes)
    }

    /// Add a bounded positioned polygon or polyline with optional rich text.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid identity, view-box coordinates, geometry, rich text, bounds,
    /// duplicate names, or splice failure.
    pub fn put_drawing_polygon(&mut self, sheet: &str, polygon: &DrawingPolygon) -> Result<()> {
        let bytes = crate::advanced::put_drawing_polygon(
            &self.candidate,
            sheet,
            polygon,
            self.before.limits.package_bytes,
        )?;
        self.stage_spliced("drawing-polygon.put", &polygon.name, bytes)
    }

    /// Atomically add a package-backed chart drawing and compact chart content dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for an unsafe object path, noncompact/malformed chart XML, collision,
    /// bounds, drawing publication, or full chart readback failure.
    pub fn put_chart_object(
        &mut self,
        sheet: &str,
        chart: &ChartObject,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        let content_path = format!("{}/content.xml", chart.object_path);
        let resource = Resource::new(&content_path, "text/xml", chart.content_xml.as_bytes())?;
        let mut candidate = self.clone();
        let disposition = candidate.put_resource(resource, collision)?;
        let package = Package::from_bytes(candidate.candidate.clone())?;
        candidate.candidate = rebuild_package(
            package.package(),
            package.content_xml(),
            Vec::new(),
            vec![(
                format!("{}/", chart.object_path),
                litchi_odf_common::constants::ODF_CHART.to_string(),
            )],
            Vec::<String>::new(),
            Vec::<String>::new(),
        )?;
        let bytes = crate::advanced::put_chart_object(
            &candidate.candidate,
            sheet,
            chart,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("chart-object.put", &chart.name, bytes)?;
        let chart_snapshot = crate::charts::Snapshot::from_bytes(candidate.candidate.clone())?;
        if chart_snapshot
            .charts()
            .iter()
            .all(|value| value.name() != Some(chart.name.as_str()))
        {
            return invalid("ODS chart dependency failed typed readback");
        }
        *self = candidate;
        Ok(disposition)
    }

    /// Serialize a typed chart definition and add its complete package-backed occurrence.
    ///
    /// # Errors
    ///
    /// Returns an error when chart serialization, compactness, collision, dependency publication,
    /// or typed readback fails.
    pub fn put_chart_definition(
        &mut self,
        sheet: &str,
        drawing_name: &str,
        object_path: &str,
        definition: &crate::charts::Definition,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        let part = crate::charts::Part::from_definition(definition)?;
        self.put_chart_object(
            sheet,
            &ChartObject {
                name: drawing_name.to_string(),
                object_path: object_path.to_string(),
                content_xml: part.xml().to_string(),
            },
            collision,
        )
    }

    /// Transfer one compact package-backed chart and every member below its object root.
    ///
    /// Relative chart data references remain valid because dependency paths below the source root
    /// are reproduced below the destination root. The transfer is clone-staged and atomic.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent content part, unsafe roots, noncompact XML dependencies,
    /// collision, bounds, drawing publication, or typed chart readback failure.
    pub fn transfer_chart_dependencies(
        &mut self,
        source: &Snapshot,
        source_object_path: &str,
        sheet: &str,
        drawing_name: &str,
        destination_object_path: &str,
        collision: Collision,
    ) -> Result<Vec<TransferDisposition>> {
        let source_content_path = format!("{source_object_path}/content.xml");
        let destination_content_path = format!("{destination_object_path}/content.xml");
        validate_resource_path(&source_content_path)?;
        validate_resource_path(&destination_content_path)?;
        let source_package = Package::from_bytes(source.source.as_ref().to_vec())?;
        let content = source_package
            .package()
            .has_file(&source_content_path)?
            .then(|| source_package.package().get_file(&source_content_path))
            .transpose()?
            .ok_or_else(|| invalid_error("ODS chart source content.xml was not found"))?;
        let content_xml = String::from_utf8(content)
            .map_err(|_error| invalid_error("ODS chart source content.xml is not UTF-8"))?;
        let source_prefix = format!("{source_object_path}/");
        let destination_prefix = format!("{destination_object_path}/");
        let manifest = source_package.package().package()?;
        let mut candidate = self.clone();
        let mut dispositions = Vec::new();
        for path in source_package.package().files()? {
            let Some(relative) = path.strip_prefix(&source_prefix) else {
                continue;
            };
            if relative.is_empty() || relative == "content.xml" || path.ends_with('/') {
                continue;
            }
            let destination = format!("{destination_prefix}{relative}");
            let media_type = manifest
                .manifest()
                .get_media_type(&path)
                .unwrap_or("application/octet-stream");
            let resource = Resource::new(
                destination,
                media_type,
                source_package.package().get_file(&path)?,
            )?;
            dispositions.push(candidate.put_resource(resource, collision)?);
        }
        dispositions.push(candidate.put_chart_object(
            sheet,
            &ChartObject {
                name: drawing_name.to_string(),
                object_path: destination_object_path.to_string(),
                content_xml,
            },
            collision,
        )?);
        *self = candidate;
        Ok(dispositions)
    }

    /// Transfer one resource dependency and author a destination drawing reference atomically.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing source, unsafe destination, collision, bounds, or splice
    /// failure.
    pub fn transfer_drawing(
        &mut self,
        source: &Snapshot,
        source_path: &str,
        sheet: &str,
        drawing_name: &str,
        destination_path: &str,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        let source_resource = source.resource(source_path)?.ok_or_else(|| {
            invalid_error(format!(
                "ODS drawing resource '{source_path}' was not found"
            ))
        })?;
        let resource = Resource::new(
            destination_path,
            source_resource.media_type,
            source_resource.bytes.as_ref().to_vec(),
        )?;
        self.put_drawing_with_resource(
            sheet,
            &Drawing {
                name: drawing_name.to_string(),
                resource_path: destination_path.to_string(),
            },
            resource,
            collision,
        )
    }

    /// Atomically remove a drawing reference and its now-unreferenced resource dependency.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing drawing/resource, retained reference, bounds, or splice
    /// failure.
    pub fn remove_drawing_with_resource(
        &mut self,
        sheet: &str,
        drawing_name: &str,
        resource_path: &str,
    ) -> Result<()> {
        let mut candidate = self.clone();
        let bytes = crate::advanced::remove_drawing(
            &candidate.candidate,
            sheet,
            drawing_name,
            candidate.before.limits.package_bytes,
        )?;
        candidate.stage_spliced("drawing.remove", drawing_name, bytes)?;
        candidate.remove_resource(resource_path)?;
        *self = candidate;
        Ok(())
    }

    /// Add, reuse, or explicitly replace one bounded auxiliary resource.
    ///
    /// # Errors
    ///
    /// Returns an error for a collision, unsafe path, oversized resource, noncompact XML, signed,
    /// encrypted, malformed, or over-budget package candidate.
    pub fn put_resource(
        &mut self,
        resource: Resource,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        validate_resource_size(resource.bytes.len(), self.before.limits)?;
        let package = Package::from_bytes(self.candidate.clone())?;
        let existing = package.package().has_file(&resource.path)?;
        if existing {
            let bytes = package.package().get_file(&resource.path)?;
            let reader = package.package().package()?;
            let media_type = reader
                .manifest()
                .get_media_type(&resource.path)
                .unwrap_or("application/octet-stream");
            let equivalent =
                bytes.as_slice() == resource.as_bytes() && media_type == resource.media_type();
            match collision {
                Collision::Reject => return invalid("ODS resource destination already exists"),
                Collision::ReuseEquivalent if equivalent => {
                    return Ok(TransferDisposition::Reused);
                },
                Collision::ReuseEquivalent => {
                    return invalid("ODS resource destination is not equivalent");
                },
                Collision::Replace if equivalent => return Ok(TransferDisposition::Reused),
                Collision::Replace => {},
            }
        }
        let bytes = replace_resource(
            &package,
            Some(&resource),
            &resource.path,
            self.before.limits,
        )?;
        self.stage("resource.put", &resource.path, bytes)?;
        Ok(if existing {
            TransferDisposition::Replaced
        } else {
            TransferDisposition::Added
        })
    }

    /// Copy one bounded inert resource from another immutable ODS snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the source resource is absent or either source/destination constraint
    /// rejects the transfer.
    pub fn transfer_resource(
        &mut self,
        source: &Snapshot,
        source_path: &str,
        destination_path: &str,
        collision: Collision,
    ) -> Result<TransferDisposition> {
        let source_resource = source.resource(source_path)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODS transfer resource '{source_path}' was not found"
            ))
        })?;
        let resource = Resource::new(
            destination_path,
            source_resource.media_type,
            source_resource.bytes.as_ref().to_vec(),
        )?;
        self.put_resource(resource, collision)
    }

    /// Remove one unreferenced auxiliary resource.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is absent, referenced by retained XML, reserved, or package
    /// rebuilding fails.
    pub fn remove_resource(&mut self, path: &str) -> Result<()> {
        validate_resource_path(path)?;
        let package = Package::from_bytes(self.candidate.clone())?;
        if !package.package().has_file(path)? {
            return invalid(format!("ODS resource '{path}' was not found"));
        }
        refuse_referenced_removal(&package, path)?;
        let bytes = replace_resource(&package, None, path, self.before.limits)?;
        self.stage("resource.remove", path, bytes)
    }

    /// Restore the exact source candidate and discard every staged semantic operation.
    pub fn rollback(&mut self) {
        self.candidate = self.before.source.as_ref().to_vec();
        self.steps.clear();
        self.spliced_parts.clear();
    }

    /// Validate security policy, reopen the complete candidate, and publish one durable patch.
    ///
    /// Exact no-ops retain the source allocation and remain permitted for signed, encrypted, or
    /// protected input. Any changed transaction refuses those inputs before publication.
    ///
    /// # Errors
    ///
    /// Returns an error for a security refusal, bound failure, malformed candidate, typed readback
    /// failure, noncompact authored XML, or durable-patch construction failure.
    pub fn commit(self) -> Result<Commit> {
        self.commit_with_security(SecurityWritePolicy::default())
    }

    /// Validate and publish under an explicit signature/encryption write policy.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected policy refuses signed, encrypted, or protected input,
    /// or when bounds, compactness, readback, or durable patch construction fails.
    pub fn commit_with_security(self, policy: SecurityWritePolicy) -> Result<Commit> {
        if self.candidate.as_slice() == self.before.as_bytes() || self.steps.is_empty() {
            let patch = Patch::build(
                self.before.source.clone(),
                self.before.source.clone(),
                Vec::new(),
                self.before.limits,
            )?;
            return Ok(Commit {
                snapshot: self.before,
                patch,
            });
        }
        enforce_security_policy(&self.before, policy)?;
        validate_package_size(self.candidate.len(), self.before.limits)?;
        validate_authored_parts(&self.before.source, &self.candidate, &self.spliced_parts)?;
        let snapshot = Snapshot::from_bytes_with(self.candidate, self.before.limits)?;
        let patch = Patch::build(
            self.before.source,
            snapshot.source.clone(),
            self.steps,
            snapshot.limits,
        )?;
        Ok(Commit { snapshot, patch })
    }

    fn stage(&mut self, op: &str, target: &str, candidate: Vec<u8>) -> Result<()> {
        self.stage_shared(op, target, Arc::new(candidate))
    }

    fn stage_shared(&mut self, op: &str, target: &str, candidate: Arc<Vec<u8>>) -> Result<()> {
        validate_package_size(candidate.len(), self.before.limits)?;
        let effects = changed_effects(&self.candidate, candidate.as_slice())?;
        if effects.is_empty() {
            return Ok(());
        }
        for effect in &effects {
            if let Some(path) = effect.strip_prefix("part:") {
                self.spliced_parts.remove(path);
            }
        }
        let validated = Package::from_shared_bytes(Arc::clone(&candidate))?;
        drop(validated);
        self.candidate = take_shared_bytes(candidate);
        self.steps.push(Step {
            op: op.to_string(),
            target: target.to_string(),
            effects,
        });
        Ok(())
    }

    fn stage_spliced(&mut self, op: &str, target: &str, candidate: Vec<u8>) -> Result<()> {
        self.stage(op, target, candidate)?;
        self.spliced_parts.insert("content.xml".to_string());
        Ok(())
    }

    fn stage_spliced_shared(
        &mut self,
        op: &str,
        target: &str,
        candidate: Arc<Vec<u8>>,
    ) -> Result<()> {
        self.stage_shared(op, target, candidate)?;
        self.spliced_parts.insert("content.xml".to_string());
        Ok(())
    }
}

fn take_shared_bytes(bytes: Arc<Vec<u8>>) -> Vec<u8> {
    Arc::try_unwrap(bytes).unwrap_or_else(|shared| (*shared).clone())
}

/// A durable, exact-source, reversible unified package patch.
#[derive(Clone)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    semantic: CorePatch<Reversible>,
    steps: Vec<Step>,
    limits: Limits,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("source", &self.source.len())
            .field("target", &self.target.len())
            .field("semantic", &self.semantic)
            .field("steps", &self.steps)
            .field("limits", &self.limits)
            .finish()
    }
}

impl Patch {
    fn build(
        source: Arc<[u8]>,
        target: Arc<[u8]>,
        mut steps: Vec<Step>,
        limits: Limits,
    ) -> Result<Self> {
        if steps.is_empty() {
            if source != target {
                return invalid("ODS package changed without a semantic owner operation");
            }
            let mut blobs = BlobBundle::new(limits.patch.blobs());
            let _source_id = blobs
                .insert_shared(Arc::clone(&source))
                .map_err(patch_error)?;
            let semantic = CorePatch::<Reversible>::new(
                limits.patch,
                FORMAT,
                Vec::<ReversibleOperation>::new(),
                blobs.clone(),
                blobs,
            )
            .map_err(patch_error)?;
            return Ok(Self {
                source,
                target,
                semantic,
                steps,
                limits,
            });
        }
        let actual_effects = changed_effects(&source, &target)?
            .into_iter()
            .collect::<BTreeSet<_>>();
        for step in &mut steps {
            step.effects
                .retain(|effect| actual_effects.contains(effect));
        }
        steps.retain(|step| !step.effects.is_empty());
        if source != target && steps.is_empty() {
            return invalid("ODS package changed without a semantic owner operation");
        }
        let mut forward_blobs = BlobBundle::new(limits.patch.blobs());
        let mut reverse_blobs = BlobBundle::new(limits.patch.blobs());
        let target_id = forward_blobs
            .insert_shared(Arc::clone(&target))
            .map_err(patch_error)?;
        let source_id = reverse_blobs
            .insert_shared(Arc::clone(&source))
            .map_err(patch_error)?;
        let source_hash = source_id.as_hex();
        let target_hash = target_id.as_hex();
        let mut operations = Vec::new();
        operations
            .try_reserve_exact(steps.len())
            .map_err(|_allocation_error| invalid_error("ODS patch operation allocation failed"))?;
        for step in &steps {
            let forward =
                semantic_operation(limits.patch, step, &source_hash, &target_hash, "forward")?;
            let inverse =
                semantic_operation(limits.patch, step, &target_hash, &source_hash, "inverse")?;
            operations.push(ReversibleOperation::new(forward, inverse));
        }
        let semantic = CorePatch::<Reversible>::new(
            limits.patch,
            FORMAT,
            operations,
            forward_blobs,
            reverse_blobs,
        )
        .map_err(patch_error)?;
        Ok(Self {
            source,
            target,
            semantic,
            steps,
            limits,
        })
    }

    /// Whether this patch changes exact package bytes.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.source != self.target
    }

    /// Semantic operations in deterministic transaction order.
    #[must_use]
    pub fn operations(&self) -> &[PatchOperation] {
        self.semantic.operations()
    }

    /// Check exact source applicability.
    #[must_use]
    pub fn is_applicable_to(&self, snapshot: &Snapshot) -> bool {
        self.source.as_ref() == snapshot.as_bytes()
    }

    /// Serialize the bounded semantic envelope as canonical deterministic JSON.
    ///
    /// # Errors
    ///
    /// Returns an error when the configured output bound is exceeded.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        self.semantic.to_deterministic_json().map_err(patch_error)
    }

    /// Parse a canonical deterministic semantic envelope with explicit ODS limits.
    ///
    /// # Errors
    ///
    /// Returns an error for noncanonical/untrusted JSON, wrong format vocabulary, malformed
    /// operations or blobs, package-bound failures, or mismatched fingerprints.
    pub fn from_deterministic_json(bytes: &[u8], limits: Limits) -> Result<Self> {
        let semantic = CorePatch::<Reversible>::from_deterministic_json(bytes, limits.patch)
            .map_err(patch_error)?;
        Self::from_semantic(semantic, limits)
    }

    /// Apply only to the exact immutable package that authorized this patch.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes, security refusal, bounds, malformed target, or
    /// complete facade readback failure.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        self.apply_with_security(snapshot, SecurityWritePolicy::default())
    }

    /// Apply under an explicit signature/encryption write policy.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, a security-policy refusal, bounds, malformed target,
    /// or complete facade readback failure.
    pub fn apply_with_security(
        &self,
        snapshot: &Snapshot,
        policy: SecurityWritePolicy,
    ) -> Result<Commit> {
        if !self.is_applicable_to(snapshot) {
            return invalid("ODS unified patch source snapshot does not match");
        }
        if self.changed() {
            enforce_security_policy(snapshot, policy)?;
        }
        let target = Snapshot::from_arc(self.target.clone(), self.limits)?;
        Ok(Commit {
            snapshot: target,
            patch: self.clone(),
        })
    }

    /// Return the durable semantic patch restoring the exact accepted source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let semantic = self.semantic.inverse();
        let mut steps = self.steps.clone();
        steps.reverse();
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            semantic,
            steps,
            limits: self.limits,
        }
    }

    /// Deterministically join two provably disjoint patches from one exact source.
    ///
    /// # Errors
    ///
    /// Returns a structured lineage, limit, semantic-overlap, or physical-overlap refusal.
    pub fn join(&self, other: &Self) -> std::result::Result<Self, JoinError> {
        if self.limits != other.limits {
            return Err(JoinError {
                failure: JoinFailure::DifferentLimits,
                detail: "ODS semantic patches use different finite bounds".to_string(),
            });
        }
        let mut joined = JoinedSubEdits::new(self.source.clone(), self.limits.composition);
        joined
            .join(self.as_sub_edit().map_err(JoinError::candidate)?)
            .map_err(JoinError::from_core)?;
        joined
            .join(other.as_sub_edit().map_err(JoinError::candidate)?)
            .map_err(JoinError::from_core)?;
        joined_patch(joined, self.limits).map_err(JoinError::candidate)
    }

    /// Build a non-applying deterministic three-way plan from two branches.
    ///
    /// # Errors
    ///
    /// Returns an error when lineage, limits, or finite planning bounds differ.
    pub fn three_way(left: &Self, right: &Self) -> Result<ThreeWayPlan> {
        if left.limits != right.limits {
            return invalid("ODS three-way branches use different finite bounds");
        }
        let mut left_branch = JoinedSubEdits::new(left.source.clone(), left.limits.composition);
        left_branch
            .join(left.as_sub_edit()?)
            .map_err(|error| composition_error(error.failure()))?;
        let mut right_branch = JoinedSubEdits::new(right.source.clone(), right.limits.composition);
        right_branch
            .join(right.as_sub_edit()?)
            .map_err(|error| composition_error(error.failure()))?;
        let plan = ThreeWayMergePlan::new(left_branch, right_branch).map_err(|error| {
            invalid_error(format!(
                "ODS three-way planning failed: {:?}",
                error.failure()
            ))
        })?;
        Ok(ThreeWayPlan {
            plan: Some(plan),
            limits: left.limits,
        })
    }

    fn as_sub_edit(&self) -> Result<SubEdit<Arc<[u8]>, Branch>> {
        let id = self
            .semantic
            .fingerprint()
            .unwrap_or_else(|_error| DiagnosticFingerprint::of(&self.target))
            .as_hex();
        SubEdit::new(
            self.source.clone(),
            self.limits.composition,
            id,
            Vec::new(),
            patch_effects(&self.steps),
            Branch {
                target: self.target.clone(),
                steps: self.steps.clone(),
            },
        )
        .map_err(|error| invalid_error(format!("ODS composition sub-edit failed: {error}")))
    }

    fn from_semantic(semantic: CorePatch<Reversible>, limits: Limits) -> Result<Self> {
        if semantic.format() != FORMAT {
            return invalid("durable patch does not use the ODS document vocabulary");
        }
        let inverse = semantic.inverse();
        let target = only_blob(semantic.blobs())?;
        let source = only_blob(inverse.blobs())?;
        validate_package_size(source.len(), limits)?;
        validate_package_size(target.len(), limits)?;
        let source_hash = DiagnosticFingerprint::of(&source).as_hex();
        let target_hash = DiagnosticFingerprint::of(&target).as_hex();
        let (forward_direction, inverse_direction) =
            match semantic.operations().first().and_then(operation_direction) {
                Some("forward") | None => ("forward", "inverse"),
                Some("inverse") => ("inverse", "forward"),
                Some(_) => return invalid("ODS durable patch operation direction is invalid"),
            };
        let steps = parse_steps(
            semantic.operations(),
            &source_hash,
            &target_hash,
            forward_direction,
        )?;
        let inverse_steps = parse_steps(
            inverse.operations(),
            &target_hash,
            &source_hash,
            inverse_direction,
        )?;
        if inverse_steps.iter().rev().ne(steps.iter()) {
            return invalid("ODS durable patch inverse operations differ");
        }
        let actual_effects = changed_effects(&source, &target)?;
        if patch_effects(&steps) != actual_effects
            || (source == target && !steps.is_empty())
            || (source != target && steps.is_empty())
        {
            return invalid("ODS durable patch effects do not match its package members");
        }
        let source: Arc<[u8]> = Arc::from(source);
        let target: Arc<[u8]> = Arc::from(target);
        let _source_snapshot = Snapshot::from_arc(source.clone(), limits)?;
        let _target_snapshot = Snapshot::from_arc(target.clone(), limits)?;
        Ok(Self {
            source,
            target,
            semantic,
            steps,
            limits,
        })
    }
}

#[derive(Clone, Debug)]
struct Branch {
    target: Arc<[u8]>,
    steps: Vec<Step>,
}

/// Why two unified package patches could not be joined.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum JoinFailure {
    /// Exact package source lineage differs.
    DifferentLineage,
    /// The branches use different finite bounds.
    DifferentLimits,
    /// Stable operation identity or semantic effects overlap.
    Conflict(ConflictSet<SubEditConflict>),
    /// A finite composition bound was exceeded.
    Limit,
    /// Disjoint semantic declarations still produced incompatible physical members.
    PhysicalOverlap,
    /// A merged package could not be validated or rebuilt.
    InvalidCandidate,
}

/// Structured recoverable join refusal.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct JoinError {
    failure: JoinFailure,
    detail: String,
}

impl JoinError {
    /// Structured refusal class.
    #[must_use]
    pub const fn failure(&self) -> &JoinFailure {
        &self.failure
    }

    /// Stable diagnostic detail without package bytes.
    #[must_use]
    pub fn detail(&self) -> &str {
        &self.detail
    }

    fn candidate(error: Error) -> Self {
        let detail = format!("{error}");
        let failure = if detail.contains("physical member overlap") {
            JoinFailure::PhysicalOverlap
        } else {
            JoinFailure::InvalidCandidate
        };
        Self { failure, detail }
    }

    fn from_core(error: litchi_core::SubEditJoinError<Arc<[u8]>, Branch>) -> Self {
        let failure = match error.failure() {
            SubEditJoinFailure::DifferentLineage => JoinFailure::DifferentLineage,
            SubEditJoinFailure::DifferentLimits => JoinFailure::DifferentLimits,
            SubEditJoinFailure::DuplicateId => {
                JoinFailure::Conflict(ConflictSet::new([SubEditConflict::DuplicateId(
                    error.rejected().id().to_string(),
                )]))
            },
            SubEditJoinFailure::Overlap(conflicts) => JoinFailure::Conflict(conflicts.clone()),
            SubEditJoinFailure::Limit(_) => JoinFailure::Limit,
            _ => JoinFailure::Limit,
        };
        Self {
            failure,
            detail: "ODS semantic patch join was refused".to_string(),
        }
    }
}

/// Non-applying three-way merge plan for two exact-source branches.
pub struct ThreeWayPlan {
    plan: Option<ThreeWayMergePlan<Arc<[u8]>, Branch>>,
    limits: Limits,
}

impl fmt::Debug for ThreeWayPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ThreeWayPlan")
            .field("plan", &self.plan)
            .field("limits", &self.limits)
            .finish()
    }
}

impl ThreeWayPlan {
    /// Deterministically ordered semantic conflicts.
    #[must_use]
    pub fn conflicts(&self) -> Option<&ConflictSet<SubEditConflict>> {
        self.plan.as_ref().map(ThreeWayMergePlan::conflicts)
    }

    /// Resolve the complete conservative conflict group explicitly.
    pub fn resolve(&mut self, choice: MergeChoice) -> &mut Self {
        if let Some(plan) = self.plan.as_mut() {
            plan.resolve(choice);
        }
        self
    }

    /// Finish the selected plan and build one validated unified package patch.
    ///
    /// # Errors
    ///
    /// Returns an error while conflicts remain unresolved or merged package validation fails.
    pub fn finish(mut self) -> Result<Patch> {
        let plan = self
            .plan
            .take()
            .ok_or_else(|| invalid_error("ODS three-way plan was already consumed"))?;
        let joined = plan.finish().map_err(|_plan| {
            invalid_error("ODS three-way conflicts require explicit resolution")
        })?;
        joined_patch(joined, self.limits)
    }
}

/// One immutable package publication with durable reversible semantics.
#[derive(Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether exact package bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.changed()
    }

    /// Resulting immutable package snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Durable exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume this publication into its immutable package snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// Explicit bounded undo/redo over immutable ODS package snapshots.
pub struct History {
    inner: CoreHistory<Snapshot>,
}

impl History {
    /// Start history with the snapshot's configured finite history bounds.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        Self {
            inner: CoreHistory::new(snapshot.clone(), snapshot.limits.history),
        }
    }

    /// Current immutable package snapshot.
    #[must_use]
    pub fn current(&self) -> &Snapshot {
        self.inner.current()
    }

    /// Whether an undo transition is available.
    #[must_use]
    pub fn can_undo(&self) -> bool {
        self.inner.can_undo()
    }

    /// Whether a redo transition is available.
    #[must_use]
    pub fn can_redo(&self) -> bool {
        self.inner.can_redo()
    }

    /// Record one exact-source commit under its serialized transfer weight.
    ///
    /// # Errors
    ///
    /// Returns an error for stale lineage, serialization failure, or history-weight refusal.
    pub fn record(&mut self, commit: Commit) -> Result<Vec<Snapshot>> {
        if !commit.patch.is_applicable_to(self.current()) {
            return invalid("ODS history commit does not descend from the current snapshot");
        }
        let weight = u64::try_from(commit.patch.to_deterministic_json()?.len())
            .map_err(|_error| invalid_error("ODS history weight exceeds u64"))?;
        self.inner
            .record(commit.snapshot, weight)
            .map_err(patch_error)
    }

    /// Move one retained transition backward.
    pub fn undo(&mut self) -> bool {
        self.inner.undo()
    }

    /// Move one retained transition forward.
    pub fn redo(&mut self) -> bool {
        self.inner.redo()
    }
}

fn changed_effects(source: &[u8], target: &[u8]) -> Result<Vec<String>> {
    Ok(changed_package_files(source, target)?
        .into_iter()
        .map(|(path, _target)| format!("part:{path}"))
        .collect())
}

fn composition_error(failure: &SubEditJoinFailure) -> Error {
    invalid_error(format!("ODS composition was refused: {failure:?}"))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn content_validation_error(error: crate::content_validation::Error) -> Error {
    invalid_error(format!("ODS content-validation edit failed: {error}"))
}

fn is_reserved_path(path: &str) -> bool {
    CORE_PATHS.contains(&path) || path.starts_with("META-INF/")
}

fn is_xml_media_type(media_type: &str) -> bool {
    media_type == "text/xml"
        || media_type == "application/xml"
        || media_type == "application/rdf+xml"
        || media_type.ends_with("+xml")
}

fn joined_patch(joined: JoinedSubEdits<Arc<[u8]>, Branch>, limits: Limits) -> Result<Patch> {
    let source = joined.lineage().clone();
    let mut branches = joined
        .into_sub_edits()
        .map(SubEdit::into_payload)
        .collect::<Vec<_>>();
    if branches.is_empty() {
        return Patch::build(source.clone(), source, Vec::new(), limits);
    }
    branches.sort_by(|left, right| {
        DiagnosticFingerprint::of(&left.target)
            .as_hex()
            .cmp(&DiagnosticFingerprint::of(&right.target).as_hex())
    });
    let targets = branches
        .iter()
        .map(|branch| branch.target.as_ref())
        .collect::<Vec<_>>();
    let target = Arc::from(merge_package_targets(&source, &targets, limits)?);
    let steps = branches
        .into_iter()
        .flat_map(|branch| branch.steps)
        .collect::<Vec<_>>();
    Patch::build(source, target, steps, limits)
}

fn merge_package_targets(source: &[u8], targets: &[&[u8]], limits: Limits) -> Result<Vec<u8>> {
    let base_files = package_files(source)?;
    let target_maps = targets
        .iter()
        .map(|target| package_files(target))
        .collect::<Result<Vec<_>>>()?;
    let mut paths = BTreeSet::new();
    paths.extend(base_files.keys().cloned());
    for target in &target_maps {
        paths.extend(target.keys().cloned());
    }
    let mut selected = base_files.clone();
    for path in paths {
        let base = base_files.get(&path);
        let mut changed = target_maps.iter().filter_map(|target| {
            let candidate = target.get(&path);
            (candidate != base).then_some(candidate)
        });
        let Some(first) = changed.next() else {
            continue;
        };
        if changed.any(|candidate| candidate != first) {
            return invalid(format!(
                "ODS joined branches have a physical member overlap at '{path}'"
            ));
        }
        match first {
            Some(file) => {
                selected.insert(path, file.clone());
            },
            None => {
                selected.remove(&path);
            },
        }
    }
    rebuild_selected(source, &base_files, &selected, limits)
}

fn only_blob(bundle: &BlobBundle) -> Result<Vec<u8>> {
    if bundle.len() != 1 {
        return invalid("ODS durable patch must contain exactly one package blob per direction");
    }
    let id = bundle
        .ids()
        .next()
        .ok_or_else(|| invalid_error("ODS durable patch package blob is missing"))?;
    bundle
        .get(id)
        .map(<[u8]>::to_vec)
        .ok_or_else(|| invalid_error("ODS durable patch package blob cannot be resolved"))
}

fn package_files(bytes: &[u8]) -> Result<BTreeMap<String, File>> {
    let package = Package::from_bytes(bytes.to_vec())?;
    let reader = package.package().package()?;
    let mut files = BTreeMap::new();
    for path in reader.files()? {
        if matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        let bytes = reader.get_file(&path)?;
        let media_type = reader
            .manifest()
            .get_media_type(&path)
            .unwrap_or("application/octet-stream")
            .to_string();
        files.insert(path, File { bytes, media_type });
    }
    Ok(files)
}

fn changed_package_files(source: &[u8], target: &[u8]) -> Result<Vec<(String, Option<File>)>> {
    let raw_identical = raw_identical_members(source, target)
        .filter(|paths| paths.contains("META-INF/manifest.xml"));
    let source_package = Package::from_bytes(source.to_vec())?;
    let target_package = Package::from_bytes(target.to_vec())?;
    let source_reader = source_package.package().package()?;
    let target_reader = target_package.package().package()?;
    let mut paths = BTreeSet::new();
    paths.extend(source_reader.files()?);
    paths.extend(target_reader.files()?);
    let mut changed = Vec::new();
    for path in paths {
        if matches!(path.as_str(), "mimetype" | "META-INF/manifest.xml") || path.ends_with('/') {
            continue;
        }
        if raw_identical
            .as_ref()
            .is_some_and(|identical| identical.contains(&path))
        {
            continue;
        }
        let source_file = package_file(&source_reader, &path)?;
        let target_file = package_file(&target_reader, &path)?;
        if source_file != target_file {
            changed.push((path, target_file));
        }
    }
    Ok(changed)
}

fn package_file(
    package: &litchi_odf_common::core::package::Package<'_>,
    path: &str,
) -> Result<Option<File>> {
    if !package.has_file(path) {
        return Ok(None);
    }
    Ok(Some(File {
        bytes: package.get_file(path)?,
        media_type: package
            .manifest()
            .get_media_type(path)
            .unwrap_or("application/octet-stream")
            .to_string(),
    }))
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct File {
    bytes: Vec<u8>,
    media_type: String,
}

fn parse_steps(
    operations: &[PatchOperation],
    expected_source: &str,
    expected_target: &str,
    expected_direction: &str,
) -> Result<Vec<Step>> {
    let mut steps = Vec::new();
    steps
        .try_reserve_exact(operations.len())
        .map_err(|_allocation_error| invalid_error("ODS patch operation allocation failed"))?;
    for operation in operations {
        let Some(op) = operation.op.strip_prefix("ods.") else {
            return invalid("durable ODS patch contains an unknown operation");
        };
        if !known_operation(op) {
            return invalid("durable ODS patch contains an unknown operation");
        }
        if operation.preconditions.get("source_sha256") != Some(&json!(expected_source))
            || operation.preconditions.get("target_sha256") != Some(&json!(expected_target))
        {
            return invalid("durable ODS patch fingerprint precondition does not match its blobs");
        }
        let object = operation
            .value
            .as_object()
            .ok_or_else(|| invalid_error("durable ODS patch operation value must be an object"))?;
        if object.get("direction").and_then(Value::as_str) != Some(expected_direction) {
            return invalid("durable ODS patch operation direction is invalid");
        }
        let effects = object
            .get("effects")
            .and_then(Value::as_array)
            .ok_or_else(|| invalid_error("durable ODS patch operation effects are missing"))?
            .iter()
            .map(|value| {
                value
                    .as_str()
                    .map(str::to_string)
                    .ok_or_else(|| invalid_error("durable ODS patch effect must be text"))
            })
            .collect::<Result<Vec<_>>>()?;
        if effects.is_empty()
            || effects.windows(2).any(|pair| pair[0] >= pair[1])
            || effects.iter().any(|effect| !effect.starts_with("part:"))
        {
            return invalid("durable ODS patch effects are empty or noncanonical");
        }
        steps.push(Step {
            op: op.to_string(),
            target: operation.target.clone(),
            effects,
        });
    }
    Ok(steps)
}

fn operation_direction(operation: &PatchOperation) -> Option<&str> {
    operation
        .value
        .as_object()
        .and_then(|object| object.get("direction"))
        .and_then(Value::as_str)
}

fn known_operation(operation: &str) -> bool {
    matches!(
        operation,
        "worksheet.edit"
            | "definition.edit"
            | "annotation.edit"
            | "rdf.edit"
            | "protection.edit"
            | "data-pilot.edit"
            | "tracked-change.edit"
            | "chart.edit"
            | "resource.put"
            | "resource.remove"
            | "cell.rich-text"
            | "cell.formula"
            | "cell.style"
            | "row.insert"
            | "row.remove"
            | "column.insert"
            | "column.remove"
            | "sheet.insert"
            | "sheet.remove"
            | "style.put"
            | "style-graph.put"
            | "style-graph.replace"
            | "style-graph.remove"
            | "conditional-format.edit"
            | "sparkline.edit"
            | "drawing.put"
            | "drawing.remove"
            | "drawing-frame.put"
            | "drawing-group.put"
            | "drawing-geometry.put"
            | "drawing-polygon.put"
            | "chart-object.put"
            | "form.edit"
            | "form.bindings"
            | "form.rich"
    )
}

fn patch_effects(steps: &[Step]) -> Vec<String> {
    steps
        .iter()
        .flat_map(|step| step.effects.iter().cloned())
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect()
}

fn patch_error(error: PatchError) -> Error {
    invalid_error(format!("ODS durable patch failed: {error}"))
}

fn rebuild_selected(
    source: &[u8],
    base: &BTreeMap<String, File>,
    selected: &BTreeMap<String, File>,
    limits: Limits,
) -> Result<Vec<u8>> {
    for path in ["styles.xml", "meta.xml", "settings.xml"] {
        if base.get(path) != selected.get(path) {
            return invalid(format!(
                "ODS unified merge cannot replace reserved part '{path}'"
            ));
        }
    }
    let content = selected
        .get("content.xml")
        .ok_or_else(|| invalid_error("ODS unified merge removed content.xml"))?;
    let content = std::str::from_utf8(&content.bytes)
        .map_err(|_error| invalid_error("ODS merged content.xml is not UTF-8"))?;
    let source_package = Package::from_bytes(source.to_vec())?;
    let mut additions = Vec::new();
    let mut excluded = Vec::new();
    let mut paths = BTreeSet::new();
    paths.extend(base.keys().cloned());
    paths.extend(selected.keys().cloned());
    for path in paths {
        if matches!(path.as_str(), "content.xml" | "styles.xml" | "meta.xml") {
            continue;
        }
        if base.get(&path) == selected.get(&path) {
            continue;
        }
        excluded.push(path.clone());
        if let Some(file) = selected.get(&path) {
            validate_xml_resource(&file.media_type, &file.bytes)?;
            additions.push(Addition {
                path,
                bytes: file.bytes.clone(),
                media_type: file.media_type.clone(),
            });
        }
    }
    let bytes = rebuild_package(
        source_package.package(),
        content,
        additions,
        Vec::new(),
        excluded,
        Vec::<String>::new(),
    )?;
    validate_package_size(bytes.len(), limits)?;
    Ok(bytes)
}

fn refuse_referenced_removal(package: &Package, path: &str) -> Result<()> {
    for candidate in package.package().files()? {
        if candidate == path
            || matches!(candidate.as_str(), "mimetype" | "META-INF/manifest.xml")
            || candidate.ends_with('/')
        {
            continue;
        }
        let bytes = package.package().get_file(&candidate)?;
        let Ok(text) = std::str::from_utf8(&bytes) else {
            continue;
        };
        if text.contains(path) || text.contains(&format!("./{path}")) {
            return invalid(format!(
                "ODS resource '{path}' is referenced by retained part '{candidate}'"
            ));
        }
    }
    Ok(())
}

fn enforce_security_policy(snapshot: &Snapshot, policy: SecurityWritePolicy) -> Result<()> {
    let package = Package::from_bytes(snapshot.source.as_ref().to_vec())?;
    let reader = package.package().package()?;
    if reader.manifest().has_encrypted_entries() {
        return match policy.encryption {
            EncryptionWritePolicy::Refuse => {
                invalid("changed unified ODS transactions refuse encrypted package members")
            },
            EncryptionWritePolicy::DecryptAndReencrypt => invalid(
                "ODS decrypt/re-encrypt publication requires an unavailable credential capability",
            ),
        };
    }
    if (reader.has_file(DOCUMENT_SIGNATURE_PATH) || reader.has_file(MACRO_SIGNATURE_PATH))
        && matches!(policy.signatures, SignatureWritePolicy::Refuse)
    {
        return invalid("changed unified ODS transactions require explicit signature stripping");
    }
    let protection =
        crate::protection::Snapshot::parse(package.content_xml(), package.styles_xml())?;
    if protection.document().structure_protected == Some(true)
        || protection
            .sheets()
            .iter()
            .any(|sheet| sheet.is_protected() == Some(true))
    {
        return invalid("changed unified ODS transactions require an explicit unlock capability");
    }
    Ok(())
}

fn replace_resource(
    package: &Package,
    resource: Option<&Resource>,
    path: &str,
    limits: Limits,
) -> Result<Vec<u8>> {
    let additions = resource.map_or_else(Vec::new, |value| {
        vec![Addition {
            path: value.path.clone(),
            bytes: value.bytes.as_ref().to_vec(),
            media_type: value.media_type.clone(),
        }]
    });
    let bytes = rebuild_package(
        package.package(),
        package.content_xml(),
        additions,
        Vec::new(),
        [path.to_string()],
        Vec::<String>::new(),
    )?;
    validate_package_size(bytes.len(), limits)?;
    Ok(bytes)
}

fn semantic_operation(
    limits: PatchLimits,
    step: &Step,
    source_hash: &str,
    target_hash: &str,
    direction: &str,
) -> Result<PatchOperation> {
    let mut preconditions = BTreeMap::new();
    preconditions.insert("source_sha256".to_string(), json!(source_hash));
    preconditions.insert("target_sha256".to_string(), json!(target_hash));
    PatchOperation::new(
        limits,
        format!("ods.{}", step.op),
        step.target.clone(),
        preconditions,
        json!({"direction":direction,"effects":step.effects}),
    )
    .map_err(patch_error)
}

fn validate_authored_parts(
    source: &[u8],
    target: &[u8],
    provenance_spliced: &BTreeSet<String>,
) -> Result<()> {
    for (path, target_file) in changed_package_files(source, target)? {
        let Some(file) = target_file else {
            continue;
        };
        if is_xml_media_type(&file.media_type) && !provenance_spliced.contains(&path) {
            litchi_odf_common::compact_xml::validate(&file.bytes).map_err(Error::from)?;
        }
    }
    let package = Package::from_bytes(target.to_vec())?;
    let manifest = package.package().get_file("META-INF/manifest.xml")?;
    litchi_odf_common::compact_xml::validate(&manifest).map_err(Error::from)
}

fn validate_media_type(media_type: &str) -> Result<()> {
    if media_type.is_empty()
        || media_type.len() > 256
        || media_type.chars().any(char::is_control)
        || media_type.chars().any(char::is_whitespace)
    {
        invalid("ODS resource media type is invalid")
    } else {
        Ok(())
    }
}

fn validate_package_size(length: usize, limits: Limits) -> Result<()> {
    if length > limits.package_bytes {
        invalid(format!(
            "ODS package byte limit exceeded: observed {length}, limit {}",
            limits.package_bytes
        ))
    } else {
        Ok(())
    }
}

fn validate_resource_count(package: &Package, limits: Limits) -> Result<()> {
    let count = package
        .package()
        .files()?
        .into_iter()
        .filter(|path| !is_reserved_path(path) && !path.ends_with('/'))
        .count();
    if count > limits.resources {
        invalid(format!(
            "ODS resource count limit exceeded: observed {count}, limit {}",
            limits.resources
        ))
    } else {
        Ok(())
    }
}

fn validate_resource_path(path: &str) -> Result<()> {
    if path.is_empty()
        || path.len() > MAX_PATH_BYTES
        || path.starts_with('/')
        || path.ends_with('/')
        || path.contains('\\')
        || path.split('/').any(|segment| {
            segment.is_empty()
                || segment == "."
                || segment == ".."
                || segment.chars().any(char::is_control)
        })
        || is_reserved_path(path)
    {
        invalid("ODS resource path is unsafe or reserved")
    } else {
        Ok(())
    }
}

pub(crate) fn validate_detached_resource_path(path: &str) -> Result<()> {
    validate_resource_path(path)
}

fn validate_resource_size(length: usize, limits: Limits) -> Result<()> {
    if length > limits.resource_bytes {
        invalid(format!(
            "ODS resource byte limit exceeded: observed {length}, limit {}",
            limits.resource_bytes
        ))
    } else {
        Ok(())
    }
}

fn validate_xml_resource(media_type: &str, bytes: &[u8]) -> Result<()> {
    validate_media_type(media_type)?;
    if is_xml_media_type(media_type) {
        litchi_odf_common::compact_xml::validate(bytes).map_err(Error::from)?;
    }
    Ok(())
}

#[cfg(test)]
mod raw_package_diff_tests {
    #![allow(
        clippy::unwrap_used,
        reason = "fixed in-memory ZIP fixtures keep physical-diff assertions concise"
    )]

    use std::{
        io::{Cursor, Write},
        sync::Arc,
    };

    use litchi_core::DiagnosticFingerprint;

    use super::{Limits, Patch, Step, changed_effects};

    const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:spreadsheet><table:table xml:id="sheet" table:name="Sheet1"/></office:spreadsheet></office:body></office:document-content>"#;
    const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

    fn raw_package(
        compression: zip::CompressionMethod,
        resource_media_type: &str,
        resource: &[u8],
    ) -> Vec<u8> {
        let manifest = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIMETYPE}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Pictures/opaque.bin" manifest:media-type="{resource_media_type}"/></manifest:manifest>"#
        );
        let mut output = Cursor::new(Vec::new());
        {
            let mut archive = zip::ZipWriter::new(&mut output);
            let stored = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            let selected = zip::write::SimpleFileOptions::default().compression_method(compression);
            archive.start_file("mimetype", stored).unwrap();
            archive.write_all(MIMETYPE.as_bytes()).unwrap();
            for (path, bytes) in [
                ("META-INF/manifest.xml", manifest.as_bytes()),
                ("content.xml", CONTENT.as_bytes()),
                ("Pictures/opaque.bin", resource),
            ] {
                archive.start_file(path, selected).unwrap();
                archive.write_all(bytes).unwrap();
            }
            archive.finish().unwrap();
        }
        output.into_inner()
    }

    #[test]
    fn raw_fast_path_retains_logical_fallback_semantics() {
        let source = raw_package(
            zip::CompressionMethod::Stored,
            "application/octet-stream",
            b"same logical payload",
        );
        let recompressed = raw_package(
            zip::CompressionMethod::Deflated,
            "application/octet-stream",
            b"same logical payload",
        );
        assert!(changed_effects(&source, &recompressed).unwrap().is_empty());

        let changed_media_type = raw_package(
            zip::CompressionMethod::Deflated,
            "image/png",
            b"same logical payload",
        );
        assert_eq!(
            changed_effects(&source, &changed_media_type).unwrap(),
            ["part:Pictures/opaque.bin"]
        );

        let changed_payload = raw_package(
            zip::CompressionMethod::Deflated,
            "application/octet-stream",
            b"different logical payload",
        );
        assert_eq!(
            changed_effects(&source, &changed_payload).unwrap(),
            ["part:Pictures/opaque.bin"]
        );
    }

    #[test]
    fn durable_patch_blobs_share_exact_package_allocations_and_hashes() {
        let source: Arc<[u8]> = Arc::from(raw_package(
            zip::CompressionMethod::Stored,
            "application/octet-stream",
            b"source payload",
        ));
        let target: Arc<[u8]> = Arc::from(raw_package(
            zip::CompressionMethod::Stored,
            "application/octet-stream",
            b"target payload",
        ));
        let source_pointer = source.as_ptr();
        let target_pointer = target.as_ptr();
        let source_hash = DiagnosticFingerprint::of(&source).as_hex();
        let target_hash = DiagnosticFingerprint::of(&target).as_hex();
        let effects = changed_effects(&source, &target).unwrap();

        let patch = Patch::build(
            Arc::clone(&source),
            Arc::clone(&target),
            vec![Step {
                op: "resource.put".to_string(),
                target: "Pictures/opaque.bin".to_string(),
                effects,
            }],
            Limits::default(),
        )
        .unwrap();

        let target_id = patch.semantic.blobs().ids().next().unwrap();
        assert_eq!(target_id.as_hex(), target_hash);
        assert_eq!(
            patch.semantic.blobs().get(target_id).map(<[u8]>::as_ptr),
            Some(target_pointer)
        );
        let inverse = patch.inverse();
        let source_id = inverse.semantic.blobs().ids().next().unwrap();
        assert_eq!(source_id.as_hex(), source_hash);
        assert_eq!(
            inverse.semantic.blobs().get(source_id).map(<[u8]>::as_ptr),
            Some(source_pointer)
        );
        assert_eq!(
            patch.operations()[0].preconditions.get("source_sha256"),
            Some(&serde_json::json!(source_hash))
        );
        assert_eq!(
            patch.operations()[0].preconditions.get("target_sha256"),
            Some(&serde_json::json!(target_hash))
        );
    }

    #[test]
    fn exact_noop_patch_reuses_one_shared_allocation_in_both_directions() {
        let source: Arc<[u8]> = Arc::from(raw_package(
            zip::CompressionMethod::Stored,
            "application/octet-stream",
            b"unchanged payload",
        ));
        let source_pointer = source.as_ptr();
        let source_hash = DiagnosticFingerprint::of(&source).as_hex();

        let patch = Patch::build(
            Arc::clone(&source),
            Arc::clone(&source),
            Vec::new(),
            Limits::default(),
        )
        .unwrap();

        assert!(!patch.changed());
        assert!(patch.operations().is_empty());
        assert!(patch.steps.is_empty());
        assert_eq!(patch.semantic.blobs().len(), 1);
        let target_id = patch.semantic.blobs().ids().next().unwrap();
        assert_eq!(target_id.as_hex(), source_hash);
        assert_eq!(
            patch.semantic.blobs().get(target_id).map(<[u8]>::as_ptr),
            Some(source_pointer)
        );
        let inverse = patch.inverse();
        assert_eq!(inverse.semantic.blobs().len(), 1);
        let source_id = inverse.semantic.blobs().ids().next().unwrap();
        assert_eq!(
            inverse.semantic.blobs().get(source_id).map(<[u8]>::as_ptr),
            Some(source_pointer)
        );
        assert_eq!(
            patch.to_deterministic_json().unwrap(),
            inverse.to_deterministic_json().unwrap()
        );

        let changed: Arc<[u8]> = Arc::from(raw_package(
            zip::CompressionMethod::Stored,
            "application/octet-stream",
            b"changed without an owner",
        ));
        assert!(
            Patch::build(source, changed, Vec::new(), Limits::default()).is_err(),
            "changed packages still require a semantic owner operation"
        );
    }
}
