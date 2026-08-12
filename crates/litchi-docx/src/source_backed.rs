//! Lazy DOCX access and guarded main-document publication over an immutable
//! positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes. Main-document
//! transactions retain the raw XML and may be published to a sequential sink
//! while raw-copying every unselected ZIP member.

pub mod paragraph_copy;
pub mod paragraph_remove;

use crate::document::{Commit, Edit, Snapshot, TransactionResult};
use crate::error::{Error, Result};
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{
    document_paragraph, document_paragraph_count, document_paragraphs, visible_document_xml,
};
use crate::redact;
use crate::sanitize::{self, RelationshipState};
use crate::settings::DocumentSettings;
use crate::variables;
use litchi_core::ReadAt;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, Part, SourceArtifact, SourceArtifactFingerprint, SourceBackedPackage};
use sha2::{Digest as _, Sha256};
use smallvec::SmallVec;
use std::io::Write;
use std::sync::Arc;

/// A DOCX package that leaves ordinary part bodies cold at open.
pub struct Package {
    package: SourceBackedPackage,
}

struct DocumentVariablesSource {
    partname: litchi_opc::PackURI,
    snapshot: variables::Snapshot,
    protected: bool,
    unique_inbound_owner: bool,
}

/// Products of one exact source-backed document-variable publication.
pub struct DocumentVariablesPublication {
    snapshot: variables::Snapshot,
    original_snapshot: variables::Snapshot,
    original_artifact: SourceArtifact,
    published_artifact: SourceArtifactFingerprint,
}

impl DocumentVariablesPublication {
    /// Borrow the published settings snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &variables::Snapshot {
        &self.snapshot
    }
}

struct FingerprintingWriter<W> {
    inner: W,
    hasher: Sha256,
}

impl<W: Write> Write for FingerprintingWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        let written = self.inner.write(bytes)?;
        if written > bytes.len() {
            return Err(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!(
                    "document-variable sink reported {written} bytes for a {}-byte write",
                    bytes.len()
                ),
            ));
        }
        self.hasher.update(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

impl Package {
    /// Open a DOCX source using the standard bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at(source)?)
    }

    /// Open a DOCX source with an explicit bounded OPC read policy.
    ///
    /// This validates the main-document relationship and content type but
    /// does not decompress or materialize the main-document payload.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        limits: litchi_opc::ReadLimits,
    ) -> Result<Self> {
        Self::from_source_backed(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    fn from_source_backed(package: SourceBackedPackage) -> Result<Self> {
        let main = package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        Ok(Self { package })
    }

    /// Load and pin the main document for read-only semantic queries.
    ///
    /// The first call reads only the main-document part. The returned document
    /// owns its normalized XML bytes, so repeated text and paragraph queries do
    /// not revisit the positional source.
    pub fn document(&self) -> Result<Document> {
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let xml = visible_document_xml(main.data()?.into_arc())?;
        let source_version = self.package.source_version()?;
        Ok(Document {
            xml,
            source_version,
        })
    }

    /// Load only the mandatory main-document payload and capture its immutable
    /// section inventory at the exact opened source version.
    ///
    /// Header/footer relationship IDs are retained as inert values. Their
    /// target parts are neither resolved nor read.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_with_limits(&crate::section::Limits::default())
    }

    /// Capture a source-bound section inventory with explicit semantic limits.
    pub fn section_inventory_snapshot_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Snapshot> {
        let main = self.package.main_document_part()?;
        validate_document_main_content_type(main.content_type())?;
        let source_version = self.package.source_version()?;
        let raw = main.data()?;
        let snapshot =
            crate::section::Snapshot::from_source_xml(raw.as_bytes(), source_version, limits)?;
        // Detect hostile adapters that mutate immediately after the payload
        // read and semantic scan, before publishing the closure to the caller.
        let observed = self.package.source_version()?;
        if observed != source_version {
            return Err(litchi_opc::OpcError::SourceChanged {
                expected: source_version,
                actual: observed,
            }
            .into());
        }
        Ok(snapshot)
    }

    /// Return content-free payload-cache activity for this lazy package.
    ///
    /// This does not read any part payload or expose member identities.
    #[must_use]
    pub fn cache_diagnostics(&self) -> litchi_opc::SourceCacheDiagnostics {
        self.package.cache_diagnostics()
    }

    /// Load the exact raw main-document bytes as a semantic transaction
    /// snapshot.
    ///
    /// Source-backed edits currently refuse documents whose markup-
    /// compatibility preprocessing selects or rewrites branches. Keeping the
    /// raw bytes prevents semantic selectors from being applied to a different
    /// XML representation than the one eventually published.
    ///
    /// # Errors
    ///
    /// Returns a typed package, document, resource-limit, or unsafe-edit error.
    pub fn document_snapshot(&self) -> TransactionResult<Snapshot> {
        let (_, snapshot) = self.main_document_snapshot("document_snapshot")?;
        Ok(snapshot)
    }

    /// Start an isolated semantic edit over the exact raw main document.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::document_snapshot`].
    pub fn edit_document(&self) -> TransactionResult<Edit> {
        Ok(self.document_snapshot()?.edit())
    }

    /// Capture document variables from the existing internal settings Part.
    ///
    /// This source-backed capability never creates a settings relationship or
    /// Part. It accepts exactly one Strict or Transitional internal settings
    /// relationship with the Word settings content type. Markup-compatibility
    /// branch selection is refused so the semantic snapshot and publish bytes
    /// cannot diverge.
    ///
    /// # Errors
    ///
    /// Returns a typed source, relationship, content-type, MCE, XML, or limit
    /// error.
    pub fn document_variables_snapshot(&self) -> Result<variables::Snapshot> {
        Ok(self
            .settings_document_variables_source("document_variables_snapshot")?
            .snapshot)
    }

    /// Start an isolated edit of variables in the existing settings Part.
    ///
    /// # Errors
    ///
    /// Returns the same failures as [`Self::document_variables_snapshot`].
    pub fn edit_document_variables(&self) -> Result<variables::Transaction> {
        self.document_variables_snapshot()?.try_edit()
    }

    /// Publish a source-checked document-variable commit to a sequential sink.
    ///
    /// Only the existing settings payload may change. Every other physical ZIP
    /// member is raw-copied. Exact no-ops reproduce the complete source artifact
    /// byte for byte, including signatures and protected settings. A real change
    /// refuses signed, protected, MCE-selected, or unmodeled selected-variable
    /// sources before output begins.
    ///
    /// # Errors
    ///
    /// Returns a typed source-conflict, relationship, content-type, protection,
    /// preservation, signature, limit, XML-publication, or incomplete-output
    /// error.
    pub fn publish_document_variables_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &variables::Commit,
    ) -> Result<DocumentVariablesPublication> {
        let current =
            self.settings_document_variables_source("publish_document_variables_commit_to_stream")?;
        let target = commit.patch().apply(&current.snapshot)?;
        if !commit.patch().is_empty() {
            if !current.unique_inbound_owner {
                return Err(Error::DocumentVariablesPreservation(
                    "the settings Part has an additional internal inbound relationship",
                ));
            }
            if current.protected {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "publish_document_variables_commit_to_stream",
                    reason: "document or write protection is enforced",
                });
            }
            variables::ensure_source_backed_rewrite_safe(
                current.snapshot.xml_bytes(),
                current.snapshot.variables(),
            )?;
        }
        let target_xml = target.shared_xml();
        let mut replacement = Vec::new();
        replacement
            .try_reserve_exact(target_xml.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed document-variable replacement",
                source,
            })?;
        replacement.extend_from_slice(target_xml.as_slice());
        let original_artifact = self.package.source_artifact();
        let mut output = FingerprintingWriter {
            inner: writer,
            hasher: Sha256::new(),
        };
        self.package
            .write_part_overlay_to_stream(&mut output, &current.partname, replacement)?;
        let published_artifact =
            SourceArtifactFingerprint::from_sha256(output.hasher.finalize().into());
        Ok(DocumentVariablesPublication {
            snapshot: target,
            original_snapshot: commit.patch().before_snapshot().clone(),
            original_artifact,
            published_artifact,
        })
    }

    /// Apply the exact inverse of a completed source-backed publication.
    ///
    /// The supplied package must be the byte-exact artifact emitted by
    /// `publication`; a reopened foreign or subsequently modified artifact is
    /// rejected before output begins. Successful publication copies the exact
    /// retained original artifact, including all untouched physical ZIP bytes.
    ///
    /// # Errors
    ///
    /// Returns a typed conflict, source-change, allocation, or incomplete-
    /// output error.
    pub fn publish_document_variables_inverse_to_stream<W: Write>(
        self,
        writer: W,
        publication: &DocumentVariablesPublication,
    ) -> Result<variables::Snapshot> {
        let current = self
            .settings_document_variables_source("publish_document_variables_inverse_to_stream")?;
        if !current.snapshot.same_content(&publication.snapshot)
            || self.package.source_artifact().fingerprint()? != publication.published_artifact
        {
            return Err(Error::DocumentVariablesConflict);
        }
        publication.original_artifact.write_to_stream(writer)?;
        Ok(publication.original_snapshot.clone())
    }

    /// Capture the exact source closure used to detach external hyperlinks in
    /// the main document.
    ///
    /// The closure binds the raw main-document XML and its complete outbound
    /// relationship set. Enforced document or write protection is refused.
    /// Encrypted OOXML is rejected earlier because it is not a plaintext OPC
    /// package. No external target is fetched or executed.
    ///
    /// # Errors
    ///
    /// Returns a typed package, protection, markup-compatibility, XML, or
    /// resource-limit error.
    pub fn external_hyperlink_sanitization_snapshot(&self) -> Result<sanitize::Snapshot> {
        self.external_hyperlink_sanitization_snapshot_with_limits(sanitize::Limits::default())
    }

    /// Capture the external-hyperlink sanitization closure with explicit
    /// semantic scanner limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot`].
    pub fn external_hyperlink_sanitization_snapshot_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<sanitize::Snapshot> {
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_sanitization_snapshot")
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        let main = self.package.main_document_part()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(main.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink relationship closure",
                source,
            })?;
        for relationship in main.rels().iter() {
            relationships.push(RelationshipState::new(
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            ));
        }
        sanitize::Snapshot::from_source(document.shared_xml(), relationships, limits)
    }

    /// Build a non-mutating plan that detaches all external main-document
    /// hyperlink wrappers while retaining their visible child markup.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot`].
    pub fn plan_external_hyperlink_detachment(&self) -> Result<sanitize::SanitizePlan> {
        Ok(self.external_hyperlink_sanitization_snapshot()?.plan())
    }

    /// Build the external-hyperlink detachment plan with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns the same failures as
    /// [`Self::external_hyperlink_sanitization_snapshot_with_limits`].
    pub fn plan_external_hyperlink_detachment_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<sanitize::SanitizePlan> {
        Ok(self
            .external_hyperlink_sanitization_snapshot_with_limits(limits)?
            .plan())
    }

    /// Capture the exact main-document closure for explicit irreversible
    /// external-hyperlink redaction.
    ///
    /// Unlike reversible wrapper detachment, this inventory is intended for a
    /// later consuming publication that removes relationship records. It
    /// refuses external link owners outside the main document, external
    /// relationship forms outside `w:hyperlink`, protection, signatures, and
    /// unsupported owner syntax. No target is fetched or executed.
    pub fn external_hyperlink_redaction_snapshot(&self) -> Result<redact::Snapshot> {
        self.external_hyperlink_redaction_snapshot_with_limits(sanitize::Limits::default())
    }

    /// Capture an irreversible redaction inventory with explicit XML limits.
    pub fn external_hyperlink_redaction_snapshot_with_limits(
        &self,
        limits: sanitize::Limits,
    ) -> Result<redact::Snapshot> {
        let (_, document) = self
            .main_document_snapshot("external_hyperlink_redaction_snapshot")
            .map_err(transaction_error_to_document)?;
        self.refuse_protected_external_hyperlink_detachment()?;
        self.refuse_external_hyperlink_redaction_topology()?;
        let main = self.package.main_document_part()?;
        let mut relationships = Vec::new();
        relationships
            .try_reserve_exact(main.rels().len())
            .map_err(|source| Error::Allocation {
                resource: "external-hyperlink redaction relationship closure",
                source,
            })?;
        for relationship in main.rels().iter() {
            relationships.push(RelationshipState::new(
                relationship.r_id().to_owned(),
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.is_external(),
            ));
        }
        let source_version = self.package.source_version()?;
        let source_fingerprint = self.package.source_artifact().fingerprint()?;
        redact::Snapshot::from_source(
            document.shared_xml(),
            relationships,
            source_version,
            source_fingerprint,
            limits,
        )
    }

    /// Build a non-mutating forward-only plan for exact target URL values.
    pub fn plan_external_hyperlink_redaction(&self, target_urls: &[&str]) -> Result<redact::Plan> {
        self.external_hyperlink_redaction_snapshot()?
            .plan_target_urls(target_urls)
    }

    /// Publish one exact-source-checked main-document commit to a sequential
    /// stream while preserving every other physical ZIP member.
    ///
    /// Only operations confined to the main-document payload are accepted.
    /// Cross-package paragraph transfers are refused because their dependency
    /// graph requires package-level publication. A no-op commit reproduces the
    /// complete source artifact byte for byte. A changed signed package is
    /// refused by the underlying OPC publisher.
    ///
    /// All semantic, source-version, topology, signature, and replacement-XML
    /// checks happen before output. A sink failure after output begins is
    /// reported through the underlying typed incomplete-output error.
    ///
    /// # Errors
    ///
    /// Returns a typed transaction, package, unsafe-edit, signature, source,
    /// XML-publication, or sink error.
    pub fn publish_document_commit_to_stream<W: Write>(
        self,
        writer: W,
        commit: &Commit,
    ) -> TransactionResult<Snapshot> {
        let (main, current) = self.main_document_snapshot("publish_document_commit_to_stream")?;
        let target = commit.patch().apply(&current)?;
        if commit
            .patch()
            .operations()
            .iter()
            .any(|operation| !operation.supports_source_backed_main_document_overlay())
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "publish_document_commit_to_stream",
                reason: "paragraph transfers require package-level dependency publication",
            }
            .into());
        }
        self.package
            .write_part_overlay_to_stream(writer, &main, target.xml_bytes().to_vec())
            .map_err(Error::from)?;
        Ok(target)
    }

    /// Publish an explicit external-hyperlink sanitization commit to a
    /// sequential stream.
    ///
    /// Only the main-document payload is regenerated. Every other physical
    /// member and the relationship topology are preserved. Consequently the
    /// detached hyperlinks' external relationship records remain present and
    /// are counted in [`sanitize::EffectReport`]. An exact no-op reproduces
    /// every source byte, including signatures; a real change to a signed
    /// package is refused because this API has no resigning policy.
    ///
    /// All source, protection, patch, XML, signature, and preservation checks
    /// complete before output. A later sink failure is reported as the
    /// underlying typed incomplete-output error.
    ///
    /// # Errors
    ///
    /// Returns a typed package, source-conflict, protection, signature, XML,
    /// resource-limit, or sink error.
    pub fn publish_external_hyperlink_sanitization_to_stream<W: Write>(
        self,
        writer: W,
        commit: &sanitize::Commit,
    ) -> Result<sanitize::Snapshot> {
        let main = self.package.main_document_part()?.partname().clone();
        let current =
            self.external_hyperlink_sanitization_snapshot_with_limits(commit.patch().limits())?;
        let target = commit.patch().apply(&current)?;
        self.package
            .write_part_overlay_to_stream(writer, &main, target.xml_bytes().to_vec())?;
        Ok(target)
    }

    /// Irreversibly publish a source-checked external-hyperlink redaction.
    ///
    /// Selected wrappers are unwrapped, selected external relationship records
    /// are removed, visible children are retained, and every other ZIP member
    /// is raw-copied. This operation consumes the package and exposes no inverse
    /// API. It is never called by ordinary save or detachment operations.
    pub fn publish_external_hyperlink_redaction_to_stream<W: Write>(
        self,
        writer: W,
        commit: &redact::Commit,
    ) -> Result<redact::EffectReport> {
        let main = self.package.main_document_part()?.partname().clone();
        let current =
            self.external_hyperlink_redaction_snapshot_with_limits(commit.patch().limits())?;
        commit.patch().validate_source(&current)?;
        let report = commit.effect_report();
        if report.is_noop() {
            self.package.source_artifact().write_to_stream(writer)?;
            return Ok(report);
        }
        let (replacement, removed_ids) = redact::publication_parts(commit);
        self.package
            .write_part_overlay_with_external_relationship_removals_to_stream(
                writer,
                &main,
                replacement,
                removed_ids,
            )?;
        Ok(report)
    }

    fn main_document_snapshot(
        &self,
        operation: &'static str,
    ) -> TransactionResult<(litchi_opc::PackURI, Snapshot)> {
        let main = self.package.main_document_part().map_err(Error::from)?;
        validate_document_main_content_type(main.content_type())?;
        let partname = main.partname().clone();
        let raw = main.data().map_err(Error::from)?.into_arc();
        let visible = visible_document_xml(Arc::clone(&raw))?;
        if !Arc::ptr_eq(&raw, &visible) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "source-backed document transactions do not support markup-compatibility branch selection",
            }
            .into());
        }
        Ok((partname, Snapshot::from_shared_xml(raw)?))
    }

    fn settings_document_variables_source(
        &self,
        operation: &'static str,
    ) -> Result<DocumentVariablesSource> {
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        let main = self.package.main_document_part()?;
        let package_strict = self.package.rels().iter().find_map(|relationship| {
            matches!(
                relationship.reltype(),
                rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
            )
            .then_some(relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT)
        });
        let mut matching = main.rels().iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
        });
        let relationship = matching.next().ok_or_else(|| {
            Error::InvalidRelationship(
                "source-backed document-variable editing requires an existing settings relationship"
                    .into(),
            )
        })?;
        if matching.next().is_some() {
            return Err(Error::InvalidRelationship(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::InvalidRelationship(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let mut inbound_owners = 0usize;
        for candidate in self.package.rels().iter() {
            if !candidate.is_external() && candidate.target_partname()? == target {
                inbound_owners = inbound_owners.saturating_add(1);
            }
        }
        for part in self.package.iter_parts() {
            for candidate in part.rels().iter() {
                if !candidate.is_external() && candidate.target_partname()? == target {
                    inbound_owners = inbound_owners.saturating_add(1);
                }
            }
        }
        let settings_part = self.package.part(&target)?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::InvalidContentType {
                expected: ct::WML_SETTINGS.into(),
                got: settings_part.content_type().into(),
            });
        }
        let raw = settings_part.data()?.into_arc();
        if raw.len() > variables::MAX_DOCUMENT_VARIABLE_XML_BYTES {
            return Err(Error::InvalidFormat(format!(
                "settings XML exceeds the {} byte document-variable limit",
                variables::MAX_DOCUMENT_VARIABLE_XML_BYTES
            )));
        }
        let staged = BlobPart::new_shared(
            target.clone(),
            ct::WML_SETTINGS.to_owned(),
            Arc::clone(&raw),
        );
        let visible = litchi_ooxml_common::mce::process_part_arc(&staged)?;
        if !Arc::ptr_eq(&raw, &visible) {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation,
                reason: "source-backed document-variable transactions do not support markup-compatibility branch selection",
            });
        }
        let policy = variables::inspect_source_policy(raw.as_slice())?;
        let settings_strict = relationship.reltype() == STRICT_SETTINGS;
        let root_strict = policy.dialect == variables::SettingsDialect::Strict;
        if package_strict != Some(settings_strict) || settings_strict != root_strict {
            return Err(Error::InvalidRelationship(
                "main document, settings relationship, and settings XML use mixed OOXML conformance families"
                    .into(),
            ));
        }
        Ok(DocumentVariablesSource {
            partname: target,
            snapshot: variables::Snapshot::from_source_xml(raw, self.package.source_version()?)?,
            protected: policy.protected,
            unique_inbound_owner: inbound_owners == 1,
        })
    }

    fn refuse_protected_external_hyperlink_detachment(&self) -> Result<()> {
        const STRICT_SETTINGS: &str =
            "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

        let main = self.package.main_document_part()?;
        let mut settings_relationships = main.rels().iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
        });
        let Some(relationship) = settings_relationships.next() else {
            return Ok(());
        };
        if settings_relationships.next().is_some() {
            return Err(Error::InvalidRelationship(
                "document has multiple settings relationships".into(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::InvalidRelationship(
                "settings relationship cannot be external".into(),
            ));
        }
        let target = relationship.target_partname()?;
        let settings_part = self.package.part(&target)?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::InvalidContentType {
                expected: ct::WML_SETTINGS.into(),
                got: settings_part.content_type().into(),
            });
        }
        let mut staged = BlobPart::new(
            target,
            ct::WML_SETTINGS.to_owned(),
            settings_part.data()?.as_bytes().to_vec(),
        );
        for relationship in settings_part.rels().iter() {
            staged.rels_mut().try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )?;
        }
        let settings = DocumentSettings::extract_from_part(&staged)?;
        if settings.is_protected() || settings.is_write_protected() {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "external_hyperlink_wrapper_detachment",
                reason: "document or write protection is enforced",
            });
        }
        Ok(())
    }

    fn refuse_external_hyperlink_redaction_topology(&self) -> Result<()> {
        const SIGNATURE_TYPES: &[&str] = &[
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
        ];
        let main = self.package.main_document_part()?;
        if self
            .package
            .rels()
            .iter()
            .any(|relationship| SIGNATURE_TYPES.contains(&relationship.reltype()))
        {
            return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
        }
        if self
            .package
            .rels()
            .iter()
            .any(|relationship| relationship.is_external())
        {
            return Err(Error::UnsafeEdit {
                format: "DOCX",
                operation: "external_hyperlink_redaction",
                reason: "a package-level relationship has an external target outside the redaction closure",
            });
        }
        for part in self.package.iter_parts() {
            if part.partname().as_str().starts_with("/_xmlsignatures/")
                || matches!(
                    part.content_type(),
                    ct::OPC_DIGITAL_SIGNATURE_ORIGIN
                        | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                        | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
                )
                || part
                    .rels()
                    .iter()
                    .any(|relationship| SIGNATURE_TYPES.contains(&relationship.reltype()))
            {
                return Err(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy.into());
            }
            if part.partname() != main.partname()
                && part
                    .rels()
                    .iter()
                    .any(|relationship| relationship.is_external())
            {
                return Err(Error::UnsafeEdit {
                    format: "DOCX",
                    operation: "external_hyperlink_redaction",
                    reason: "a non-main-document part owns an external hyperlink relationship",
                });
            }
        }
        Ok(())
    }
}

fn transaction_error_to_document(error: crate::document::TransactionError) -> Error {
    match error {
        crate::document::TransactionError::Document(error) => error,
        other => Error::InvalidFormat(format!(
            "external hyperlink-wrapper detachment snapshot could not be captured: {other}"
        )),
    }
}

/// A pinned read-only view of a DOCX main document.
///
/// This view owns the main-document bytes loaded by [`Package::document`].
/// It intentionally exposes semantic text and paragraph queries only; use
/// the established [`crate::Package`] APIs for mutable package access.
#[derive(Clone)]
pub struct Document {
    xml: Arc<Vec<u8>>,
    source_version: litchi_core::SourceVersion,
}

impl Document {
    /// Extract all visible paragraph text from the pinned document.
    pub fn extract_text(&self) -> Result<String> {
        crate::paragraph::extract_word_text(self.xml.as_slice())
    }

    /// Count visible paragraphs in the pinned document.
    pub fn paragraph_count(&self) -> Result<usize> {
        document_paragraph_count(self.xml.as_slice())
    }

    /// Return visible paragraphs sharing the pinned main-document allocation.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        document_paragraphs(Arc::clone(&self.xml))
    }

    /// Return one visible paragraph without allocating all paragraph views.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        document_paragraph(Arc::clone(&self.xml), index)
    }

    /// Capture the immutable section inventory from the pinned main document.
    pub fn section_inventory(&self) -> Result<crate::section::Inventory> {
        self.section_inventory_with_limits(&crate::section::Limits::default())
    }

    /// Capture the section inventory with caller-provided semantic limits.
    pub fn section_inventory_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Inventory> {
        crate::section::Inventory::parse_with_limits(self.xml.as_slice(), limits)
    }

    /// Capture a cheaply cloneable inventory bound to the exact opened source.
    pub fn section_inventory_snapshot(&self) -> Result<crate::section::Snapshot> {
        self.section_inventory_snapshot_with_limits(&crate::section::Limits::default())
    }

    /// Capture a source-bound inventory with caller-provided semantic limits.
    pub fn section_inventory_snapshot_with_limits(
        &self,
        limits: &crate::section::Limits,
    ) -> Result<crate::section::Snapshot> {
        crate::section::Snapshot::from_source_xml(self.xml.as_slice(), self.source_version, limits)
    }
}
