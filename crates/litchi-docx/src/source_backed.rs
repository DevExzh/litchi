//! Lazy DOCX access and guarded main-document publication over an immutable
//! positional source.
//!
//! [`Package::from_read_at`] validates the OPC package, its main-document
//! relationship, and the main-document content type without materializing the
//! main document. [`Package::document`] performs that first payload read and
//! returns a pinned semantic view which owns the loaded bytes. Main-document
//! transactions retain the raw XML and may be published to a sequential sink
//! while raw-copying every unselected ZIP member.

use crate::document::{Commit, Edit, Snapshot, TransactionResult};
use crate::error::{Error, Result};
use crate::package::validate_document_main_content_type;
use crate::paragraph::Paragraph;
use crate::parts::document_part::{
    document_paragraph, document_paragraph_count, document_paragraphs, visible_document_xml,
};
use crate::sanitize::{self, RelationshipState};
use crate::settings::DocumentSettings;
use litchi_core::ReadAt;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, Part, SourceBackedPackage};
use smallvec::SmallVec;
use std::io::Write;
use std::sync::Arc;

/// A DOCX package that leaves ordinary part bodies cold at open.
pub struct Package {
    package: SourceBackedPackage,
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
        Ok(Document { xml })
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
}
