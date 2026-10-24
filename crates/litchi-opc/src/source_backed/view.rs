//! Owned read handles for an immutable effective topology.

use super::*;

/// Exact decoded candidate metadata, retaining its source or overlay admission.
///
/// Clones share bytes and retain the complete prepared candidate, including
/// managed reservations and donor contexts. Captured bytes remain readable
/// after source changes; acquiring a new handle checks source and execution
/// guards again. This is a metadata image, not authority to replace metadata.
#[derive(Clone)]
pub struct PreparedMetadataData {
    topology: PreparedTopology,
    payload: PreparedMetadataPayload,
}

#[derive(Clone)]
enum PreparedMetadataPayload {
    Source(PartData),
    Changed(usize),
    Appended(usize),
}

impl std::fmt::Debug for PreparedMetadataData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PreparedMetadataData")
            .field("bytes", &self.as_bytes().len())
            .finish()
    }
}

impl PreparedMetadataData {
    /// Borrow the exact decoded metadata bytes captured by this handle.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        match &self.payload {
            PreparedMetadataPayload::Source(data) => data.as_bytes(),
            PreparedMetadataPayload::Changed(index) => {
                self.topology.state.changed[*index].replacement.as_slice()
            },
            PreparedMetadataPayload::Appended(index) => {
                self.topology.state.retained().policy_appended[*index]
                    .payload
                    .as_slice()
            },
        }
    }

    /// Recheck the execution contexts retained by this captured metadata.
    pub fn check_execution(&self) -> Result<()> {
        self.topology.check_execution()
    }
}

/// An owned, lazy Part view retaining its complete effective package.
///
/// Cloning shares the candidate without reading or copying payloads. Views
/// remain attached to their captured revision after later edits or owner drop.
/// Metadata accessors return this captured metadata without fresh source or
/// cancellation checks; deferred data and proof acquisition recheck both.
#[derive(Clone)]
pub struct PreparedPartView {
    topology: PreparedTopology,
    pub(super) index: usize,
}

/// Immutable decoded bytes retaining their payload and managed admissions.
///
/// This handle does not expose a detachable Arc. Already captured bytes remain
/// readable if the physical source later changes; acquiring a new payload or
/// source proof through the Part view rechecks physical and transfer sources.
/// Each handle intentionally pins the complete candidate, including unrelated
/// overlays and transfer monitors, until its final share is dropped.
#[derive(Clone)]
pub struct PreparedPartData {
    part: PreparedPartView,
    source: Option<PartData>,
}

impl std::fmt::Debug for PreparedPartView {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PreparedPartView").finish_non_exhaustive()
    }
}

impl std::fmt::Debug for PreparedPartData {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PreparedPartData")
            .field("bytes", &self.as_bytes().len())
            .finish()
    }
}

impl PreparedTopology {
    /// Capture the exact effective relationship member for a Part or `/`.
    ///
    /// `None` means the member is absent, including after owner removal. A
    /// physically present empty relationship file yields `Some`. Changed and
    /// appended bytes are shared with publication; unchanged metadata is read
    /// lazily under the source's managed payload policy. No other Part is read.
    ///
    /// # Errors
    /// Returns source, execution, invalid owner URI, or metadata limit errors.
    pub fn relationships_data_for_with_limit(
        &self,
        owner: &PackURI,
        maximum: usize,
    ) -> Result<Option<PreparedMetadataData>> {
        self.check_source()?;
        // PackURI identity is case-equivalent, while ZIP member lookup is
        // exact. Derive the member name from the admitted owner spelling,
        // including removed physical owners and newly added candidate owners.
        let owner = self
            .state
            .effective
            .as_ref()
            .and_then(|effective| {
                effective.parts.iter().find_map(|part| {
                    let name = part.partname(&self.package);
                    name.is_equivalent_to(owner).then_some(name)
                })
            })
            .or_else(|| {
                self.package
                    .part_index(owner)
                    .map(|index| &self.package.catalog.parts[index].partname)
            })
            .unwrap_or(owner);
        let uri = owner.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let maximum = maximum.min(self.package.limits.max_relationship_xml_bytes());
        self.capture_metadata(
            uri.membername(),
            maximum,
            SourceLimitKind::RelationshipXmlBytes,
            || {
                self.package
                    .relationships_data_for_with_limit(owner, maximum)
            },
        )
    }

    /// Capture the exact effective content-types manifest under a byte ceiling.
    ///
    /// The image matches decoded publication bytes, preserving lexical trivia.
    /// This does not regenerate XML from the effective typed content-type map.
    ///
    /// # Errors
    /// Returns source, execution, missing manifest, or metadata limit errors.
    pub fn content_types_data_with_limit(&self, maximum: usize) -> Result<PreparedMetadataData> {
        let maximum = maximum.min(self.package.limits.max_content_types_bytes());
        self.capture_metadata(
            &self.package.catalog.content_types_member,
            maximum,
            SourceLimitKind::ContentTypesBytes,
            || {
                self.package
                    .content_types_data_with_limit(maximum)
                    .map(Some)
            },
        )?
        .ok_or_else(|| OpcError::PartNotFound(self.package.catalog.content_types_member.clone()))
    }

    fn capture_metadata(
        &self,
        member: &str,
        maximum: usize,
        kind: SourceLimitKind,
        source: impl FnOnce() -> Result<Option<PartData>>,
    ) -> Result<Option<PreparedMetadataData>> {
        self.read_capture(|| {
            let retained = self.state.retained.as_ref();
            if retained.is_some_and(|state| {
                state
                    .omitted_members
                    .iter()
                    .any(|name| name.eq_ignore_ascii_case(member))
            }) {
                return Ok(None);
            }
            let changed = self.state.changed.iter().position(|overlay| {
                matches!(&overlay.target, ChangedOverlayTarget::Member(name)
                    if name.eq_ignore_ascii_case(member))
            });
            let appended = retained.and_then(|state| {
                state
                    .policy_appended
                    .iter()
                    .position(|entry| entry.name.eq_ignore_ascii_case(member))
            });
            let payload = if let Some(index) = changed {
                PreparedMetadataPayload::Changed(index)
            } else if let Some(index) = appended {
                PreparedMetadataPayload::Appended(index)
            } else if let Some(data) = source()? {
                PreparedMetadataPayload::Source(data)
            } else {
                return Ok(None);
            };
            let data = PreparedMetadataData {
                topology: self.clone(),
                payload,
            };
            self.package
                .limits
                .check(kind, data.as_bytes().len() as u64, maximum as u64)?;
            Ok(Some(data))
        })
    }

    fn read_capture<T>(&self, capture: impl FnOnce() -> Result<T>) -> Result<T> {
        self.check_source()?;
        let result = capture();
        self.check_source()?;
        result
    }

    /// Recheck the physical base, transfer sources and their execution contexts.
    ///
    /// # Errors
    /// Returns changed-source or execution failures, including cancellation.
    pub fn check_source(&self) -> Result<()> {
        self.state.check_readback_sources(&self.package)
    }

    /// Check all retained execution contexts without re-observing source bytes.
    ///
    /// Pinned semantic views use this after capture: immutable bytes remain
    /// readable after a physical source change, but cancellation and deadlines
    /// on either the base or a managed donor still govern new semantic work.
    pub fn check_execution(&self) -> Result<()> {
        self.package.check_execution()?;
        if let Some(sources) = self
            .state
            .retained
            .as_ref()
            .and_then(|retained| retained.transfer_sources.as_ref())
        {
            for source in sources.iter() {
                if let Some(context) = &source.context {
                    context.check().map_err(map_execution_error)?;
                }
            }
        }
        Ok(())
    }

    /// Whether publication uses the exact original-artifact path.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.state.no_op
    }

    /// Limits captured by the original source and enforced on candidates.
    #[must_use]
    pub const fn read_limits(&self) -> SourceLimits {
        self.package.read_limits()
    }

    /// Whether the physical base or any retained donor requires managed retention.
    ///
    /// An unmanaged base can retain managed donors. Such a candidate must not
    /// detach payloads into unaccounted owning representations.
    #[must_use]
    pub fn is_managed(&self) -> bool {
        self.package.cache.is_managed() || self.has_managed_donors()
    }

    /// Whether any retained donor carries an execution context.
    ///
    /// Payload consumers must retain and check these contexts even when the
    /// physical base has its own separate execution context.
    #[must_use]
    pub fn has_managed_donors(&self) -> bool {
        self.state
            .retained
            .as_ref()
            .and_then(|retained| retained.transfer_sources.as_ref())
            .is_some_and(|sources| {
                sources.iter().any(|source| {
                    source.context.is_some() && source.lineage != self.package.source.lineage
                })
            })
    }

    /// Share the explicit execution context, when this owner is managed.
    #[must_use]
    pub fn execution_context(&self) -> Option<ExecutionContext> {
        self.package.execution_context()
    }

    /// The physical archive lineage, distinct from a logical candidate revision.
    ///
    /// This is not a key for caching edited XML or its semantic projections.
    #[must_use]
    pub fn physical_source_lineage(&self) -> SourceLineage {
        self.package.source_lineage()
    }

    /// Observe the unchanged physical source version after checking all guards.
    ///
    /// Logical edits do not change this version; do not use it as a candidate
    /// cache key.
    ///
    /// # Errors
    /// Returns changed-source or execution failures.
    pub fn physical_source_version(&self) -> Result<SourceVersion> {
        self.read_capture(|| self.package.source_version())
    }

    /// Retain the original physical artifact, which excludes candidate edits.
    ///
    /// Use [`Self::publish_to_stream`] to publish the effective package.
    #[must_use]
    pub fn physical_source_artifact(&self) -> SourceArtifact {
        self.package.source_artifact()
    }

    /// Resolve an owned view from the effective catalog without reading XML.
    ///
    /// Unlike commit-time [`Self::with_candidate`], ordinary reads also work
    /// for an unchanged candidate. A removed Part is absent and an added Part
    /// resolves to its retained payload, never to the physical source catalog.
    ///
    /// ```
    /// use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter,
    ///     SourceBackedPackage, SourceCacheLimits, SourceLimits, SourceTopologyPlan};
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let name = PackURI::new("/example.xml")?;
    /// let mut package = OpcPackage::new();
    /// package.try_add_part(Box::new(BlobPart::new(
    ///     name.clone(), "application/xml".to_owned(), b"<before/>".to_vec(),
    /// )))?;
    /// let source = SourceBackedPackage::from_vec(PackageWriter::to_bytes(&package)?, SourceLimits::default(), SourceCacheLimits::default(), None)?;
    /// let mut plan = SourceTopologyPlan::new();
    /// plan.try_replace_part(name.clone(), b"<after/>".to_vec())?;
    /// let candidate = source.prepare_topology(plan)?;
    /// let part = candidate.part(&name)?;
    /// drop(candidate);
    /// assert_eq!(part.content_type(), "application/xml");
    /// assert_eq!(part.data()?.as_bytes(), b"<after/>");
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns missing-Part, changed-source or execution-context failures.
    pub fn part(&self, partname: &PackURI) -> Result<PreparedPartView> {
        self.state.check_readback_sources(&self.package)?;
        let index = if self.state.no_op {
            self.package.part_index(partname)
        } else {
            self.state
                .effective_data()
                .parts
                .iter()
                .position(|part| part.partname(&self.package).is_equivalent_to(partname))
        };
        self.state.check_readback_sources(&self.package)?;
        let index = index.ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
        Ok(PreparedPartView {
            topology: self.clone(),
            index,
        })
    }

    /// Iterate owned metadata views in deterministic source/addition order.
    ///
    /// The iterator owns its candidate and can outlive this handle. Resolving
    /// views reads no payload; [`PreparedPartView::data`] performs lazy reads.
    ///
    /// # Errors
    /// Returns changed-source or execution-context failures at capture.
    pub fn iter_parts(&self) -> Result<impl ExactSizeIterator<Item = PreparedPartView> + use<>> {
        self.state.check_readback_sources(&self.package)?;
        let count = if self.state.no_op {
            self.package.catalog.parts.len()
        } else {
            self.state.effective_data().parts.len()
        };
        let topology = self.clone();
        self.check_source()?;
        Ok((0..count).map(move |index| PreparedPartView {
            topology: topology.clone(),
            index,
        }))
    }

    /// Borrow captured effective package-root relationships without fresh checks.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        if self.state.no_op {
            self.package.rels()
        } else {
            self.state
                .effective_data()
                .package_relationships
                .as_deref()
                .unwrap_or_else(|| self.package.rels())
        }
    }

    /// Resolve the unique effective main document relationship to an owned Part.
    ///
    /// # Errors
    /// Refuses missing, duplicate, external or dangling main relationships,
    /// changed sources and execution failures.
    pub fn main_document_part(&self) -> Result<PreparedPartView> {
        self.state.check_readback_sources(&self.package)?;
        let name = self.read_capture(|| main_document_partname(self.rels()))?;
        self.part(&name)
    }
}

impl PreparedPartView {
    fn effective(&self) -> Option<&PreparedPart> {
        (!self.topology.state.no_op)
            .then(|| self.topology.state.effective_data().parts[self.index].as_ref())
    }

    /// Return the captured effective Part URI.
    #[must_use]
    pub fn partname(&self) -> &PackURI {
        match self.effective() {
            Some(part) => part.partname(&self.topology.package),
            None => &self.topology.package.catalog.parts[self.index].partname,
        }
    }

    /// Return the captured effective content type.
    #[must_use]
    pub fn content_type(&self) -> &str {
        match self.effective() {
            Some(part) => part.content_type(&self.topology.package),
            None => self.topology.package.catalog.parts[self.index]
                .content_type
                .as_str(),
        }
    }

    /// Borrow the captured effective relationships owned by this Part.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        match self.effective() {
            Some(part) => part.relationships(&self.topology.package),
            None => &self.topology.package.catalog.parts[self.index].relationships,
        }
    }

    /// The source XML audit's retained proof of this Part's effective bytes,
    /// when one is kept; nothing is audited (record 0317).
    ///
    /// A replaced Part answers with the proof its replacement's audit
    /// returned, which an unmanaged revision keeps for at most four Parts. An
    /// unchanged physical Part answers with the indexed proof its source
    /// bytes earned when the package last audited a replacement of them
    /// ([`SourceBackedPackage::xml_audit_proof`]). The proof covers exactly
    /// the bytes [`Self::data`] returns, so a caller may hand it to a source
    /// XML audit of an edit of them, such as
    /// [`xml_minifier::audit::VerifiedSource::verify_successor`], whose
    /// verdict is the edit's own complete audit either way.
    #[must_use]
    pub fn xml_audit_proof(&self) -> Option<xml_minifier::audit::VerifiedSource> {
        match self.effective() {
            Some(part) if part.replacement.is_some() => part.xml_proof.clone(),
            Some(part) => part
                .source_index
                .and_then(|index| self.topology.package.retained_xml_audit_proof(index)),
            None => self.topology.package.retained_xml_audit_proof(self.index),
        }
    }

    /// Lazily capture immutable payload bytes without detaching admissions.
    ///
    /// Prepared payloads share their existing allocation; unchanged payloads
    /// use the physical package cache. Source/execution fences run on both sides.
    ///
    /// # Errors
    /// Returns deferred read, changed-source, limit or execution failures.
    pub fn data(&self) -> Result<PreparedPartData> {
        self.topology.read_capture(|| {
            let source = match self.effective() {
                Some(part) if part.replacement.is_some() => None,
                part => {
                    let index = part
                        .map_or(Some(self.index), |part| part.source_index)
                        .ok_or_else(|| {
                            overlay_unavailable("effective Part has no payload source")
                        })?;
                    Some(self.topology.package.read_part(index)?)
                },
            };
            Ok(PreparedPartData {
                part: self.clone(),
                source,
            })
        })
    }

    /// Read the effective payload length without loading unchanged bytes.
    ///
    /// A physical value is an untrusted ZIP declaration until [`Self::data`]
    /// verifies it. A replacement or addition reports its retained byte length.
    ///
    /// # Errors
    /// Returns metadata, source or execution failures.
    pub fn declared_uncompressed_size(&self) -> Result<u64> {
        self.topology.read_capture(|| match self.effective() {
            Some(part) if part.replacement.is_some() => Ok(part
                .replacement
                .as_ref()
                .expect("replacement is present")
                .len() as u64),
            part => {
                let index = part
                    .map_or(Some(self.index), |part| part.source_index)
                    .ok_or_else(|| overlay_unavailable("effective Part has no payload source"))?;
                Ok(self
                    .topology
                    .package
                    .archive
                    .metadata_for(self.topology.package.catalog.parts[index].entry_id)?
                    .uncompressed_size())
            },
        })
    }

    /// Capture a fresh logical XML revision from the effective source Part.
    ///
    /// Physical original bytes remain the publication authority. Replacements
    /// must retain checked XML provenance; additions and unproven authored
    /// replacements are refused rather than borrowing stale physical bytes.
    ///
    /// # Errors
    /// Returns provenance, XML, source, admission or execution failures.
    pub fn source_xml(&self) -> Result<SourceXmlPart> {
        self.topology.read_capture(|| {
            let source = match self.effective() {
                Some(part) => {
                    let physical = part.source_index.ok_or_else(|| {
                        overlay_unavailable(
                            "logical XML capture requires an existing physical source Part",
                        )
                    })?;
                    match (&part.replacement, &part.source_xml) {
                        (Some(bytes), Some(source)) if bytes.as_slice() == source.bytes() => {
                            source.clone()
                        },
                        (Some(_), _) => {
                            return Err(overlay_unavailable(
                                "effective XML replacement has no retained source provenance",
                            ));
                        },
                        (None, _) => self.topology.package.source_xml_part(physical)?,
                    }
                },
                None => self.topology.package.source_xml_part(self.index)?,
            };
            source.begin_revision()
        })
    }
}

impl PreparedPartData {
    /// Check base and donor execution contexts while retaining captured bytes.
    /// This does not re-observe the physical source version.
    pub fn check_execution(&self) -> Result<()> {
        self.part.topology.check_execution()
    }

    /// Borrow the exact decoded bytes captured by this immutable handle.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        match &self.source {
            Some(source) => source.as_bytes(),
            None => self
                .part
                .effective()
                .and_then(|part| part.replacement.as_ref())
                .expect("prepared payload capture retains its replacement")
                .as_slice(),
        }
    }
}
