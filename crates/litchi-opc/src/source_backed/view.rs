//! Owned read handles for an immutable effective topology.

use super::*;

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

    /// Whether publication uses the exact original-artifact path.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.state.no_op
    }

    /// Limits captured by the original source and enforced on candidates.
    #[must_use]
    pub const fn read_limits(&self) -> ReadLimits {
        self.package.read_limits()
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
    ///     SourceBackedPackage, SourceTopologyPlan};
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let name = PackURI::new("/example.xml")?;
    /// let mut package = OpcPackage::new();
    /// package.try_add_part(Box::new(BlobPart::new(
    ///     name.clone(), "application/xml".to_owned(), b"<before/>".to_vec(),
    /// )))?;
    /// let source = SourceBackedPackage::from_vec(PackageWriter::to_bytes(&package)?)?;
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
