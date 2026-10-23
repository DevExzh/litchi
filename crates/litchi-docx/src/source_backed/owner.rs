//! Logical owner boundary over a retained effective graph.

use super::*;
use litchi_opc::PreparedTopology;

impl Package {
    /// Start a bounded composition tied to this owner's complete artifact.
    /// Reopened equal-content sources may publish it, while an untouched-part
    /// change is refused before document staging.
    pub fn compose(
        &self,
        limits: litchi_core::patch::CompositionLimits,
    ) -> TransactionResult<crate::document::composition::Composition> {
        self.check_source()?;
        let source = self.document_snapshot()?;
        let fingerprint = self.output_fingerprint()?;
        Ok(source
            .compose(limits)
            .with_artifact_fingerprint(fingerprint))
    }

    /// Adopt a validated effective topology derived from this physical source.
    ///
    /// The returned owner has fresh semantic caches. The original owner and
    /// its captured views remain unchanged. This is an OPC adoption boundary;
    /// it does not replace semantic DOCX edit validation. Operations not yet
    /// migrated to effective graphs return a typed refusal.
    ///
    /// # Errors
    /// Refuses a foreign physical lineage, stale source, cancellation or an
    /// invalid effective main-document relationship/content type.
    pub fn with_prepared_topology(&self, topology: PreparedTopology) -> Result<Self> {
        self.check_source()?;
        topology.check_source()?;
        let result = (|| {
            if topology.physical_source_lineage() != self.physical_source()?.source_lineage() {
                return Err(Error::conflict(
                    "with_prepared_topology",
                    "the candidate belongs to a different physical source",
                ));
            }
            let main = topology.main_document_part()?;
            validate_document_main_content_type(main.content_type())?;
            let mut owner =
                Self::with_execution(self.physical_source()?.clone(), self.execution.clone())?
                    .with_semantic_limits(self.semantic_limits().clone());
            owner.effective = Some(topology.clone());
            owner.artifact_fingerprint = Some(ArtifactFingerprintMemo::effective());
            Ok(owner)
        })();
        self.check_source()?;
        topology.check_source()?;
        result
    }

    /// Read a paragraph's inlines using this owner's effective main-document relationships.
    /// The paragraph must have been read from this owner's main document.
    ///
    /// # Errors
    /// Returns typed XML/relationship errors and rechecks every retained source.
    pub fn paragraph_inlines(&self, paragraph: &Paragraph) -> Result<Vec<crate::Inline>> {
        self.check_source()?;
        let result = (|| match &self.effective {
            Some(topology) => {
                paragraph.inlines_with_relationships(topology.main_document_part()?.rels())
            },
            None => paragraph
                .inlines_with_relationships(self.physical_source()?.main_document_part()?.rels()),
        })();
        self.check_source()?;
        result
    }

    /// Recheck the physical archive and every retained transfer source/context.
    pub fn check_source(&self) -> Result<()> {
        if let Some(deferred) = &self.deferred {
            deferred.base.check_source()?;
        }
        if let Some(topology) = &self.effective {
            topology.check_source()?;
        }
        // A deferred owner retains the opened physical source as its
        // freshness dependency.  Calling `physical_source` here would
        // publish and reopen the complete deferred archive merely to prove
        // that dependency is still current.  Keep this fence non-materializing
        // for every apply path, including same-owner replay.  The physical
        // package and facade context are checked even for an effective owner,
        // because both remain retained dependencies of that facade.
        self.physical.source_version()?;
        self.physical.check_execution()?;
        if let Some(context) = &self.execution {
            context.check().map_err(source_document_execution_error)?;
        }
        Ok(())
    }

    /// Publish the effective package, preserving unchanged physical members.
    pub fn write_to<W: Write>(&self, mut writer: W) -> Result<()> {
        match (&self.effective, &self.deferred) {
            (Some(topology), _) => Ok(topology.publish_to_stream(writer)?),
            (None, Some(deferred)) => deferred.write(&mut writer),
            (None, None) => Ok(self
                .physical
                .clone()
                .write_topology_to_stream(writer, SourceTopologyPlan::new())?),
        }
    }

    /// Save the effective package with full durability and atomic replacement.
    /// A failure before replacement preserves the destination and removes the
    /// sibling temporary file. A post-replacement directory-sync failure remains
    /// a typed [`litchi_opc::OpcError::Committed`] error.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.save_with_durability(path, litchi_core::Durability::Full)
    }

    /// Atomically save with an explicit synchronization policy.
    /// Every level uses a sibling temporary file and rename. Only the selected
    /// file/directory syncs differ. Managed saves admit the staging buffer in the
    /// destination owner’s execution context, as are publication workspace and
    /// output. An unmanaged destination has unmanaged staging even when it
    /// retains managed donors; donor policies still govern donor reads and
    /// retention, and every donor is included in the final fence. Source state is checked again
    /// after staging and file sync, immediately before replacement.
    pub fn save_with_durability(
        &self,
        path: impl AsRef<Path>,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        self.check_source()?;
        litchi_opc::atomic::replace_guarded_with_durability(
            path.as_ref(),
            durability,
            self.execution.as_ref(),
            |sink| self.write_to(sink),
            || self.check_source(),
        )
    }

    /// Apply an exact-source main-document patch into a new immutable owner.
    ///
    /// Existing retained graph edits survive. No ZIP is emitted or reopened.
    /// Paragraph transfers still require their dedicated graph-aware publisher.
    ///
    /// # Errors
    /// Refuses stale patches, protection/tracked revisions, signatures, unsupported
    /// source provenance, transfers, source changes and resource failures.
    fn apply_package(&self, patch: &crate::document::Patch) -> TransactionResult<Self> {
        self.check_source()?;
        // A package writing exactly the commit's before artifact (a share of
        // it, or a reopened copy of its bytes) gets the commit's exact after
        // package, so the inverse of a commit restores the source archive.
        if let Some((before, after)) = patch.restore() {
            let same_semantic_limits = self.semantic_limits() == before.semantic_limits();
            // The cache identity is preserved only by package clones and
            // semantic-limit projections. Every physical/context/effective
            // boundary mints a new identity in with_execution, so this is the
            // only restored replay path that may share the committed owner
            // without rehashing or rematerializing it.
            if same_semantic_limits
                && self.same_owner(before)
                && after.semantic_limits() == self.semantic_limits()
            {
                return Ok(after.clone());
            }
            if same_semantic_limits
                && self
                    .same_artifact(before)
                    .map_err(crate::document::Error::from)?
            {
                return self.restored(after).map_err(crate::document::Error::from);
            }
        }
        let result = (|| {
            let (name, current) = self.main_document_snapshot_with_before(
                "apply",
                !patch.is_noop(),
                Some(patch.source()),
                None,
            )?;
            let target = patch.apply(&current)?;
            if patch
                .operations()
                .iter()
                .any(|operation| !operation.supports_source_backed_main_document_overlay())
            {
                return Err(Error::refused(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "apply",
                    "paragraph transfers require package-level dependency publication",
                )
                .into());
            }
            if patch.is_noop() {
                return Ok(self.clone());
            }
            document_policy::validate_changed_document(self)?;
            let topology = match &self.effective {
                Some(topology) => topology.clone(),
                None => self
                    .physical
                    .clone()
                    .prepare_topology(SourceTopologyPlan::new())
                    .map_err(Error::from)?,
            };
            let changed = if let Some(target_xml) = target.source_xml() {
                let source_xml = current.source_xml().ok_or(Error::internal(
                    "apply",
                    "current snapshot lacks XML provenance",
                ))?;
                topology
                    .apply_source_xml(&source_xml, target_xml)
                    .map_err(Error::from)?
            } else if let Some(proof) = target.publication_proof() {
                // The commit audited exactly these bytes; the overlay accepts
                // its proof instead of auditing them again (record 0311).
                topology
                    .apply_verified_xml_bytes(&name, current.xml_bytes(), proof.clone())
                    .map_err(Error::from)?
            } else {
                let bytes = target.shared_xml()?;
                topology
                    .apply_xml_bytes(&name, current.xml_bytes(), bytes)
                    .map_err(Error::from)?
            };
            self.with_prepared_topology(changed)
                .map_err(crate::document::Error::from)
        })();
        self.check_source()?;
        result
    }

    /// Apply a main-document patch and return the commit: the new package,
    /// the patch carrying this exact package pair, and diagnostics. The
    /// commit's inverse patch, applied to its package or to a reopened copy
    /// of its output, restores this package's exact artifact.
    ///
    /// # Errors
    /// Returns the failures of [`Self::apply`].
    pub fn apply(
        &self,
        patch: &crate::document::Patch,
    ) -> TransactionResult<crate::document::Commit<Self>> {
        let package = self.apply_package(patch)?;
        let diagnostics =
            crate::document::Diagnostics::new(patch.operations().len(), !patch.is_noop());
        let patch = patch.clone().with_restore(self.clone(), package.clone());
        Ok(crate::document::Commit::from_parts(
            package,
            patch,
            diagnostics,
        ))
    }

    pub(super) fn effective_document_snapshot(
        &self,
        topology: &PreparedTopology,
        operation: &'static str,
        source_authorized: bool,
        before: Option<&Snapshot>,
        edit_options: Option<&crate::document::EditOptions>,
    ) -> TransactionResult<(PackURI, Snapshot)> {
        self.check_source()?;
        let result = (|| {
            // Raw snapshots retain at most the base execution context. They
            // cannot carry or check an independent donor's budget/context.
            if topology.has_managed_donors() {
                return Err(
                    Error::from(litchi_opc::OpcError::ManagedPackageMaterialization).into(),
                );
            }
            // Exact-artifact no-ops have no logical overlay. Preserve the
            // physical path's signed managed snapshot fallback without
            // weakening the guard on any changed effective owner.
            if topology.is_noop() {
                let physical =
                    Self::with_execution(self.physical_source()?.clone(), self.execution.clone())?;
                return physical.main_document_snapshot_with_before(
                    operation,
                    source_authorized,
                    before,
                    edit_options,
                );
            }

            let main = topology.main_document_part().map_err(Error::from)?;
            validate_document_main_content_type(main.content_type())?;
            if let Some(options) = edit_options
                && options.max_snapshot_bytes_value().is_some()
            {
                options
                    .admit_source_bytes(main.declared_uncompressed_size().map_err(Error::from)?)?;
            }
            let name = main.partname().clone();
            let lineage = topology.physical_source_lineage();
            let version = topology.physical_source_version().map_err(Error::from)?;
            let snapshot = if let Some(context) = &self.execution {
                self.paragraph_index_cache.trim_clean();
                let xml = main.source_xml().map_err(Error::from)?;
                if let Some(reused) = before.and_then(|before| {
                    before.reuse_if_source_xml_matches(&xml, &lineage, version, &name)
                }) {
                    reused
                } else {
                    self.ensure_source_document_xml(xml.bytes(), Some(context))?;
                    Snapshot::from_source_xml(xml, lineage, version, &name, context.clone())?
                }
            } else {
                let data = main.data().map_err(Error::from)?;
                self.admit_main_document(data.as_bytes(), version, operation)?;
                // A committed handle edited again scans its effective revision
                // once, as the physical route does (record 0311).
                if let Some(snapshot) =
                    self.memoized_main_snapshot(data.as_bytes(), &lineage, version, &name)
                {
                    let snapshot =
                        self.with_retained_main_proof(snapshot, || main.xml_audit_proof());
                    return Ok((name, snapshot));
                }
                // The revision's retained proof shares the effective bytes'
                // allocation, so the snapshot shares it too instead of copying
                // it, and carries the proof (record 0317).
                let retained = main
                    .xml_audit_proof()
                    .filter(|proof| core::ptr::eq(proof.bytes().as_slice(), data.as_bytes()));
                let bytes = match &retained {
                    Some(proof) => Arc::clone(proof.bytes()),
                    None => {
                        let mut bytes = Vec::new();
                        bytes
                            .try_reserve_exact(data.as_bytes().len())
                            .map_err(|source| {
                                Error::Opc(litchi_opc::OpcError::Allocation {
                                    resource: "effective document snapshot",
                                    source,
                                })
                            })?;
                        bytes.extend_from_slice(data.as_bytes());
                        Arc::new(bytes)
                    },
                };
                let snapshot =
                    Snapshot::from_shared_xml_with_source_identity(bytes, lineage, version, &name)?;
                self.remember_main_snapshot(&snapshot);
                self.with_retained_main_proof(snapshot, || retained)
            };
            Ok((name, snapshot))
        })();
        self.check_source()?;
        result
    }

    pub(super) fn effective_document(&self, topology: &PreparedTopology) -> Result<Document> {
        topology.check_source()?;
        let result = (|| {
            let main = topology.main_document_part()?;
            validate_document_main_content_type(main.content_type())?;
            let data = match main.data().map_err(Error::from) {
                Ok(data) => data,
                Err(error) if is_index_pressure_error(&error) => {
                    self.paragraph_index_cache.trim_clean();
                    main.data()?
                },
                Err(error) => return Err(error),
            };
            let source_version = topology.physical_source_version()?;
            let (xml, paragraph_memo, index_admission, paragraph_route, retained, positions) =
                if let Some(context) = &self.execution {
                    let build = self.paragraph_index_cache.get_or_build_source(
                        source_version,
                        ParagraphIndexRoute::Managed,
                        || {
                            ensure_source_document_xml(data.as_bytes(), Some(context))?;
                            let admission =
                                DocumentIndexAdmission::new(context, data.as_bytes().len())?;
                            let _parser =
                                admit_document_query_parser(Some(context), data.as_bytes().len())?;
                            let index = ParagraphIndex::from_xml(data.as_bytes()).ok();
                            topology.check_source()?;
                            Ok(match index {
                                Some(index) => ParagraphIndexBuild::indexed(index, Some(admission)),
                                None => ParagraphIndexBuild::unavailable(Some(admission)),
                            })
                        },
                    )?;
                    (
                        DocumentPayload::Effective(data.clone()),
                        OnceLock::from(build.memo),
                        build.admission,
                        ParagraphIndexRoute::Managed,
                        Some(data),
                        RawPositions::Preserved,
                    )
                } else {
                    let mut bytes = Vec::new();
                    bytes
                        .try_reserve_exact(data.as_bytes().len())
                        .map_err(|source| litchi_opc::OpcError::Allocation {
                            resource: "effective DOCX visible XML",
                            source,
                        })?;
                    bytes.extend_from_slice(data.as_bytes());
                    let raw = Arc::new(bytes);
                    let payload = VisiblePayload::project(raw.as_slice())?;
                    let xml = payload.xml(raw);
                    (
                        DocumentPayload::Owned(xml),
                        OnceLock::new(),
                        None,
                        ParagraphIndexRoute::Visible,
                        Some(data),
                        payload.positions(),
                    )
                };
            Ok(Document {
                xml,
                _retained: retained,
                paragraph_memo,
                _index_admission: index_admission,
                source_version,
                execution: self.execution.clone(),
                limits: Arc::new(self.semantic_limits().clone()),
                paragraph_index_cache: Arc::clone(&self.paragraph_index_cache),
                paragraph_route,
                positions,
                body_memo: OnceLock::new(),
                allows_detached_views: !topology.is_managed(),
                tolerance: Arc::new(Mutex::new(self.opened_tolerance())),
            })
        })();
        match topology.check_source().map_err(Error::from) {
            Ok(()) => {
                if result.is_err() {
                    self.paragraph_index_cache.trim_clean();
                }
                result
            },
            Err(error) => {
                drop(result);
                self.paragraph_index_cache.trim_clean();
                Err(error)
            },
        }
    }
}

#[allow(
    clippy::large_enum_variant,
    reason = "inline retained payloads avoid an unadmitted allocation"
)]
pub(super) enum MainTextData {
    Physical(PartData),
    Effective(litchi_opc::PreparedPartData),
}

impl MainTextData {
    pub(super) fn as_bytes(&self) -> &[u8] {
        match self {
            Self::Physical(data) => data.as_bytes(),
            Self::Effective(data) => data.as_bytes(),
        }
    }
}

impl Package {
    pub(super) fn main_text_data(&self) -> Result<MainTextData> {
        let preflight = |content_type: &str, declared: u64| -> Result<()> {
            validate_document_main_content_type(content_type)?;
            let limit = u64::try_from(crate::paragraph::semantic_text_raw_xml_limit())
                .map_err(|_| Error::InvalidFormat("semantic DOCX XML limit overflow".into()))?;
            if declared > limit {
                return Err(Error::InvalidFormat(format!(
                    "semantic DOCX declared XML exceeds {limit} bytes"
                )));
            }
            Ok(())
        };
        match &self.effective {
            Some(topology) => {
                let main = topology.main_document_part()?;
                preflight(main.content_type(), main.declared_uncompressed_size()?)?;
                let data = match main.data().map_err(Error::from) {
                    Ok(data) => data,
                    Err(error) if is_index_pressure_error(&error) => {
                        self.paragraph_index_cache.trim_clean();
                        main.data()?
                    },
                    Err(error) => return Err(error),
                };
                Ok(MainTextData::Effective(data))
            },
            None => {
                let main = self.physical_source()?.main_document_part()?;
                preflight(main.content_type(), main.declared_uncompressed_size()?)?;
                Ok(MainTextData::Physical(self.main_data(main)?))
            },
        }
    }
}
