//! Explicit owning conversion of the effective graph.

use super::*;

impl PreparedTopology {
    /// Materialize the effective catalog and payloads into an owning package.
    ///
    /// Replacements, additions, removals, content types and relationships come
    /// from this candidate, not the original archive. Unchanged payloads use
    /// one serial decoder session; retained replacements share their immutable
    /// allocation. This is an explicit whole-package operation, not an edit
    /// or publication fast path. Exact physical ZIP framing is not preserved
    /// by the owning representation; use [`Self::publish_to_stream`] for that.
    ///
    /// Managed base or transfer sources are refused before payload reads so
    /// their reservations cannot escape into an unaccounted owning package.
    /// All source and execution guards are checked before and after conversion,
    /// including when a decoding or construction step fails. Signed unchanged
    /// packages retain the owning layer's borrowed-ingress signature policy.
    ///
    /// # Errors
    /// Returns [`OpcError::ManagedPackageMaterialization`] for managed sources,
    /// or a source, execution, allocation, payload or relationship error.
    pub fn to_opc_package(&self) -> Result<OpcPackage> {
        self.check_source()?;
        let result = self.materialize_effective(false);
        self.check_source()?;
        result
    }

    /// Materialize the effective graph while retaining a bounded exact source
    /// archive for later owning publication.
    ///
    /// This is the explicit Eager conversion seam for a prepared candidate.
    /// The source-backed owner shares an existing in-memory archive or copies
    /// a positional source once under its input limit; unchanged members can
    /// then remain byte-for-byte preserved after the candidate is adopted.
    pub fn to_opc_package_preserving_source(&self) -> Result<OpcPackage> {
        self.check_source()?;
        let result = self.materialize_effective(true);
        self.check_source()?;
        result
    }

    fn materialize_effective(&self, preserve_source: bool) -> Result<OpcPackage> {
        if self.is_managed() {
            return Err(OpcError::ManagedPackageMaterialization);
        }
        if self.state.no_op {
            return if preserve_source {
                self.package.to_opc_package_preserving_source()
            } else {
                self.package.to_opc_package()
            };
        }
        self.package.disable_read_ahead_for_publication()?;
        let source = self.package.finish_stage(if preserve_source {
            self.package.materialization_source_shared().map(Some)
        } else {
            Ok(self.package.owned_source_shared())
        })?;
        let non_part_members = self.package.cloned_non_part_members()?;
        let mut package = OpcPackage::new();
        copy_relationships(self.rels(), package.rels_mut())?;
        self.check_source()?;
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(self.state.effective_data().parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC prepared materialized parts",
                source,
            })?;
        self.check_source()?;
        let mut session = self.package.archive.read_session();
        for effective in &self.state.effective_data().parts {
            self.check_source()?;
            let bytes = match &effective.replacement {
                Some(bytes) => Arc::clone(bytes),
                None => {
                    let index = effective.source_index.ok_or_else(|| {
                        overlay_unavailable("effective Part has neither source nor payload")
                    })?;
                    self.package
                        .read_part_with_session(index, &mut session, None)?
                        .into_arc()?
                },
            };
            self.check_source()?;
            let mut part = PartFactory::load_shared(
                effective.partname(&self.package).clone(),
                effective.content_type(&self.package).to_owned(),
                bytes,
            )?;
            copy_relationships(effective.relationships(&self.package), part.rels_mut())?;
            // A replacement is authored by the candidate. Keep it out of the
            // source XML provenance map so the preservation writer cannot
            // mistake the replacement allocation for the untouched source
            // member. Retained source parts keep their exact XML allocation.
            parts.push((part, effective.replacement.is_none()));
        }
        package.try_add_materialized_parts(parts, || self.check_source())?;
        package.set_non_part_members(non_part_members);
        package.set_tolerance_report(self.package.catalog.tolerance_report.clone());
        if let Some(source) = source {
            self.package
                .attach_owned_source_preservation(&mut package, false, source, None)?;
        }
        package.mark_source_ingress_signature_policy();
        Ok(package)
    }
}
