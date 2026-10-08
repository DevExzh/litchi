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
        let result = self.materialize_effective();
        self.check_source()?;
        result
    }

    fn materialize_effective(&self) -> Result<OpcPackage> {
        let managed_transfer = self
            .state
            .retained
            .as_ref()
            .and_then(|retained| retained.transfer_sources.as_ref())
            .is_some_and(|sources| sources.iter().any(|source| source.context.is_some()));
        if self.package.cache.is_managed() || managed_transfer {
            return Err(OpcError::ManagedPackageMaterialization);
        }
        if self.state.no_op {
            return self.package.to_opc_package();
        }
        self.package.disable_read_ahead_for_publication()?;
        let non_part_members = self.package.cloned_non_part_members()?;
        let mut package = OpcPackage::new();
        copy_relationships(self.rels(), package.rels_mut())?;
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
            package.try_add_source_part(part)?;
        }
        package.set_non_part_members(non_part_members);
        package.set_tolerance_report(self.package.catalog.tolerance_report.clone());
        package.mark_source_ingress_signature_policy();
        Ok(package)
    }
}
