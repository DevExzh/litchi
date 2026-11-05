//! Immutable XML revisions of a retained effective topology.

use super::*;
use crate::xml_splice::XmlSplicePublication;

impl PreparedTopology {
    /// Edit an existing source XML Part against this retained candidate.
    ///
    /// The callback receives a revision-specific source and a checked splice
    /// transaction. Success returns a new flat candidate over the original
    /// archive; failure leaves this owner unchanged. No ZIP is materialized.
    /// Exact logical no-ops return a share of this candidate, preserving prior
    /// topology edits. Source-backed XML replacements retain their provenance.
    /// Added Parts and authored replacements without source provenance are not
    /// yet supported by this low-level source revision entry.
    ///
    /// ```
    /// use litchi_opc::{AuthoredXmlFragment, BlobPart, OpcPackage, PackURI,
    ///     PackageWriter, SourceBackedPackage, SourceCacheLimits, SourceLimits,
    ///     SourceTopologyPlan};
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let name = PackURI::new("/example.xml")?;
    /// let mut package = OpcPackage::new();
    /// package.try_add_part(Box::new(BlobPart::new(
    ///     name.clone(), "application/xml".to_owned(), b"<first/>".to_vec(),
    /// )))?;
    /// let source = SourceBackedPackage::from_vec(PackageWriter::to_bytes(&package)?, SourceLimits::default(), SourceCacheLimits::default(), None)?;
    /// let mut current = source.prepare_topology(SourceTopologyPlan::new())?;
    /// for next in [b"<second/>".as_slice(), b"<third/>".as_slice()] {
    ///     current = current.edit_xml(&name, |xml, edit| {
    ///         edit.replace(xml.checked_range(0..xml.bytes().len(), xml.bytes())?,
    ///             AuthoredXmlFragment::markup(next.to_vec())?)
    ///     })?;
    /// }
    /// current.with_candidate(|view| {
    ///     assert_eq!(view.part(&name)?.data()?.as_bytes(), b"<third/>");
    ///     Ok(())
    /// })?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Refuses missing or unproven Parts, foreign transactions, invalid XML,
    /// changed sources, cancellation and admission failures.
    pub fn edit_xml<F>(&self, partname: &PackURI, edit: F) -> Result<Self>
    where
        F: FnOnce(&SourceXmlPart, &mut XmlSplicePublication) -> Result<()>,
    {
        self.state.check_readback_sources(&self.package)?;
        if let Some(context) = self.package.cache.context() {
            context
                .consume(Resource::Work, self.package.catalog.parts.len() as u64)
                .map_err(map_execution_error)?;
        }
        let part = self.part(partname)?;
        let effective_index = (!self.state.no_op).then_some(part.index);
        let source = part.source_xml()?;
        let revision = source.clone();
        let mut publication = revision.clone().into_publication()?;
        edit(&revision, &mut publication)?;
        publication.check_edit_base(&revision)?;
        self.state.check_readback_sources(&self.package)?;
        let result = publication.finish()?;
        if result.bytes() == source.bytes() {
            return Ok(self.clone());
        }
        let Some(index) = effective_index else {
            let mut plan = SourceTopologyPlan::new();
            plan.try_replace_source_xml_part(partname.clone(), result)?;
            return self.package.clone().prepare_topology(plan);
        };
        self.replace_effective_xml(index, result)
    }

    /// Apply a source-authorized XML target with an exact effective precondition.
    ///
    /// Both tokens must belong to the destination physical Part. The expected
    /// token must match the current effective bytes; unrelated retained graph
    /// changes survive. This low-level operation validates XML provenance, not
    /// format-specific semantic operations or protection policy.
    ///
    /// # Errors
    /// Refuses foreign/stale tokens, unproven current Parts, invalid XML,
    /// changed sources, cancellation or admission failures.
    pub fn apply_source_xml(
        &self,
        expected: &SourceXmlPart,
        target: SourceXmlPart,
    ) -> Result<Self> {
        self.check_source()?;
        let result = (|| {
            let part = self.part(expected.partname())?;
            let current = part.source_xml()?;
            let physical = self.package.part(expected.partname())?;
            let original = physical.data()?;
            expected.check_for_replacement(
                &self.package.source,
                physical.partname(),
                physical.content_type(),
                original.as_bytes(),
                self.package.limits,
            )?;
            if expected.bytes() != current.bytes() {
                return Err(overlay_unavailable(
                    "source XML precondition differs from effective Part",
                ));
            }
            target.check_for_replacement(
                &self.package.source,
                physical.partname(),
                physical.content_type(),
                original.as_bytes(),
                self.package.limits,
            )?;
            if target.bytes() == current.bytes() {
                return Ok(self.clone());
            }
            if self.state.no_op {
                let mut plan = SourceTopologyPlan::new();
                plan.try_replace_source_xml_part(expected.partname().clone(), target)?;
                self.package.clone().prepare_topology(plan)
            } else {
                self.replace_effective_xml(part.index, target)
            }
        })();
        self.check_source()?;
        result
    }

    /// Replace an unmanaged physical XML Part under an exact effective-byte precondition.
    ///
    /// Uses the ordinary full-document publication audit, including producer
    /// namespace spellings that cannot authorize positional XML splices.
    /// Format-specific semantics and source identity are the caller's responsibility.
    /// XML is audited even when the requested bytes are unchanged; callers that
    /// need an arbitrary exact-artifact no-op should retain the current candidate.
    ///
    /// # Errors
    /// Refuses managed base/donors, added Parts, non-XML Parts, mismatched bytes,
    /// invalid XML, signatures, source changes or resource limits.
    pub fn apply_xml_bytes(
        &self,
        partname: &PackURI,
        expected: &[u8],
        replacement: Arc<Vec<u8>>,
    ) -> Result<Self> {
        self.apply_xml_payload(partname, expected, replacement, None)
    }

    /// [`Self::apply_xml_bytes`] for bytes the source XML audit has already
    /// passed: `replacement` is the audit's own proof, and its allocation is
    /// published.
    ///
    /// The proof stands in for auditing the replacement again only when it
    /// covers its bytes under the default audit limits, the limits this
    /// publication audits with; a proof under any other limits is ignored and
    /// the bytes are audited in full. The physical Part's original bytes are
    /// audited as by [`Self::apply_xml_bytes`] (once per unmanaged source
    /// revision). Every refusal, and its order, is that method's.
    ///
    /// # Errors
    /// As [`Self::apply_xml_bytes`].
    pub fn apply_verified_xml_bytes(
        &self,
        partname: &PackURI,
        expected: &[u8],
        replacement: xml_minifier::audit::VerifiedSource,
    ) -> Result<Self> {
        let bytes = Arc::clone(replacement.bytes());
        self.apply_xml_payload(partname, expected, bytes, Some(replacement))
    }

    fn apply_xml_payload(
        &self,
        partname: &PackURI,
        expected: &[u8],
        replacement: Arc<Vec<u8>>,
        proof: Option<xml_minifier::audit::VerifiedSource>,
    ) -> Result<Self> {
        self.check_source()?;
        let result = (|| {
            if self.is_managed() {
                return Err(OpcError::ManagedPackageMaterialization);
            }
            let part = self.part(partname)?;
            let current = part.data()?;
            if current.as_bytes() != expected {
                return Err(overlay_unavailable(
                    "XML precondition differs from effective Part",
                ));
            }
            if !xml_minifier::audit::package::is_xml_part(partname.as_str(), part.content_type()) {
                return Err(overlay_unavailable(
                    "full XML replacement requires an XML Part",
                ));
            }
            if !self.state.no_op
                && self.state.effective_data().parts[part.index]
                    .source_index
                    .is_none()
            {
                return Err(overlay_unavailable(
                    "full XML replacement requires a physical Part",
                ));
            }
            self.package.limits.check(
                SourceLimitKind::PartBytes,
                replacement.len() as u64,
                self.package.limits.max_part_bytes(),
            )?;
            let physical = self.package.part(partname)?;
            let original = physical.data()?;
            // The effective revision's own proof, when it keeps one: the
            // replacement edits those bytes, not the physical original's
            // (record 0317).
            let predecessor = (!self.state.no_op)
                .then(|| {
                    self.state.effective_data().parts[part.index]
                        .xml_proof
                        .as_ref()
                })
                .flatten();
            let proof = self.package.audit_physical_replacement_xml(
                physical.index,
                &original,
                &replacement,
                None,
                proof.as_ref(),
                predecessor,
            )?;
            if replacement.as_slice() == current.as_bytes() {
                return Ok(self.clone());
            }
            if self.state.no_op {
                // The audit's proof travels with the plan, so publication
                // does not audit the same pair a second time.
                let mut plan = SourceTopologyPlan::new();
                match proof {
                    Some(proof) => plan.try_replace_part_verified(partname.clone(), proof)?,
                    None => plan.try_replace_part_shared(partname.clone(), replacement)?,
                }
                self.package.clone().prepare_topology(plan)
            } else {
                self.replace_effective_payload(part.index, replacement, None, proof)
            }
        })();
        self.check_source()?;
        result
    }

    fn replace_effective_xml(&self, index: usize, result: SourceXmlPart) -> Result<Self> {
        let data = self.state.effective_data();
        let old_part = &data.parts[index];
        let target = old_part
            .source_index
            .expect("edit resolved a physical Part");
        let original = self.package.read_part(target)?;
        result.check_for_replacement(
            &self.package.source,
            old_part.partname(&self.package),
            old_part.content_type(&self.package),
            original.as_bytes(),
            self.package.limits,
        )?;
        self.replace_effective_payload(index, result.payload_arc(), Some(result), None)
    }

    fn replace_effective_payload(
        &self,
        index: usize,
        payload: Arc<Vec<u8>>,
        proof: Option<SourceXmlPart>,
        xml_proof: Option<xml_minifier::audit::VerifiedSource>,
    ) -> Result<Self> {
        let data = self.state.effective_data();
        let old_part = &data.parts[index];
        let target = old_part
            .source_index
            .expect("edit resolved a physical Part");
        let original = self.package.read_part(target)?;
        let old_bytes = old_part
            .replacement
            .as_ref()
            .map_or(original.as_bytes().len(), |bytes| bytes.len()) as u64;
        let new_bytes = payload.as_slice().len() as u64;
        self.package.limits.check(
            SourceLimitKind::EntryBytes,
            new_bytes,
            self.package.limits.max_entry_bytes(),
        )?;
        let previous = self
            .state
            .totals
            .expect("changed candidate has validated totals");
        let parts_total = adjusted_overlay_total(
            previous.parts,
            old_bytes,
            new_bytes,
            SourceLimitKind::TotalPartBytes,
            self.package.limits.max_total_part_bytes(),
        )?;
        let archive_total = adjusted_overlay_total(
            previous.archive,
            old_bytes,
            new_bytes,
            SourceLimitKind::AggregateBytes,
            self.package.limits.max_aggregate_bytes(),
        )?;
        self.package.limits.check(
            SourceLimitKind::TotalPartBytes,
            parts_total,
            self.package.limits.max_total_part_bytes(),
        )?;
        self.package.limits.check(
            SourceLimitKind::AggregateBytes,
            archive_total,
            self.package.limits.max_aggregate_bytes(),
        )?;
        let old_output = self
            .state
            .changed
            .iter()
            .find_map(|overlay| {
                matches!(overlay.target, ChangedOverlayTarget::Part(part) if part == target)
                    .then_some(overlay.replacement.len() as u64)
            })
            .unwrap_or(0);
        let new_output = if payload.as_slice() == original.as_bytes() {
            0
        } else {
            new_bytes
        };
        let output_bound = old_output
            .checked_mul(2)
            .and_then(|old| previous.output_bound.checked_sub(old))
            .and_then(|n| new_output.checked_mul(2).and_then(|new| n.checked_add(new)))
            .ok_or_else(|| overlay_unavailable("revision output bound overflows"))?;
        let retained = self.state.retained();
        let removed = self
            .package
            .catalog
            .parts
            .len()
            .checked_add(retained._additions.len())
            .and_then(|n| n.checked_sub(data.parts.len()))
            .ok_or_else(|| overlay_unavailable("revision removal count overflows"))?;
        let other_replacements = self.state.changed.iter().filter(|overlay| {
            matches!(overlay.target, ChangedOverlayTarget::Part(part) if part != target)
        }).count();
        let operations = removed
            .checked_add(retained._additions.len())
            .and_then(|n| n.checked_add(other_replacements))
            .and_then(|n| n.checked_add(usize::from(new_output != 0)))
            .ok_or_else(|| overlay_unavailable("revision operation count overflows"))?;
        if operations > MAX_SOURCE_TOPOLOGY_PARTS {
            return Err(overlay_unavailable(
                "revision exceeds topology Part operation bound",
            ));
        }
        // Admit every copied slot/name before cloning. Immutable payloads,
        // relationship graphs and all unaffected Part records remain shared.
        let count = self
            .state
            .changed
            .len()
            .checked_add(1)
            .ok_or_else(|| overlay_unavailable("revision overlay count overflows"))?;
        let mut bytes = count
            .checked_mul(size_of::<ChangedOverlay>())
            .and_then(|n| {
                n.checked_add(
                    data.parts
                        .len()
                        .checked_mul(size_of::<Arc<PreparedPart>>())?,
                )
            })
            .and_then(|n| {
                n.checked_add(size_of::<PreparedTopologyState>() + size_of::<[usize; 2]>())
            })
            .ok_or_else(|| overlay_unavailable("revision metadata bound overflows"))?;
        for overlay in &self.state.changed {
            let name = match &overlay.target {
                ChangedOverlayTarget::Member(name) => name.len(),
                _ => 0,
            };
            let payload = match &overlay.replacement {
                ChangedOverlayPayload::Owned(bytes) => bytes.len(),
                _ => 0,
            };
            bytes = bytes
                .checked_add(name)
                .and_then(|n| n.checked_add(payload))
                .ok_or_else(|| overlay_unavailable("revision metadata bound overflows"))?;
        }
        let reservation = self.package.reserve_topology_memory(
            u64::try_from(bytes)
                .map_err(|_| overlay_unavailable("revision metadata exceeds u64"))?,
        )?;
        let part_reservation = self.package.reserve_topology_memory(
            (size_of::<PreparedPart>() + size_of::<[usize; 2]>()) as u64,
        )?;
        let mut changed = Vec::new();
        changed
            .try_reserve_exact(count)
            .map_err(|source| OpcError::Allocation {
                resource: "source topology revision overlays",
                source,
            })?;
        changed.extend(self.state.changed.iter().filter(|overlay| {
            !matches!(overlay.target, ChangedOverlayTarget::Part(part) if part == target)
        }).cloned());
        if payload.as_slice() != original.as_bytes() {
            changed.push(ChangedOverlay {
                target: ChangedOverlayTarget::Part(target),
                replacement: ChangedOverlayPayload::Shared(Arc::clone(&payload)),
            });
        }
        let mut parts = Vec::new();
        parts
            .try_reserve_exact(data.parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "source topology revision Part catalog",
                source,
            })?;
        parts.extend(data.parts.iter().cloned());
        let xml_proof = xml_proof.filter(|xml_proof| {
            !self.package.cache.is_managed()
                && xml_proof.covers(payload.as_slice(), xml_minifier::audit::Limits::default())
        });
        if xml_proof.is_some() {
            retain_effective_xml_proofs(&mut parts, index);
        }
        let mut part = old_part.as_ref().clone();
        part.replacement = Some(Arc::clone(&payload));
        part.source_xml = proof;
        part.xml_proof = xml_proof;
        part._metadata_reservation = part_reservation;
        parts[index] = Arc::new(part);
        let retained = self.state.retained();
        let no_op = changed.is_empty()
            && retained.omitted_members.is_empty()
            && retained.appended.is_empty();
        let effective = EffectiveTopologyData {
            package_relationships: data.package_relationships.clone(),
            parts,
            content_types: data.content_types.clone(),
            added_member_names: Arc::clone(&data.added_member_names),
        };
        self.state.check_readback_sources(&self.package)?;
        Ok(Self {
            package: self.package.clone(),
            state: Arc::new(PreparedTopologyState {
                changed,
                retained: self.state.retained.clone(),
                effective: Some(effective),
                totals: Some(PreparedTopologyTotals {
                    parts: parts_total,
                    archive: archive_total,
                    output_bound,
                }),
                _catalog_reservation: None,
                no_op,
            }),
            _state_reservation: reservation,
        })
    }
}

/// Make room for one more retained replacement proof at `index`: when
/// [`RETAINED_EFFECTIVE_XML_PROOFS`] other Parts of the revision keep one,
/// the first of them gives its proof up. Its verdict is unaffected; its next
/// edit is audited from the physical original's proof instead.
fn retain_effective_xml_proofs(parts: &mut [Arc<PreparedPart>], index: usize) {
    let kept = parts
        .iter()
        .enumerate()
        .filter(|(at, part)| *at != index && part.xml_proof.is_some())
        .count();
    if kept < RETAINED_EFFECTIVE_XML_PROOFS {
        return;
    }
    if let Some((_, part)) = parts
        .iter_mut()
        .enumerate()
        .find(|(at, part)| *at != index && part.xml_proof.is_some())
    {
        let mut stripped = part.as_ref().clone();
        stripped.xml_proof = None;
        *part = Arc::new(stripped);
    }
}
