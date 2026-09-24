//! Bounded relationship indexing for the slide-owned InkAction owner.
//!
//! The package owner needs package-wide inbound references when deciding
//! whether an action target is shared.  Building that set once per inventory
//! avoids a full package walk for every selected target and, more
//! importantly, makes case-insensitive OPC resolution a single checked
//! operation.  The index borrows the package's relationship objects.  It does
//! not clone every relationship ID, type, and target string; callers copy only
//! the bounded edges that belong to a selected snapshot.
//!
//! Index construction scans package parts and relationship records once,
//! then performs deterministic part-name and per-bucket sorting.  Target
//! resolution uses binary search over the retained part-name index.  The
//! retained metadata is `O(parts + unique internal targets + relationships)`
//! map/vector storage, with at most one borrowed outbound reference and one
//! borrowed inbound reference per internal edge; selected snapshot copies
//! remain bounded by the action target-relationship limit.  The OPC graph-node
//! ceiling applies to unique internal target keys, while source buckets are
//! bounded by the admitted part inventory.

use std::cmp::Ordering;
use std::collections::HashMap;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcError, OpcPackage, PackURI, Relationship, TargetMode};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{InboundReference, Limits};
use crate::{Error, Result};

/// The package pseudo-part used as the source of package-root relationships.
const PACKAGE_ROOT: &str = "/";

// These relationship forms are part of the ISO/IEC 29500 inventory but are
// not exposed by the older litchi-opc constant table.  Keep them local to the
// classifier rather than widening that public OPC constants API for this
// package-owner diagnostic.
const CUSTOM_XML_PROPS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const STRICT_CUSTOM_XML_PROPS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps";
const STRICT_SLIDE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const STRICT_SLIDE_LAYOUT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";
const STRICT_SLIDE_MASTER_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_HANDOUT_MASTER_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster";
const STRICT_COMMENT_AUTHORS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors";
const STRICT_PRES_PROPS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps";
const STRICT_VIEW_PROPS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps";
const STRICT_THEME_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";
const STRICT_THEME_OVERRIDE_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride";
const STRICT_TABLE_STYLES_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles";

/// Relationship fields are copied only when a selected edge is materialized.
/// Keep this owner ceiling below an unbounded `String` allocation even for a
/// package authored in memory rather than read through `litchi-opc`.
const MAX_RELATIONSHIP_FIELD_BYTES: usize = 4096;

/// A relationship record borrowed from the package inventory.
///
/// The source part name and the relationship itself remain owned by the
/// caller's [`OpcPackage`].  This small record is `Copy`, so indexing all
/// package edges does not copy relationship payload strings.  Use
/// [`RelationshipReference::to_owned`] at the package-owner boundary when an
/// immutable snapshot needs to retain an edge after the index is dropped.
#[derive(Clone, Copy, Debug)]
pub(crate) struct RelationshipReference<'package> {
    source_part: &'package str,
    relationship: &'package Relationship,
}

impl<'package> RelationshipReference<'package> {
    /// Source part URI, or `/` for the package-root relationship collection.
    #[must_use]
    pub(crate) fn source_part_name(self) -> &'package str {
        self.source_part
    }

    /// Relationship ID retained by the source package.
    #[must_use]
    pub(crate) fn relationship_id(self) -> &'package str {
        self.relationship.r_id()
    }

    /// Relationship type URI retained by the source package.
    #[must_use]
    pub(crate) fn relationship_type(self) -> &'package str {
        self.relationship.reltype()
    }

    /// Original target reference, including its lexical spelling.
    #[must_use]
    pub(crate) fn target_ref(self) -> &'package str {
        self.relationship.target_ref()
    }

    /// Relationship target mode.
    #[must_use]
    pub(crate) fn target_mode(self) -> TargetMode {
        self.relationship.target_mode()
    }

    /// Materialize the edge into the source-backed InkAction model.
    ///
    /// This is deliberately a fallible boundary: source names and relationship
    /// fields can be mutated on an in-memory package after construction, so the
    /// copy is checked again before it becomes retained snapshot state.
    pub(crate) fn to_owned(self) -> Result<InboundReference> {
        let source_part = PackURI::new(self.source_part).map_err(|error| {
            Error::Relationship(format!(
                "ink-action relationship source '{}' is invalid: {error}",
                self.source_part
            ))
        })?;
        Ok(InboundReference {
            source_part,
            relationship_id: clone_field(
                self.relationship_id(),
                "ink-action relationship metadata bytes",
            )?,
            relationship_type: clone_field(
                self.relationship_type(),
                "ink-action relationship metadata bytes",
            )?,
            target_ref: clone_field(self.target_ref(), "ink-action relationship metadata bytes")?,
            target_mode: self.target_mode(),
        })
    }
}

/// A target's borrowed relationship closure.
///
/// Both slices are deterministic: references are sorted by source part, ID,
/// type, target spelling, and target mode.  `outbound` includes unknown and
/// external relationships so a scalar profile edit can preserve their
/// diagnostic read-set.  Known ECMA outbound edges are rejected by
/// [`RelationshipIndex::closure`] for the selected target only.
#[derive(Clone, Copy, Debug)]
pub(crate) struct RelationshipClosure<'index, 'package> {
    inbound: &'index [RelationshipReference<'package>],
    outbound: &'index [RelationshipReference<'package>],
}

impl<'index, 'package> RelationshipClosure<'index, 'package> {
    /// All inbound package and part edges to the target.
    #[allow(
        dead_code,
        reason = "retained for graph diagnostics and future owner lifecycles"
    )]
    #[must_use]
    pub(crate) fn inbound(&self) -> &'index [RelationshipReference<'package>] {
        self.inbound
    }

    /// All outbound target edges, including unknown edges kept for diagnostics.
    #[allow(
        dead_code,
        reason = "retained for graph diagnostics and future owner lifecycles"
    )]
    #[must_use]
    pub(crate) fn outbound(&self) -> &'index [RelationshipReference<'package>] {
        self.outbound
    }

    /// Materialize inbound edges for an immutable owner snapshot.
    pub(crate) fn owned_inbound(&self) -> Result<Vec<InboundReference>> {
        materialize(self.inbound)
    }

    /// Materialize outbound edges for an immutable owner snapshot.
    pub(crate) fn owned_outbound(&self) -> Result<Vec<InboundReference>> {
        materialize(self.outbound)
    }
}

/// Shared package relationship index for one owner inventory.
///
/// The index borrows relationship objects and part names from `package`, so
/// its lifetime prevents a caller from mutating or replacing the package while
/// the index is in use.  The selected target's inbound and outbound closure is
/// capped by [`Limits::target_relationships`]; the index itself retains
/// borrowed references for every admitted relationship so an unrelated part
/// cannot make a selected target fail during the inventory pass.  Aggregate
/// edge and key allocations are bounded by the package's retained
/// [`litchi_opc::ReadLimits`], and all fallible collection growth uses
/// `try_reserve`.
#[derive(Debug)]
pub(crate) struct RelationshipIndex<'package> {
    package: &'package OpcPackage,
    target_relationships: usize,
    /// Physical part names sorted under OPC's ASCII-case-insensitive order.
    /// The borrowed vector replaces a linear `OpcPackage::get_part` fallback
    /// scan for every internal relationship target.
    part_names: Vec<&'package PackURI>,
    inbound: HashMap<PackURI, Vec<RelationshipReference<'package>>>,
    outbound: HashMap<PackURI, Vec<RelationshipReference<'package>>>,
}

impl<'package> RelationshipIndex<'package> {
    /// Build one bounded index over package-root and every part relationship
    /// collection.
    ///
    /// Internal target resolution is fail-closed.  An internal query,
    /// fragment, malformed relative target, or other resolver error returns a
    /// typed relationship failure instead of silently omitting the edge.
    /// Missing parts are retained under the checked target URI so a later
    /// selected-owner lookup can report the missing target with its source
    /// edge intact.
    pub(crate) fn build(package: &'package OpcPackage, limits: &Limits) -> Result<Self> {
        let read_limits = package.read_limits();
        let part_count = package.part_count();
        if part_count > read_limits.max_parts() {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::Parts,
                actual: to_u64(part_count, "ink-action package part count")?,
                maximum: to_u64(read_limits.max_parts(), "ink-action package part limit")?,
            }));
        }
        validate_package_part_bytes(package, read_limits)?;

        let mut part_names = Vec::new();
        part_names
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("ink-action canonical part-name index", source))?;
        for part in package.iter_parts() {
            part_names.push(part.partname());
        }
        part_names.sort_unstable_by(|left, right| {
            ascii_case_insensitive_cmp(left.as_str(), right.as_str())
                .then_with(|| left.as_str().cmp(right.as_str()))
        });
        for pair in part_names.windows(2) {
            if pair[0].is_equivalent_to(pair[1]) {
                if pair[0].as_str() == pair[1].as_str() {
                    return Err(Error::Opc(OpcError::DuplicatePartName(clone_field(
                        pair[1].as_str(),
                        "ink-action canonical part-name ambiguity",
                    )?)));
                }
                return Err(Error::Opc(OpcError::EquivalentPartNames {
                    existing: clone_field(
                        pair[0].as_str(),
                        "ink-action canonical part-name ambiguity",
                    )?,
                    candidate: clone_field(
                        pair[1].as_str(),
                        "ink-action canonical part-name ambiguity",
                    )?,
                }));
            }
        }

        let mut index = Self {
            package,
            target_relationships: limits.target_relationships,
            part_names,
            inbound: HashMap::new(),
            outbound: HashMap::new(),
        };
        // There is at most one outbound source bucket per part, plus the
        // package root.  Reserve both maps before the first insertion so an
        // adversarial package cannot force an infallible growth during the
        // initial inventory pass.
        let source_buckets = part_count
            .checked_add(1)
            .ok_or_else(|| limit("ink-action relationship source buckets", part_count))?;
        // Outbound buckets are keyed by relationship *sources*, so they are
        // bounded by the admitted part inventory rather than the OPC graph
        // node ceiling.  That ceiling counts unique internal target nodes;
        // applying it to sources would reject a package with many owners
        // converging on one target at the exact node budget.
        index
            .inbound
            .try_reserve(source_buckets.min(read_limits.max_relationship_graph_nodes()))
            .map_err(|source| allocation("ink-action inbound relationship index", source))?;
        index
            .outbound
            .try_reserve(source_buckets)
            .map_err(|source| allocation("ink-action outbound relationship index", source))?;

        let mut total_edges = 0usize;
        index.index_relationships(PACKAGE_ROOT, package.rels().iter(), &mut total_edges)?;
        for part in package.iter_parts() {
            index.index_relationships(
                part.partname().as_str(),
                part.rels().iter(),
                &mut total_edges,
            )?;
        }

        for references in index.inbound.values_mut() {
            sort_references(references);
        }
        for references in index.outbound.values_mut() {
            sort_references(references);
        }
        Ok(index)
    }

    /// Return the canonical physical part name for a relationship target.
    ///
    /// The bounded part-name index performs the OPC-required ASCII
    /// case-insensitive lookup.  A missing target keeps its checked URI as a
    /// diagnostic key, allowing the selected owner to produce a typed missing
    /// target error later.
    pub(crate) fn canonical_target(&self, target: &PackURI) -> PackURI {
        canonical_target_name(&self.part_names, target)
    }

    /// Borrow the package-wide inbound set for `target` after canonical OPC
    /// resolution.  This accessor never applies the selected-target outbound
    /// conformance rule; use [`Self::closure`] when reading a typed target.
    #[allow(dead_code, reason = "retained for generic graph callers")]
    #[must_use]
    pub(crate) fn inbound(&self, target: &PackURI) -> &[RelationshipReference<'package>] {
        let canonical = self.canonical_target(target);
        self.inbound
            .get(&canonical)
            .map(Vec::as_slice)
            .unwrap_or(&[])
    }

    /// Borrow a target's complete outbound diagnostic set.
    ///
    /// This low-level accessor intentionally does not reject known ECMA edges;
    /// it is useful to preserve a malformed/unsupported source through an
    /// opaque path.  Typed action publication should call [`Self::closure`].
    #[allow(dead_code, reason = "retained for generic graph callers")]
    #[must_use]
    pub(crate) fn outbound(&self, target: &PackURI) -> &[RelationshipReference<'package>] {
        let canonical = self.canonical_target(target);
        self.outbound
            .get(&canonical)
            .map(Vec::as_slice)
            .unwrap_or(&[])
    }

    /// Return the selected target's complete closure.
    ///
    /// Unknown outbound edges are retained in the returned closure.  An edge
    /// to an ECMA-376-defined part is a typed conformance refusal only here,
    /// for the selected target; unrelated package parts do not make an action
    /// inventory fail merely because they have ordinary known relationships.
    pub(crate) fn closure(&self, target: &PackURI) -> Result<RelationshipClosure<'_, 'package>> {
        let canonical = self.canonical_target(target);
        let inbound = self
            .inbound
            .get(&canonical)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        let outbound = self
            .outbound
            .get(&canonical)
            .map(Vec::as_slice)
            .unwrap_or(&[]);
        check_target_relationship_limit(
            inbound.len(),
            self.target_relationships,
            "ink-action relationship edges",
        )?;
        check_target_relationship_limit(
            outbound.len(),
            self.target_relationships,
            "ink-action relationship edges",
        )?;
        for reference in outbound {
            if self.is_known_ecma_outbound(*reference)? {
                return Err(Error::Relationship(format!(
                    "ink-action target '{}' has an outbound relationship to an ECMA-defined part",
                    canonical.as_str()
                )));
            }
        }
        Ok(RelationshipClosure { inbound, outbound })
    }

    fn is_known_ecma_outbound(&self, reference: RelationshipReference<'package>) -> Result<bool> {
        // A known relationship URI does not make an external URL an ECMA
        // package-part edge.  Keep external links in the diagnostic closure;
        // the generic owner only refuses an internal outbound edge.
        if reference.target_mode() != TargetMode::Internal {
            return Ok(false);
        }
        if is_normative_ecma_relationship_type(reference.relationship_type()) {
            return Ok(true);
        }

        // Producers may use a relationship URI that this inventory does not
        // know.  If its internal target is nevertheless an ECMA part, the
        // selected Content part still violates the generic no-outbound-edge
        // rule.  Build already validated this target, so a resolver failure
        // here is treated as an unknown diagnostic edge rather than omitted.
        let Ok(target) =
            resolve_internal_target(reference.relationship, reference.source_part_name())
        else {
            return Ok(false);
        };
        let Some(canonical) = canonical_target_ref(&self.part_names, &target) else {
            return Ok(false);
        };
        // Only an absent part reads as a non-ECMA target. A payload decode
        // refusal is propagated: reading it as `false` would skip the typed
        // conformance refusal above (ADR 0030).
        let part = match self.package.get_part(canonical) {
            Ok(part) => part,
            Err(OpcError::PartNotFound(_)) => return Ok(false),
            Err(error) => return Err(error.into()),
        };
        Ok(is_known_ecma_content_type(part.content_type()) || has_known_ecma_root(part.blob()))
    }

    fn index_relationships<I>(
        &mut self,
        source_name: &'package str,
        relationships: I,
        total_edges: &mut usize,
    ) -> Result<()>
    where
        I: Iterator<Item = &'package Relationship>,
    {
        let source = PackURI::new(source_name).map_err(|error| {
            Error::Relationship(format!(
                "ink-action relationship source '{}' is invalid: {error}",
                source_name
            ))
        })?;
        // `source_name` is borrowed from a package part or the static root;
        // this clone is only a hash-map key and does not duplicate edge
        // payload fields.
        let source_key = source.clone();

        let read_limits = self.package.read_limits();
        let relationships_per_part_limit = to_u64(
            read_limits.max_relationships_per_part(),
            "ink-action relationships-per-part limit",
        )?;
        let total_relationships_limit = to_u64(
            read_limits.max_total_relationships(),
            "ink-action aggregate relationship edge limit",
        )?;
        let mut source_edges = 0usize;
        let source_bucket_limit = self.package.part_count().checked_add(1).ok_or_else(|| {
            limit(
                "ink-action relationship source buckets",
                self.package.part_count(),
            )
        })?;

        for relationship in relationships {
            source_edges = source_edges.checked_add(1).ok_or({
                Error::Opc(OpcError::ReadLimit {
                    resource: litchi_opc::ReadResource::RelationshipsPerPart,
                    actual: u64::MAX,
                    maximum: relationships_per_part_limit,
                })
            })?;
            if source_edges > read_limits.max_relationships_per_part() {
                return Err(Error::Opc(OpcError::ReadLimit {
                    resource: litchi_opc::ReadResource::RelationshipsPerPart,
                    actual: to_u64(source_edges, "ink-action relationships-per-part count")?,
                    maximum: relationships_per_part_limit,
                }));
            }
            *total_edges = total_edges.checked_add(1).ok_or({
                Error::Opc(OpcError::ReadLimit {
                    resource: litchi_opc::ReadResource::TotalRelationships,
                    actual: u64::MAX,
                    maximum: total_relationships_limit,
                })
            })?;
            if *total_edges > read_limits.max_total_relationships() {
                return Err(Error::Opc(OpcError::ReadLimit {
                    resource: litchi_opc::ReadResource::TotalRelationships,
                    actual: to_u64(*total_edges, "ink-action aggregate relationship edge count")?,
                    maximum: total_relationships_limit,
                }));
            }
            check_relationship_fields(relationship, read_limits)?;

            let reference = RelationshipReference {
                source_part: source_name,
                relationship,
            };
            add_edge(
                &mut self.outbound,
                source_key.clone(),
                reference,
                None,
                None,
                source_bucket_limit,
                "ink-action relationship edges",
            )?;

            if relationship.target_mode() != TargetMode::Internal {
                continue;
            }
            let target = resolve_internal_target(relationship, source_name)?;
            let canonical = canonical_target_name(&self.part_names, &target);
            add_edge(
                &mut self.inbound,
                canonical,
                reference,
                None,
                Some(read_limits.max_relationship_graph_nodes()),
                source_bucket_limit,
                "ink-action relationship edges",
            )?;
        }
        Ok(())
    }
}

fn validate_package_part_bytes(
    package: &OpcPackage,
    read_limits: litchi_opc::ReadLimits,
) -> Result<()> {
    let mut total = 0u64;
    // Every part's inflated byte count is charged, so this pass decodes every
    // payload, where a decode refusal is still reported (ADR 0030).
    for part in package.try_iter_parts() {
        let part = part?;
        let actual = to_u64(part.blob().len(), "ink-action part byte count")?;
        if actual > read_limits.max_part_bytes() {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::PartBytes,
                actual,
                maximum: read_limits.max_part_bytes(),
            }));
        }
        total = total.checked_add(actual).ok_or_else(|| {
            Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::TotalPartBytes,
                actual: u64::MAX,
                maximum: read_limits.max_total_part_bytes(),
            })
        })?;
        if total > read_limits.max_total_part_bytes() {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::TotalPartBytes,
                actual: total,
                maximum: read_limits.max_total_part_bytes(),
            }));
        }
    }
    Ok(())
}

fn materialize<'package>(
    references: &[RelationshipReference<'package>],
) -> Result<Vec<InboundReference>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(references.len())
        .map_err(|source| allocation("ink-action relationship edges", source))?;
    for reference in references {
        owned.push((*reference).to_owned()?);
    }
    Ok(owned)
}

fn clone_field(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    owned.push_str(value);
    Ok(owned)
}

fn add_edge<'package>(
    index: &mut HashMap<PackURI, Vec<RelationshipReference<'package>>>,
    key: PackURI,
    reference: RelationshipReference<'package>,
    maximum_edges: Option<usize>,
    maximum_keys: Option<usize>,
    maximum_source_buckets: usize,
    resource: &'static str,
) -> Result<()> {
    if let Some(maximum_keys) = maximum_keys {
        if !index.contains_key(&key) && index.len() >= maximum_keys {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::RelationshipGraphNodes,
                actual: to_u64(
                    index.len().saturating_add(1),
                    "ink-action relationship graph node count",
                )?,
                maximum: to_u64(maximum_keys, "ink-action relationship graph node limit")?,
            }));
        }
    } else if !index.contains_key(&key) && index.len() >= maximum_source_buckets {
        // Every outbound source is one package part or the package root.  The
        // inventory count checked before indexing bounds this map separately
        // from the OPC graph-node policy, which counts unique internal targets.
        return Err(Error::Opc(OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::Parts,
            actual: to_u64(
                index.len().saturating_add(1),
                "ink-action relationship source bucket count",
            )?,
            maximum: to_u64(
                maximum_source_buckets,
                "ink-action relationship source bucket limit",
            )?,
        }));
    }
    index
        .try_reserve(1)
        .map_err(|source| allocation("ink-action relationship index", source))?;
    let edges = index.entry(key).or_default();
    if let Some(maximum_edges) = maximum_edges {
        if edges.len() >= maximum_edges {
            return Err(limit(resource, maximum_edges));
        }
    }
    edges
        .try_reserve(1)
        .map_err(|source| allocation(resource, source))?;
    edges.push(reference);
    Ok(())
}

fn check_target_relationship_limit(
    actual: usize,
    maximum: usize,
    resource: &'static str,
) -> Result<()> {
    if actual > maximum {
        return Err(limit(resource, maximum));
    }
    Ok(())
}

fn resolve_internal_target(relationship: &Relationship, source_name: &str) -> Result<PackURI> {
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        return Err(Error::Relationship(format!(
            "ink-action relationship '{}' from '{}' has an internal query or fragment",
            relationship.r_id(),
            source_name
        )));
    }
    relationship.target_partname().map_err(|error| {
        Error::Relationship(format!(
            "ink-action relationship '{}' from '{}' has an invalid internal target: {error}",
            relationship.r_id(),
            source_name
        ))
    })
}

fn canonical_target_name(part_names: &[&PackURI], target: &PackURI) -> PackURI {
    canonical_target_ref(part_names, target).map_or_else(|| target.clone(), |name| name.clone())
}

fn canonical_target_ref<'package>(
    part_names: &[&'package PackURI],
    target: &PackURI,
) -> Option<&'package PackURI> {
    part_names
        .binary_search_by(|candidate| {
            ascii_case_insensitive_cmp(candidate.as_str(), target.as_str())
        })
        .ok()
        .map(|index| part_names[index])
}

fn ascii_case_insensitive_cmp(left: &str, right: &str) -> Ordering {
    left.as_bytes()
        .iter()
        .map(u8::to_ascii_lowercase)
        .cmp(right.as_bytes().iter().map(u8::to_ascii_lowercase))
}

fn check_relationship_fields(
    relationship: &Relationship,
    read_limits: litchi_opc::ReadLimits,
) -> Result<()> {
    let field_limit = MAX_RELATIONSHIP_FIELD_BYTES.min(read_limits.max_xml_attribute_bytes());
    for (name, value) in [
        ("relationship ID", relationship.r_id()),
        ("relationship type", relationship.reltype()),
        ("relationship target", relationship.target_ref()),
    ] {
        if value.len() > field_limit {
            return Err(limit("ink-action relationship metadata bytes", field_limit));
        }
        if name == "relationship target"
            && value.len() > read_limits.max_relationship_target_bytes()
        {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: litchi_opc::ReadResource::RelationshipTargetBytes,
                actual: to_u64(value.len(), "ink-action relationship target bytes")?,
                maximum: to_u64(
                    read_limits.max_relationship_target_bytes(),
                    "ink-action relationship target limit",
                )?,
            }));
        }
    }
    Ok(())
}

fn sort_references(references: &mut [RelationshipReference<'_>]) {
    references.sort_by(|left, right| {
        left.source_part_name()
            .cmp(right.source_part_name())
            .then_with(|| left.relationship_id().cmp(right.relationship_id()))
            .then_with(|| left.relationship_type().cmp(right.relationship_type()))
            .then_with(|| left.target_ref().cmp(right.target_ref()))
            .then_with(|| {
                target_mode_rank(left.target_mode()).cmp(&target_mode_rank(right.target_mode()))
            })
    });
}

const fn target_mode_rank(mode: TargetMode) -> u8 {
    match mode {
        TargetMode::Internal => 0,
        TargetMode::External => 1,
    }
}

const fn limit(resource: &'static str, maximum: usize) -> Error {
    Error::Limit {
        resource,
        limit: maximum,
    }
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

fn to_u64(value: usize, resource: &'static str) -> Result<u64> {
    u64::try_from(value).map_err(|_| Error::Invalid(resource.to_owned()))
}

/// Relationship types whose target part is enumerated by OPC or ECMA-376.
///
/// This is intentionally an explicit normative inventory. In particular, the
/// OPC constants module also contains Microsoft extension relationships
/// (threaded comments, VBA, and media) that do not establish an ECMA-defined
/// target for this conformance check.
fn is_normative_ecma_relationship_type(reltype: &str) -> bool {
    matches!(
        reltype,
        rt::CORE_PROPERTIES
            | rt::EXTENDED_PROPERTIES
            | rt::CUSTOM_PROPERTIES
            | rt::THUMBNAIL
            | rt::DIGITAL_SIGNATURE_ORIGIN
            | rt::OFFICE_DOCUMENT
            | rt::STRICT_OFFICE_DOCUMENT
            | rt::COMMENTS
            | rt::STRICT_COMMENTS
            | rt::ENDNOTES
            | rt::FONT
            | rt::FONT_TABLE
            | rt::FOOTER
            | rt::FOOTNOTES
            | rt::HEADER
            | rt::NUMBERING
            | rt::SETTINGS
            | rt::STYLES
            | rt::STRICT_STYLES
            | rt::SHARED_STRINGS
            | rt::STRICT_SHARED_STRINGS
            | rt::CALC_CHAIN
            | rt::STRICT_CALC_CHAIN
            | rt::EXTERNAL_LINK
            | rt::STRICT_EXTERNAL_LINK
            | rt::EXTERNAL_LINK_PATH
            | rt::STRICT_EXTERNAL_LINK_PATH
            | rt::WEB_SETTINGS
            | rt::SHEET_METADATA
            | rt::WORKSHEET
            | rt::STRICT_WORKSHEET
            | rt::ALTERNATIVE_FORMAT_IMPORT
            | rt::PIVOT_CACHE_DEFINITION
            | rt::STRICT_PIVOT_CACHE_DEFINITION
            | rt::PIVOT_CACHE_RECORDS
            | rt::STRICT_PIVOT_CACHE_RECORDS
            | rt::PIVOT_TABLE
            | rt::STRICT_PIVOT_TABLE
            | rt::TABLE
            | rt::STRICT_TABLE
            | rt::IMAGE
            | rt::STRICT_IMAGE
            | rt::AUDIO
            | rt::VIDEO
            | rt::CHART
            | rt::CHART_USER_SHAPES
            | rt::STRICT_CHART
            | rt::DRAWING
            | rt::STRICT_DRAWING
            | rt::VML_DRAWING
            | rt::THEME
            | rt::THEME_OVERRIDE
            | rt::CUSTOM_XML
            | "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml"
            | CUSTOM_XML_PROPS_RELATIONSHIP
            | STRICT_CUSTOM_XML_PROPS_RELATIONSHIP
            | rt::HYPERLINK
            | rt::STRICT_HYPERLINK
            | rt::OLE_OBJECT
            | rt::STRICT_OLE_OBJECT
            | rt::PACKAGE
            | rt::STRICT_PACKAGE
            | rt::SLIDE
            | rt::SLIDE_LAYOUT
            | rt::SLIDE_MASTER
            | rt::NOTES_SLIDE
            | rt::NOTES_MASTER
            | rt::STRICT_NOTES_SLIDE
            | rt::STRICT_NOTES_MASTER
            | rt::HANDOUT_MASTER
            | rt::COMMENT_AUTHORS
            | rt::PRES_PROPS
            | rt::VIEW_PROPS
            | rt::TABLE_STYLES
            | STRICT_SLIDE_RELATIONSHIP
            | STRICT_SLIDE_LAYOUT_RELATIONSHIP
            | STRICT_SLIDE_MASTER_RELATIONSHIP
            | STRICT_HANDOUT_MASTER_RELATIONSHIP
            | STRICT_COMMENT_AUTHORS_RELATIONSHIP
            | STRICT_PRES_PROPS_RELATIONSHIP
            | STRICT_VIEW_PROPS_RELATIONSHIP
            | STRICT_THEME_RELATIONSHIP
            | STRICT_THEME_OVERRIDE_RELATIONSHIP
            | STRICT_TABLE_STYLES_RELATIONSHIP
    )
}

/// Content types assigned by ECMA-376 or OPC to package parts.
///
/// This fallback catches an otherwise unknown relationship URI whose target is
/// visibly a standard ECMA part. Microsoft extension content types are kept
/// out of this list so an unrelated extension remains an unknown diagnostic
/// edge unless its relationship URI is itself normative.
fn is_known_ecma_content_type(content_type: &str) -> bool {
    matches!(
        content_type,
        ct::BMP
            | ct::GIF
            | ct::JPEG
            | ct::PNG
            | ct::TIFF
            | ct::X_EMF
            | ct::X_WMF
            | ct::DML_CHART
            | ct::DML_CHARTSHAPES
            | ct::DML_DIAGRAM_COLORS
            | ct::DML_DIAGRAM_DATA
            | ct::DML_DIAGRAM_LAYOUT
            | ct::DML_DIAGRAM_STYLE
            | ct::OFC_CUSTOM_PROPERTIES
            | ct::OFC_CUSTOM_XML_PROPERTIES
            | ct::OFC_DRAWING
            | ct::OFC_EXTENDED_PROPERTIES
            | ct::OFC_OLE_OBJECT
            | ct::OFC_PACKAGE
            | ct::OFC_THEME
            | ct::OFC_THEME_OVERRIDE
            | ct::OFC_VML_DRAWING
            | ct::OPC_CORE_PROPERTIES
            | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
            | ct::OPC_DIGITAL_SIGNATURE_ORIGIN
            | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
            | ct::OPC_RELATIONSHIPS
            | ct::WML_COMMENTS
            | ct::WML_DOCUMENT
            | ct::WML_DOCUMENT_GLOSSARY
            | ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_ENDNOTES
            | ct::WML_FONT_TABLE
            | ct::WML_FOOTER
            | ct::WML_FOOTNOTES
            | ct::WML_HEADER
            | ct::WML_NUMBERING
            | ct::WML_SETTINGS
            | ct::WML_STYLES
            | ct::WML_WEB_SETTINGS
            | ct::SML_SHEET
            | ct::SML_SHEET_MAIN
            | ct::SML_TEMPLATE_MAIN
            | ct::SML_WORKSHEET
            | ct::SML_STYLES
            | ct::SML_SHARED_STRINGS
            | ct::SML_EXTERNAL_LINK
            | ct::SML_PIVOT_CACHE_DEFINITION
            | ct::SML_PIVOT_CACHE_RECORDS
            | ct::SML_PIVOT_TABLE
            | ct::SML_TABLE
            | ct::SML_COMMENTS
            | ct::SML_SHEET_METADATA
            | ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_SLIDE
            | ct::PML_SLIDE_LAYOUT
            | ct::PML_SLIDE_MASTER
            | ct::PML_TABLE_STYLES
            | ct::PML_VIEW_PROPS
            | ct::PML_PRES_PROPS
            | ct::PML_NOTES_SLIDE
            | ct::PML_NOTES_MASTER
            | ct::PML_HANDOUT_MASTER
            | ct::PML_COMMENTS
            | ct::PML_COMMENT_AUTHORS
            | ct::AUDIO_MPEG
            | ct::AUDIO_WAV
            | ct::AUDIO_M4A
            | ct::VIDEO_MP4
            | ct::VIDEO_AVI
            | ct::VIDEO_MOV
    )
}

// Namespace families in which the first root element identifies an
// ECMA-defined XML part even when its content type was declared generically.
const ECMA_WML_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const ECMA_WML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const ECMA_SML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const ECMA_SML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/spreadsheetml/main";
const ECMA_PML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const ECMA_PML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const ECMA_DML_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const ECMA_DML_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const ECMA_CORE_PROPERTIES_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const ECMA_CORE_PROPERTIES_STRICT: &[u8] =
    b"http://purl.oclc.org/ooxml/package/metadata/core-properties";
const ECMA_EXTENDED_PROPERTIES_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
const ECMA_EXTENDED_PROPERTIES_STRICT: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/extended-properties";
const ECMA_CUSTOM_PROPERTIES_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
const ECMA_CUSTOM_PROPERTIES_STRICT: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/custom-properties";
const ECMA_CUSTOM_XML_PROPERTIES_TRANSITIONAL: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/customXml";
const ECMA_CUSTOM_XML_PROPERTIES_STRICT: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/customXml";

fn has_known_ecma_root(xml: &[u8]) -> bool {
    let mut reader = NsReader::from_reader(xml);
    loop {
        let Ok((namespace, event)) = reader.read_resolved_event() else {
            return false;
        };
        match event {
            Event::Start(element) | Event::Empty(element) => {
                return is_known_ecma_root_namespace(&namespace, element.name().as_ref());
            },
            Event::Eof => return false,
            _ => {},
        }
    }
}

fn is_known_ecma_root_namespace(namespace: &ResolveResult<'_>, _local_name: &[u8]) -> bool {
    let ResolveResult::Bound(Namespace(uri)) = namespace else {
        return false;
    };
    [
        ECMA_WML_TRANSITIONAL,
        ECMA_WML_STRICT,
        ECMA_SML_TRANSITIONAL,
        ECMA_SML_STRICT,
        ECMA_PML_TRANSITIONAL,
        ECMA_PML_STRICT,
        ECMA_DML_TRANSITIONAL,
        ECMA_DML_STRICT,
        ECMA_CORE_PROPERTIES_TRANSITIONAL,
        ECMA_CORE_PROPERTIES_STRICT,
        ECMA_EXTENDED_PROPERTIES_TRANSITIONAL,
        ECMA_EXTENDED_PROPERTIES_STRICT,
        ECMA_CUSTOM_PROPERTIES_TRANSITIONAL,
        ECMA_CUSTOM_PROPERTIES_STRICT,
        ECMA_CUSTOM_XML_PROPERTIES_TRANSITIONAL,
        ECMA_CUSTOM_XML_PROPERTIES_STRICT,
    ]
    .contains(uri)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, Part};

    fn package_with_shared_case_target() -> (OpcPackage, PackURI) {
        let target_name = PackURI::new("/ppt/custom/Action.xml").unwrap();
        let first_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let second_name = PackURI::new("/ppt/slides/slide2.xml").unwrap();
        let mut first = BlobPart::new(first_name, ct::PML_SLIDE.to_owned(), Vec::new());
        let mut second = BlobPart::new(second_name, ct::PML_SLIDE.to_owned(), Vec::new());
        first.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "../custom/action.xml".to_owned(),
            "rId2".to_owned(),
            false,
        );
        second.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "../custom/ACTION.xml".to_owned(),
            "rId1".to_owned(),
            false,
        );
        let target = BlobPart::new(target_name.clone(), "text/xml".to_owned(), Vec::new());
        let mut package = OpcPackage::new();
        package.add_part(Box::new(first));
        package.add_part(Box::new(second));
        package.add_part(Box::new(target));
        (package, target_name)
    }

    #[test]
    fn indexes_root_and_case_insensitive_part_edges_once() {
        let (mut package, target_name) = package_with_shared_case_target();
        package.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "ppt/custom/action.xml".to_owned(),
            "rIdRoot".to_owned(),
            false,
        );
        let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
        let inbound = index.inbound(&PackURI::new("/PPT/CUSTOM/aCtIoN.xml").unwrap());
        assert_eq!(inbound.len(), 3);
        assert_eq!(inbound[0].source_part_name(), "/");
        assert_eq!(inbound[1].source_part_name(), "/ppt/slides/slide1.xml");
        assert_eq!(inbound[2].source_part_name(), "/ppt/slides/slide2.xml");
        assert_eq!(index.canonical_target(&target_name), target_name);
        assert_eq!(index.outbound(&target_name).len(), 0);
    }

    #[test]
    fn rejects_case_equivalent_physical_part_names_before_indexing() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/custom/Action.xml").unwrap(),
            "text/xml".to_owned(),
            Vec::new(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/PPT/CUSTOM/action.xml").unwrap(),
            "text/xml".to_owned(),
            Vec::new(),
        )));
        let error = RelationshipIndex::build(&package, &Limits::default()).unwrap_err();
        assert!(matches!(
            error,
            Error::Opc(OpcError::EquivalentPartNames { .. })
        ));
    }

    #[test]
    fn malformed_internal_query_is_not_omitted() {
        let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut slide = BlobPart::new(name, ct::PML_SLIDE.to_owned(), Vec::new());
        slide.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "../custom/action.xml?bad".to_owned(),
            "rIdBad".to_owned(),
            false,
        );
        let mut package = OpcPackage::new();
        package.add_part(Box::new(slide));
        let error = RelationshipIndex::build(&package, &Limits::default()).unwrap_err();
        assert!(matches!(error, Error::Relationship(_)));
    }

    #[test]
    fn known_ecma_edge_refuses_only_selected_target() {
        let (mut package, target_name) = package_with_shared_case_target();
        let image_name = PackURI::new("/ppt/media/image.png").unwrap();
        package.add_part(Box::new(BlobPart::new(
            image_name,
            "image/png".to_owned(),
            Vec::new(),
        )));
        package
            .get_part_mut(&target_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.to_owned(),
                "../media/image.png".to_owned(),
                "rIdImage".to_owned(),
                false,
            );
        let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
        let error = index
            .closure(&PackURI::new("/ppt/custom/action.xml").unwrap())
            .unwrap_err();
        assert!(matches!(error, Error::Relationship(_)));
        // The edge remains available for source-preserving diagnostics.
        assert_eq!(index.outbound(&target_name).len(), 1);
    }

    #[test]
    fn known_external_relationship_is_kept_as_diagnostic() {
        let (mut package, target_name) = package_with_shared_case_target();
        package
            .get_part_mut(&target_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.to_owned(),
                "https://example.invalid/image.png".to_owned(),
                "rIdExternalImage".to_owned(),
                true,
            );
        let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
        let closure = index.closure(&target_name).unwrap();
        assert_eq!(closure.outbound().len(), 1);
        assert_eq!(closure.outbound()[0].target_mode(), TargetMode::External);
    }

    #[test]
    fn strict_ecma_relationship_inventory_refuses_selected_target() {
        let strict_relationships = [
            CUSTOM_XML_PROPS_RELATIONSHIP,
            STRICT_CUSTOM_XML_PROPS_RELATIONSHIP,
            STRICT_SLIDE_RELATIONSHIP,
            STRICT_SLIDE_LAYOUT_RELATIONSHIP,
            STRICT_SLIDE_MASTER_RELATIONSHIP,
            STRICT_HANDOUT_MASTER_RELATIONSHIP,
            STRICT_COMMENT_AUTHORS_RELATIONSHIP,
            STRICT_PRES_PROPS_RELATIONSHIP,
            STRICT_VIEW_PROPS_RELATIONSHIP,
            STRICT_THEME_RELATIONSHIP,
            STRICT_THEME_OVERRIDE_RELATIONSHIP,
            STRICT_TABLE_STYLES_RELATIONSHIP,
        ];
        for (ordinal, relationship_type) in strict_relationships.into_iter().enumerate() {
            let (mut package, target_name) = package_with_shared_case_target();
            package
                .get_part_mut(&target_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    relationship_type.to_owned(),
                    "Action.xml".to_owned(),
                    format!("rIdStrict{ordinal}"),
                    false,
                );
            let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
            assert!(matches!(
                index.closure(&target_name),
                Err(Error::Relationship(message)) if message.contains("ECMA-defined")
            ));
        }
    }

    #[test]
    fn unknown_internal_ecma_content_type_or_root_is_refused() {
        for (content_type, blob) in [
            (ct::PML_SLIDE, Vec::new()),
            (
                "application/octet-stream",
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
                    .to_vec(),
            ),
        ] {
            let (mut package, target_name) = package_with_shared_case_target();
            let ecma_name = PackURI::new("/ppt/known/ecma.xml").unwrap();
            package.add_part(Box::new(BlobPart::new(
                ecma_name,
                content_type.to_owned(),
                blob,
            )));
            package
                .get_part_mut(&target_name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:example:unknown-ecma-target".to_owned(),
                    "../known/ecma.xml".to_owned(),
                    "rIdUnknownEcma".to_owned(),
                    false,
                );
            let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
            assert!(matches!(
                index.closure(&target_name),
                Err(Error::Relationship(message)) if message.contains("ECMA-defined")
            ));
        }
    }

    #[test]
    fn unknown_internal_microsoft_extension_remains_diagnostic() {
        let (mut package, target_name) = package_with_shared_case_target();
        let extension_name = PackURI::new("/ppt/known/vba.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            extension_name,
            ct::OFC_VBA_PROJECT.to_owned(),
            Vec::new(),
        )));
        package
            .get_part_mut(&target_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:example:unknown-extension".to_owned(),
                "../known/vba.bin".to_owned(),
                "rIdUnknownExtension".to_owned(),
                false,
            );
        let index = RelationshipIndex::build(&package, &Limits::default()).unwrap();
        assert!(index.closure(&target_name).is_ok());
        assert_eq!(index.outbound(&target_name).len(), 1);
    }

    #[test]
    fn per_target_cap_is_checked_when_that_target_is_selected() {
        let (mut package, shared_name) = package_with_shared_case_target();
        let small_name = PackURI::new("/ppt/custom/one.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            small_name.clone(),
            "text/xml".to_owned(),
            Vec::new(),
        )));
        package.rels_mut().add_relationship(
            rt::CUSTOM_XML.to_owned(),
            "ppt/custom/one.xml".to_owned(),
            "rIdOne".to_owned(),
            false,
        );
        let limits = Limits {
            target_relationships: 1,
            ..Limits::default()
        };
        // The shared target has three inbound references, but building the
        // package-wide borrowed index retains them so an unrelated target can
        // still be queried independently.
        let index = RelationshipIndex::build(&package, &limits).unwrap();
        assert!(index.closure(&small_name).is_ok());
        let shared_result = index.closure(&shared_name);
        assert!(matches!(
            shared_result,
            Err(Error::Limit {
                resource: "ink-action relationship edges",
                limit: 1
            })
        ));
    }
}
