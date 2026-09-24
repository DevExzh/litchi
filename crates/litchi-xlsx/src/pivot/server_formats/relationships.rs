//! Bounded, operation-local relationship capture for the PivotTable owner.
//!
//! `OpcPackage::source_relationships` is intentionally a compatibility
//! convenience and uses the default read profile.  A source-bound XLSX
//! operation must forward the package's retained [`ReadLimits`] instead.  It
//! also needs package-wide counters: a separate source capture for each
//! workbook, table, cache, and connections part would otherwise reset the
//! relationship XML, event, and edge budgets on every call.
//!
//! [`RelationshipIndex`] performs that work once.  It retains one exact
//! [`OwnedRelationships`] token for the package root and for each owner that
//! has a relationship member; an absent owner is represented by `None` rather
//! than a newly materialized canonical XML token.  It indexes every internal
//! edge by its canonical target part and keeps the edge payload borrowed from
//! the immutable package.  `SourcePart` can clone the selected token (the XML
//! is shared) and materialize one selected incoming closure.  Unknown
//! relationship types, external edges, and unusual lexical target spellings
//! remain in the exact source token; no relationship is normalized by this
//! helper.

use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use litchi_opc::{
    OpcError, OpcPackage, OwnedRelationships, PackURI, ReadLimits, RelationshipSourcePlan,
};
#[cfg(test)]
use quick_xml::events::Event;
#[cfg(test)]
use quick_xml::reader::NsReader;

use super::RelationshipState;
use crate::error::{Error, Result, invalid};

const PACKAGE_ROOT: &str = "/";

/// A relationship record borrowed from the package graph.
///
/// Keeping the edge borrowed avoids a second package-wide copy of IDs, type
/// URIs, and target references.  The owner converts only the selected bucket
/// to its source-backed `RelationshipState` values.
#[derive(Clone, Copy, Debug)]
pub(super) struct RelationshipReference<'package> {
    source: &'package str,
    relationship: &'package litchi_opc::Relationship,
}

impl<'package> RelationshipReference<'package> {
    /// Source part name, or `/` for package-root relationships.
    #[must_use]
    pub(super) fn source(&self) -> &'package str {
        self.source
    }

    /// Borrow the complete relationship record, including unknown fields
    /// represented by the OPC model.
    #[must_use]
    pub(super) fn relationship(&self) -> &'package litchi_opc::Relationship {
        self.relationship
    }
}

/// One bounded operation-local package relationship inventory.
///
/// The lifetime ties borrowed edges and part names to the immutable
/// `OpcPackage` used for the operation.  The package may therefore be changed
/// only after this index is dropped; in-memory relationship mutations are
/// admitted by constructing a fresh index after the mutation.
#[derive(Debug)]
pub(super) struct RelationshipIndex<'package> {
    limits: ReadLimits,
    part_names: Vec<&'package PackURI>,
    part_indices: HashMap<String, usize>,
    root_source: OwnedRelationships,
    part_sources: Vec<Option<OwnedRelationships>>,
    incoming: Vec<Vec<RelationshipReference<'package>>>,
}

impl<'package> RelationshipIndex<'package> {
    /// Capture the package root and every part relationship collection once.
    ///
    /// The operation first obtains output-free source plans for every current
    /// relationship collection, admits all retained-member, XML, event, edge,
    /// and graph-node budgets, and only then materializes any source token.
    /// Internal target nodes are counted after ASCII-case-insensitive OPC
    /// canonicalization, including targets of relationships changed in
    /// memory after package open. The plans are descriptive admission data;
    /// each token is captured afresh after aggregate admission.
    pub(super) fn build(package: &'package OpcPackage) -> Result<Self> {
        let limits = package.read_limits();
        let part_count = package.part_count();
        check_limit(
            litchi_opc::ReadResource::Parts,
            part_count,
            limits.max_parts(),
        )?;

        let mut part_names = Vec::new();
        part_names
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("PivotTable relationship part index", source))?;
        for part in package.iter_parts() {
            check_part_name_limit(part.partname(), limits)?;
            part_names.push(part.partname());
        }
        part_names.sort_unstable_by(|left, right| {
            ascii_case_insensitive_cmp(left.as_str(), right.as_str())
                .then_with(|| left.as_str().cmp(right.as_str()))
        });

        let mut part_indices = HashMap::new();
        part_indices
            .try_reserve(part_count)
            .map_err(|source| allocation("PivotTable canonical relationship part index", source))?;
        for (index, name) in part_names.iter().enumerate() {
            let key = canonical_part_key(name, limits)?;
            if let Some(previous) = part_indices.insert(key, index) {
                let existing = part_names[previous].as_str();
                if existing == name.as_str() {
                    return Err(Error::Package(OpcError::DuplicatePartName(copy_bounded(
                        existing,
                        "PivotTable duplicate part name",
                    )?)));
                }
                return Err(Error::Package(OpcError::EquivalentPartNames {
                    existing: copy_bounded(existing, "PivotTable equivalent part name")?,
                    candidate: copy_bounded(name.as_str(), "PivotTable equivalent part name")?,
                }));
            }
        }

        let mut incoming = Vec::new();
        incoming
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("PivotTable incoming relationship buckets", source))?;
        for _ in 0..part_count {
            incoming.push(Vec::new());
        }

        let mut part_sources = Vec::new();
        part_sources
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("PivotTable relationship source tokens", source))?;

        let mut budget = RelationshipBudget::default();
        let mut graph_seen = Vec::new();
        graph_seen
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("PivotTable relationship graph nodes", source))?;
        graph_seen.extend(std::iter::repeat_n(false, part_count));
        let mut part_member_present = Vec::new();
        part_member_present
            .try_reserve_exact(part_count)
            .map_err(|source| allocation("PivotTable relationship member presence", source))?;
        // The root relationship collection has no source Part and is always
        // represented by `/` in the borrowed edge inventory.
        let root = PackURI::new(PACKAGE_ROOT).map_err(OpcError::InvalidPackUri)?;
        let root_plan = package.plan_source_relationships_with_limits(&root, limits)?;
        budget.admit_plan(root_plan, limits)?;
        index_relationships(
            package.rels().iter(),
            PACKAGE_ROOT,
            &part_indices,
            &mut incoming,
            &mut graph_seen,
            &mut budget,
            limits,
        )?;

        for name in &part_names {
            let part = package.get_part(name)?;
            let plan = package.plan_source_relationships_with_limits(name, limits)?;
            part_member_present.push(plan.member_present());
            budget.admit_plan(plan, limits)?;
            index_relationships(
                part.rels().iter(),
                name.as_str(),
                &part_indices,
                &mut incoming,
                &mut graph_seen,
                &mut budget,
                limits,
            )?;
        }

        // No source XML allocation or parser graph is retained until all
        // aggregate admissions above have succeeded. Capture each owner
        // afresh so a descriptive plan can never authorize stale material.
        let root_source = package.source_relationships_with_limits(&root, limits)?;
        for (name, member_present) in part_names.iter().zip(part_member_present) {
            let source = if member_present {
                Some(package.source_relationships_with_limits(name, limits)?)
            } else {
                None
            };
            // `part_sources` and `part_names` have the same canonical order.
            // Check the source owner so an accidental future reordering
            // cannot produce a stale token.
            if let Some(source) = source.as_ref()
                && source.owner().as_str() != name.as_str()
            {
                return Err(invalid("PivotTable relationship source owner mismatch"));
            }
            part_sources.push(source);
        }

        for bucket in &mut incoming {
            bucket.sort_unstable_by(compare_references);
        }
        Ok(Self {
            limits,
            part_names,
            part_indices,
            root_source,
            part_sources,
            incoming,
        })
    }

    /// Return the exact package-root relationship token.
    #[must_use]
    pub(super) fn root_source(&self) -> &OwnedRelationships {
        &self.root_source
    }

    /// Capture the exact relationship token for `owner`, retaining its
    /// original XML bytes and member-presence flag.  An owner without a
    /// relationship member returns `None`, preserving the absence proof
    /// without allocating canonical XML. A clone shares the XML allocation
    /// and does not copy relationship XML.
    pub(super) fn source(&self, owner: &PackURI) -> Result<Option<OwnedRelationships>> {
        if owner.as_str() == PACKAGE_ROOT {
            return Ok(Some(self.root_source.clone()));
        }
        let index = self.part_index(owner)?;
        self.part_sources
            .get(index)
            .cloned()
            .ok_or_else(|| invalid("PivotTable relationship source index is incomplete"))
    }

    /// Return a borrowed exact source token for an owner known to be in the
    /// package.  This accessor is useful for readset checks that do not need
    /// another cheap `OwnedRelationships` handle; absent relationship members
    /// return `None`.
    pub(super) fn source_ref(&self, owner: &PackURI) -> Result<Option<&OwnedRelationships>> {
        if owner.as_str() == PACKAGE_ROOT {
            return Ok(Some(&self.root_source));
        }
        let index = self.part_index(owner)?;
        self.part_sources
            .get(index)
            .map(Option::as_ref)
            .ok_or_else(|| invalid("PivotTable relationship source index is incomplete"))
    }

    /// Return all indexed incoming edges for an existing canonical target.
    /// Unknown relationship types are deliberately retained.  External edges
    /// have no package target and remain preserved by their source token.
    pub(super) fn incoming(&self, target: &PackURI) -> Result<&[RelationshipReference<'package>]> {
        let Some(index) = self.part_index_opt(target)? else {
            return Ok(&[]);
        };
        Ok(self.incoming.get(index).map(Vec::as_slice).unwrap_or(&[]))
    }

    /// Materialize one selected incoming closure into the owner's existing
    /// `RelationshipState` representation.  Every field copy is admitted
    /// through fallible reservation and the vector is bounded by the retained
    /// aggregate relationship limit.
    pub(super) fn capture_incoming(&self, target: &PackURI) -> Result<Vec<RelationshipState>> {
        let references = self.incoming(target)?;
        if references.len() > self.limits.max_total_relationships() {
            return Err(read_limit(
                litchi_opc::ReadResource::TotalRelationships,
                references.len(),
                self.limits.max_total_relationships(),
            ));
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(references.len())
            .map_err(|source| allocation("PivotTable relationship closure", source))?;
        for reference in references {
            let relationship = reference.relationship();
            values.push(RelationshipState {
                source: copy_bounded(reference.source(), "PivotTable relationship source")?,
                id: copy_bounded(relationship.r_id(), "PivotTable relationship ID")?,
                reltype: copy_bounded(relationship.reltype(), "PivotTable relationship type")?,
                target: copy_bounded(relationship.target_ref(), "PivotTable relationship target")?,
                external: relationship.is_external(),
            });
        }
        values.sort_unstable();
        Ok(values)
    }

    /// Resolve an internal target with OPC's ASCII-case-insensitive canonical
    /// lookup.  Missing targets return `None`; their canonical graph nodes
    /// were still charged during construction.
    pub(super) fn canonical_part(&self, target: &PackURI) -> Result<Option<&PackURI>> {
        Ok(self
            .part_index_opt(target)?
            .and_then(|index| self.part_names.get(index).copied()))
    }

    fn part_index(&self, owner: &PackURI) -> Result<usize> {
        let Some(index) = self.part_index_opt(owner)? else {
            return Err(Error::Package(OpcError::PartNotFound(copy_bounded(
                owner.as_str(),
                "PivotTable relationship owner",
            )?)));
        };
        Ok(index)
    }

    fn part_index_opt(&self, owner: &PackURI) -> Result<Option<usize>> {
        let key = canonical_part_key(owner, self.limits)?;
        Ok(self.part_indices.get(&key).copied())
    }
}

#[derive(Default)]
struct RelationshipBudget {
    relationship_parts: usize,
    xml_bytes: usize,
    xml_events: usize,
    relationships: usize,
    graph_nodes: usize,
    missing_nodes: HashSet<String>,
}

impl RelationshipBudget {
    fn admit_plan(&mut self, plan: RelationshipSourcePlan, limits: ReadLimits) -> Result<()> {
        if plan.member_present() {
            self.relationship_parts = checked_add_limit(
                self.relationship_parts,
                1,
                litchi_opc::ReadResource::RelationshipParts,
                limits.max_relationship_parts(),
            )?;
            self.xml_bytes = checked_add_limit(
                self.xml_bytes,
                plan.final_len(),
                litchi_opc::ReadResource::TotalRelationshipXmlBytes,
                limits.max_total_relationship_xml_bytes(),
            )?;
            self.xml_events = checked_add_limit(
                self.xml_events,
                plan.event_count(),
                litchi_opc::ReadResource::TotalRelationshipXmlEvents,
                limits.max_total_relationship_xml_events(),
            )?;
        }
        self.relationships = checked_add_limit(
            self.relationships,
            plan.relationship_count(),
            litchi_opc::ReadResource::TotalRelationships,
            limits.max_total_relationships(),
        )?;
        Ok(())
    }
}

fn index_relationships<'package, I>(
    relationships: I,
    source: &'package str,
    part_indices: &HashMap<String, usize>,
    incoming: &mut [Vec<RelationshipReference<'package>>],
    graph_seen: &mut [bool],
    budget: &mut RelationshipBudget,
    limits: ReadLimits,
) -> Result<()>
where
    I: Iterator<Item = &'package litchi_opc::Relationship>,
{
    for relationship in relationships {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        // OPC's graph traversal charges the canonical resolved target path in
        // addition to the authored Target attribute charged by source capture.
        // Keep that retained limit for changed in-memory relationships too.
        check_limit(
            litchi_opc::ReadResource::RelationshipTargetBytes,
            target.as_str().len(),
            limits.max_relationship_target_bytes(),
        )?;
        // `source_relationships_with_limits` charges the authored Target
        // attribute.  The resolved absolute PackURI can be longer because it
        // includes the source-part base path, so canonical graph lookup uses
        // the package member-name ceiling rather than charging that derived
        // spelling as another relationship Target attribute.
        let target_key = canonical_key(
            target.as_str(),
            limits
                .max_archive_member_name_bytes()
                .max(limits.max_relationship_target_bytes() as u64),
        )?;
        if let Some(&target_index) = part_indices.get(&target_key) {
            if !graph_seen[target_index] {
                budget.graph_nodes = checked_add_limit(
                    budget.graph_nodes,
                    1,
                    litchi_opc::ReadResource::RelationshipGraphNodes,
                    limits.max_relationship_graph_nodes(),
                )?;
                graph_seen[target_index] = true;
            }
            let bucket = incoming
                .get_mut(target_index)
                .ok_or_else(|| invalid("PivotTable relationship target index is incomplete"))?;
            bucket
                .try_reserve(1)
                .map_err(|source| allocation("PivotTable incoming relationship edges", source))?;
            bucket.push(RelationshipReference {
                source,
                relationship,
            });
        } else {
            if !budget.missing_nodes.contains(&target_key) {
                budget.graph_nodes = checked_add_limit(
                    budget.graph_nodes,
                    1,
                    litchi_opc::ReadResource::RelationshipGraphNodes,
                    limits.max_relationship_graph_nodes(),
                )?;
                budget.missing_nodes.try_reserve(1).map_err(|source| {
                    allocation("PivotTable missing relationship nodes", source)
                })?;
                budget.missing_nodes.insert(target_key);
            }
        }
    }
    Ok(())
}

fn compare_references<'package>(
    left: &RelationshipReference<'package>,
    right: &RelationshipReference<'package>,
) -> Ordering {
    left.source
        .cmp(right.source)
        .then_with(|| left.relationship.r_id().cmp(right.relationship.r_id()))
        .then_with(|| {
            left.relationship
                .reltype()
                .cmp(right.relationship.reltype())
        })
        .then_with(|| {
            left.relationship
                .target_ref()
                .cmp(right.relationship.target_ref())
        })
        .then_with(|| {
            target_mode_rank(left.relationship.target_mode())
                .cmp(&target_mode_rank(right.relationship.target_mode()))
        })
}

fn target_mode_rank(mode: litchi_opc::TargetMode) -> u8 {
    match mode {
        litchi_opc::TargetMode::Internal => 0,
        litchi_opc::TargetMode::External => 1,
    }
}

#[cfg(test)]
fn relationship_xml_metrics(bytes: &[u8]) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_reader(bytes);
    // Match `PackageReader::parse_rels_xml_with_source`, whose retained
    // ingress profile trims whitespace-only text events before charging the
    // XML event budget.  Counting with a different reader configuration would
    // reject a package at an exact aggregate event cap after ingress accepted
    // it, or admit one that ingress rejected at the same cap.
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut relationships = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("PivotTable relationship XML event count overflows"))?;
        match reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                relationships = relationships
                    .checked_add(1)
                    .ok_or_else(|| invalid("PivotTable relationship count overflows"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((events, relationships))
}

fn check_part_name_limit(name: &PackURI, limits: ReadLimits) -> Result<()> {
    let maximum = usize::try_from(limits.max_archive_member_name_bytes()).unwrap_or(usize::MAX);
    check_limit(
        litchi_opc::ReadResource::ArchiveMemberNameBytes,
        name.membername().len(),
        maximum,
    )
}

fn canonical_part_key(name: &PackURI, limits: ReadLimits) -> Result<String> {
    check_part_name_limit(name, limits)?;
    canonical_lower_key(name.as_str())
}

fn canonical_key(value: &str, maximum: u64) -> Result<String> {
    let maximum = usize::try_from(maximum).unwrap_or(usize::MAX);
    check_limit(
        litchi_opc::ReadResource::RelationshipTargetBytes,
        value.len(),
        maximum,
    )?;
    canonical_lower_key(value)
}

fn canonical_lower_key(value: &str) -> Result<String> {
    let mut key = String::new();
    key.try_reserve_exact(value.len())
        .map_err(|source| allocation("PivotTable canonical relationship key", source))?;
    key.push_str(value);
    key.make_ascii_lowercase();
    Ok(key)
}

fn copy_bounded(value: &str, resource: &'static str) -> Result<String> {
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    copy.push_str(value);
    Ok(copy)
}

fn ascii_case_insensitive_cmp(left: &str, right: &str) -> Ordering {
    for (left, right) in left.bytes().zip(right.bytes()) {
        let ordering = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if ordering != Ordering::Equal {
            return ordering;
        }
    }
    left.len().cmp(&right.len())
}

fn checked_add_limit(
    current: usize,
    additional: usize,
    resource: litchi_opc::ReadResource,
    maximum: usize,
) -> Result<usize> {
    let next = current
        .checked_add(additional)
        .ok_or_else(|| read_limit(resource, usize::MAX, maximum))?;
    if next > maximum {
        return Err(read_limit(resource, next, maximum));
    }
    Ok(next)
}

fn check_limit(resource: litchi_opc::ReadResource, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(read_limit(resource, actual, maximum));
    }
    Ok(())
}

fn read_limit(resource: litchi_opc::ReadResource, actual: usize, maximum: usize) -> Error {
    Error::Package(OpcError::ReadLimit {
        resource,
        actual: actual as u64,
        maximum: maximum as u64,
    })
}

fn allocation(resource: &'static str, source: std::collections::TryReserveError) -> Error {
    Error::Allocation { resource, source }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, ReadLimits, Relationships};
    use std::cmp::Ordering;
    use std::collections::HashMap;

    #[test]
    fn canonical_part_order_is_ascii_case_insensitive_and_stable() {
        assert_eq!(
            ascii_case_insensitive_cmp("/xl/A.xml", "/xl/a.xml"),
            Ordering::Equal
        );
        assert_eq!(
            ascii_case_insensitive_cmp("/xl/a.xml", "/xl/b.xml"),
            Ordering::Less
        );
        assert_eq!(
            ascii_case_insensitive_cmp("/xl/b.xml", "/xl/a.xml"),
            Ordering::Greater
        );
    }

    #[test]
    fn archive_part_name_budget_excludes_packuri_root_slash() {
        let name = PackURI::new("/abc").unwrap();
        let exact = ReadLimits::builder()
            .max_archive_member_name_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(check_part_name_limit(&name, exact).is_ok());

        let one_under = ReadLimits::builder()
            .max_archive_member_name_bytes(2)
            .unwrap()
            .build()
            .unwrap();
        assert!(check_part_name_limit(&name, one_under).is_err());
    }

    #[test]
    fn relationship_event_metrics_ignore_whitespace_text_like_ingress() {
        let compact = b"<Relationships><Relationship/></Relationships>";
        let spaced = b"<Relationships>\n  <Relationship/>\n</Relationships>";
        assert_eq!(
            relationship_xml_metrics(compact).unwrap(),
            relationship_xml_metrics(spaced).unwrap()
        );
        assert_eq!(relationship_xml_metrics(spaced).unwrap().1, 1);
    }

    #[test]
    fn graph_node_budget_accepts_exact_cap_and_rejects_one_over_before_mutation() {
        let limits = ReadLimits::builder()
            .max_relationship_graph_nodes(1)
            .unwrap()
            .build()
            .unwrap();
        let mut one_edge = Relationships::new("/".to_owned());
        one_edge.add_relationship(
            "urn:unknown".to_owned(),
            "/a".to_owned(),
            "rId1".to_owned(),
            false,
        );
        let mut part_indices = HashMap::new();
        part_indices.insert("/a".to_owned(), 0);
        let mut incoming = vec![Vec::new()];
        let mut graph_seen = vec![false];
        let mut budget = RelationshipBudget::default();
        index_relationships(
            one_edge.iter(),
            "/",
            &part_indices,
            &mut incoming,
            &mut graph_seen,
            &mut budget,
            limits,
        )
        .unwrap();
        assert_eq!(budget.graph_nodes, 1);
        assert!(graph_seen[0]);
        assert_eq!(incoming[0].len(), 1);

        let mut one_over = Relationships::new("/".to_owned());
        one_over.add_relationship(
            "urn:unknown".to_owned(),
            "/missing".to_owned(),
            "rId1".to_owned(),
            false,
        );
        assert!(
            index_relationships(
                one_over.iter(),
                "/",
                &part_indices,
                &mut incoming,
                &mut graph_seen,
                &mut budget,
                limits,
            )
            .is_err()
        );
        assert_eq!(budget.graph_nodes, 1);
        assert!(budget.missing_nodes.is_empty());
    }

    #[test]
    fn absent_part_relationships_keep_only_an_absence_token() {
        let mut package = OpcPackage::new();
        let owner = PackURI::new("/custom/no-rels.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/octet-stream".to_owned(),
            Vec::new(),
        )));

        let index = RelationshipIndex::build(&package).unwrap();

        assert!(index.source_ref(&owner).unwrap().is_none());
        assert!(index.source(&owner).unwrap().is_none());
    }
}
