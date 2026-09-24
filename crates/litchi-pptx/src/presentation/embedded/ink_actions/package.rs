//! OPC ownership for existing slide InkAction targets.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_drawingml::ink::actions::{self, Profile};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    OpcError, OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, Part, ReadResource,
    TargetMode,
};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::codec::{self, REQUIRED_CONTENT_TYPE};
use super::graph::RelationshipIndex;
use super::model::{
    ActionAnchor, AnchorFingerprint, Dialect, InboundReference, PackageRelationship, RawAnchor,
    Snapshot,
};
use super::transaction::{Commit, Patch};
use crate::presentation::embedded::{invalid, limit};
use crate::{Error, Result};

const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";

const MAX_RELATIONSHIP_FIELD_BYTES: usize = 4096;

pub use super::model::Limits;

struct SourceReadset {
    content_types: OwnedContentTypes,
    owner_relationships: OwnedRelationships,
    package_relationships: OwnedRelationships,
}

struct TargetCache {
    bytes: HashMap<PackURI, Arc<[u8]>>,
    profiles: HashMap<PackURI, Arc<Profile>>,
    inbound: HashMap<PackURI, Arc<[InboundReference]>>,
    outbound: HashMap<PackURI, Arc<[InboundReference]>>,
    retained_bytes: usize,
}

impl TargetCache {
    fn new() -> Self {
        Self {
            bytes: HashMap::new(),
            profiles: HashMap::new(),
            inbound: HashMap::new(),
            outbound: HashMap::new(),
            retained_bytes: 0,
        }
    }

    fn reserve(&mut self, capacity: usize) -> Result<()> {
        reserve_map(&mut self.bytes, capacity, "ink-action target byte cache")?;
        reserve_map(
            &mut self.profiles,
            capacity,
            "ink-action target profile cache",
        )?;
        reserve_map(
            &mut self.inbound,
            capacity,
            "ink-action target inbound graph cache",
        )?;
        reserve_map(
            &mut self.outbound,
            capacity,
            "ink-action target graph cache",
        )?;
        Ok(())
    }

    fn load(
        &mut self,
        package: &OpcPackage,
        target_name: &PackURI,
        limits: &Limits,
        relationship_index: &RelationshipIndex<'_>,
    ) -> Result<()> {
        if self.bytes.contains_key(target_name) {
            return Ok(());
        }
        let target = package.get_part(target_name).map_err(|error| {
            Error::PartNotFound(format!(
                "ink-action target '{}' is missing: {error}",
                target_name.as_str()
            ))
        })?;
        if target.blob().len() > limits.target_bytes {
            return Err(limit("ink-action target bytes", limits.target_bytes));
        }
        let retained = self
            .retained_bytes
            .checked_add(target.blob().len())
            .ok_or_else(|| {
                limit(
                    "ink-action aggregate target bytes",
                    limits.total_target_bytes,
                )
            })?;
        if retained > limits.total_target_bytes {
            return Err(limit(
                "ink-action aggregate target bytes",
                limits.total_target_bytes,
            ));
        }
        // Validate the complete graph closure before retaining target bytes or
        // parsing the shared profile.  This keeps an unsupported closure from
        // consuming owner-layer profile memory.
        let closure = relationship_index.closure(target_name)?;
        let inbound = Arc::from(closure.owned_inbound()?);
        let outbound = Arc::from(closure.owned_outbound()?);
        let profile = Arc::new(actions::read_profile(target.blob())?);
        // The shared DrawingML profile retains the exact source allocation.
        // Reuse it for the anchor/cache view instead of retaining the OPC
        // `Arc<Vec<u8>>` next to a second profile-source allocation.
        let bytes = profile.shared_source();
        self.reserve(1)?;
        self.retained_bytes = retained;
        self.bytes.insert(target_name.clone(), bytes);
        self.profiles.insert(target_name.clone(), profile);
        self.inbound.insert(target_name.clone(), inbound);
        self.outbound.insert(target_name.clone(), outbound);
        Ok(())
    }
}

fn reserve_map<K: Eq + std::hash::Hash, V>(
    map: &mut HashMap<K, V>,
    capacity: usize,
    resource: &'static str,
) -> Result<()> {
    map.try_reserve(capacity)
        .map_err(|source| Error::Allocation { resource, source })
}

/// Discover typed InkAction anchors from one existing slide owner.
pub fn load_slide(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
) -> Result<Vec<ActionAnchor>> {
    load_slide_catalog(package, slide_index, slide, limits).map(|(anchors, _)| anchors)
}

fn load_slide_catalog(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
) -> Result<(Vec<ActionAnchor>, Vec<RawAnchor>)> {
    load_slide_catalog_inner(package, slide_index, slide, limits, true)
}

fn load_slide_catalog_inner(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
    validate_shared_owners: bool,
) -> Result<(Vec<ActionAnchor>, Vec<RawAnchor>)> {
    limits.validate()?;
    let relationship_index = RelationshipIndex::build(package, limits)?;
    validate_relationship_sources(package)?;
    capture_source_readset(package, slide.partname())?;
    let mut target_cache = TargetCache::new();
    load_slide_catalog_with_index(
        package,
        slide_index,
        slide,
        limits,
        validate_shared_owners,
        &relationship_index,
        &mut target_cache,
    )
}

fn load_slide_catalog_with_index(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
    validate_shared_owners: bool,
    relationship_index: &RelationshipIndex<'_>,
    target_cache: &mut TargetCache,
) -> Result<(Vec<ActionAnchor>, Vec<RawAnchor>)> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "ink-action discovery requires a PresentationML slide",
        ));
    }
    let source_xml = slide.blob_arc();
    let source = source_xml.as_slice();
    let candidates = codec::scan_slide(source, limits.anchors)?;
    let candidate_capacity = candidates.len();
    let mut anchors = Vec::new();
    let mut raw_anchors = Vec::new();
    target_cache.reserve(candidate_capacity)?;
    anchors
        .try_reserve_exact(candidate_capacity)
        .map_err(|source| Error::Allocation {
            resource: "ink-action anchors",
            source,
        })?;
    raw_anchors
        .try_reserve_exact(candidate_capacity)
        .map_err(|source| Error::Allocation {
            resource: "ink-action raw anchors",
            source,
        })?;
    let mut semantic_ordinal = 0usize;

    for candidate in candidates {
        if !candidate.typed {
            let fingerprint =
                raw_fingerprint(package, slide, &candidate, source, relationship_index)?;
            raw_anchors.push(RawAnchor {
                source_ordinal: candidate.source_ordinal,
                source_xml: Arc::clone(&source_xml),
                span: candidate.span,
                fingerprint,
            });
            continue;
        }
        let relationship_id = candidate
            .relationship_id
            .as_deref()
            .ok_or_else(|| invalid("ink-action typed candidate has no relationship ID"))?;
        let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
            Error::Relationship(format!(
                "ink-action contentPart relationship '{}' is missing",
                relationship_id
            ))
        })?;
        let relationship_type = relationship.reltype();
        let expected_relationship_type = match candidate.dialect {
            Dialect::Transitional => rt::CUSTOM_XML,
            Dialect::Strict => STRICT_CUSTOM_XML,
        };
        if relationship_type != expected_relationship_type {
            return Err(Error::Relationship(format!(
                "ink-action contentPart relationship '{}' has unsupported type",
                relationship_id
            )));
        }
        if relationship.target_mode() != TargetMode::Internal || relationship.is_external() {
            return Err(Error::Relationship(format!(
                "ink-action contentPart relationship '{}' must have an internal target",
                relationship_id
            )));
        }
        if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
            return Err(Error::Relationship(format!(
                "ink-action contentPart relationship '{}' has an internal query or fragment",
                relationship_id
            )));
        }
        let target_name = relationship.target_partname().map_err(|error| {
            Error::Relationship(format!(
                "ink-action contentPart relationship '{}' has an invalid target: {error}",
                relationship_id
            ))
        })?;
        let target = package.get_part(&target_name).map_err(|error| {
            Error::PartNotFound(format!(
                "ink-action target '{}' is missing: {error}",
                target_name.as_str()
            ))
        })?;
        let target_name = target.partname().clone();
        if target.content_type() != REQUIRED_CONTENT_TYPE {
            return Err(Error::ContentType {
                expected: REQUIRED_CONTENT_TYPE.to_owned(),
                actual: target.content_type().to_owned(),
            });
        }
        target_cache.load(package, &target_name, limits, relationship_index)?;
        let profile = target_cache
            .profiles
            .get(&target_name)
            .cloned()
            .ok_or_else(|| invalid("ink-action target profile cache is missing"))?;
        let target_blob = target_cache
            .bytes
            .get(&target_name)
            .cloned()
            .ok_or_else(|| invalid("ink-action target byte cache is missing"))?;
        source
            .get(candidate.span.clone())
            .ok_or_else(|| invalid("ink-action AlternateContent span is invalid"))?;
        source
            .get(candidate.choice_span.clone())
            .ok_or_else(|| invalid("ink-action Choice span is invalid"))?;
        source
            .get(candidate.fallback_span.clone())
            .ok_or_else(|| invalid("ink-action Fallback span is invalid"))?;
        source
            .get(candidate.content_span.clone())
            .ok_or_else(|| invalid("ink-action contentPart span is invalid"))?;
        let inbound = target_cache
            .inbound
            .get(&target_name)
            .cloned()
            .ok_or_else(|| invalid("ink-action inbound graph cache is missing"))?;
        let outbound = target_cache
            .outbound
            .get(&target_name)
            .cloned()
            .ok_or_else(|| invalid("ink-action outbound graph cache is missing"))?;
        let fingerprint = AnchorFingerprint::from_closure(
            source
                .get(candidate.span.clone())
                .ok_or_else(|| invalid("ink-action AlternateContent span is invalid"))?,
            source
                .get(candidate.choice_span.clone())
                .ok_or_else(|| invalid("ink-action Choice span is invalid"))?,
            relationship.r_id().as_bytes(),
            relationship_type.as_bytes(),
            relationship.target_ref().as_bytes(),
            target_name.as_str().as_bytes(),
            target.blob(),
        )
        .with_graph(
            target.content_type().as_bytes(),
            relationship.target_mode(),
            &inbound,
            &outbound,
        );
        raw_anchors.push(RawAnchor {
            source_ordinal: candidate.source_ordinal,
            source_xml: Arc::clone(&source_xml),
            span: candidate.span.clone(),
            fingerprint,
        });
        anchors.push(ActionAnchor {
            slide_index,
            semantic_ordinal,
            source_ordinal: candidate.source_ordinal,
            fingerprint,
            source_xml: Arc::clone(&source_xml),
            owner_span: candidate.span,
            choice_span: candidate.choice_span,
            fallback_span: candidate.fallback_span,
            anchor_span: candidate.content_span,
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship_type.to_owned(),
            target_ref: relationship.target_ref().to_owned(),
            target_mode: relationship.target_mode(),
            target_part_name: target_name,
            content_type: target.content_type().to_owned(),
            target_bytes: target_blob,
            profile,
            inbound,
            outbound,
        });
        semantic_ordinal = semantic_ordinal
            .checked_add(1)
            .ok_or_else(|| limit("ink-action semantic ordinal", limits.anchors))?;
    }
    if validate_shared_owners {
        validate_shared_slide_owners(
            package,
            slide,
            &anchors,
            limits,
            relationship_index,
            target_cache,
        )?;
    }
    Ok((anchors, raw_anchors))
}

fn raw_fingerprint(
    package: &OpcPackage,
    slide: &dyn Part,
    candidate: &super::model::Candidate,
    source: &[u8],
    relationship_index: &RelationshipIndex<'_>,
) -> Result<AnchorFingerprint> {
    let Some(relationship_id) = candidate.relationship_id.as_deref() else {
        return Ok(candidate.fingerprint);
    };
    let Some(relationship) = slide.rels().get(relationship_id) else {
        return Ok(candidate.fingerprint);
    };
    let owner = source.get(candidate.span.clone()).unwrap_or_default();
    let choice = source
        .get(candidate.choice_span.clone())
        .unwrap_or_default();
    let Ok(target_ref) = relationship.target_partname() else {
        return Ok(AnchorFingerprint::from_closure(
            owner,
            choice,
            relationship.r_id().as_bytes(),
            relationship.reltype().as_bytes(),
            relationship.target_ref().as_bytes(),
            &[],
            &[],
        )
        .with_graph(b"", relationship.target_mode(), &[], &[]));
    };
    if relationship.target_query().is_some() || relationship.target_fragment().is_some() {
        // A query/fragment target is deliberately opaque.  Its target-side
        // relationship collection is outside the authority of an untouched
        // raw branch, which cannot be selected for a typed edit; the exact
        // owner XML and owner relationship token remain the source read set.
        return Ok(candidate.fingerprint);
    }
    // Only an absent target keeps the candidate fingerprint. A payload decode
    // refusal is not absence and is propagated, so it can never produce a
    // partial fingerprint (ADR 0030).
    let target = match package.get_part(&target_ref) {
        Ok(target) => target,
        Err(OpcError::PartNotFound(_)) => return Ok(candidate.fingerprint),
        Err(error) => return Err(error.into()),
    };
    let target_name = target.partname().clone();
    // Source selectors still authorize the complete relationship closure.
    // Fail closed when that closure cannot be read; a partial fingerprint
    // would make a stale selector appear current.
    let closure = relationship_index.closure(&target_name)?;
    let inbound = closure.owned_inbound()?;
    let outbound = closure.owned_outbound()?;
    Ok(AnchorFingerprint::from_closure(
        owner,
        choice,
        relationship.r_id().as_bytes(),
        relationship.reltype().as_bytes(),
        relationship.target_ref().as_bytes(),
        target_name.as_str().as_bytes(),
        target.blob(),
    )
    .with_graph(
        target.content_type().as_bytes(),
        relationship.target_mode(),
        &inbound,
        &outbound,
    ))
}

fn validate_shared_slide_owners(
    package: &OpcPackage,
    selected_slide: &dyn Part,
    anchors: &[ActionAnchor],
    limits: &Limits,
    relationship_index: &RelationshipIndex<'_>,
    target_cache: &mut TargetCache,
) -> Result<()> {
    let mut owners = HashSet::<PackURI>::new();
    for anchor in anchors {
        for reference in anchor.inbound_references() {
            if reference.source_part == *selected_slide.partname()
                || reference.source_part.as_str() == "/"
                || !matches!(
                    reference.relationship_type.as_str(),
                    rt::CUSTOM_XML | STRICT_CUSTOM_XML
                )
            {
                continue;
            }
            owners.insert(reference.source_part.clone());
        }
    }
    for owner_name in owners {
        let owner = package.get_part(&owner_name)?;
        if owner.content_type() != ct::PML_SLIDE {
            continue;
        }
        capture_owner_relationship_source(package, owner.partname())?;
        load_slide_catalog_with_index(
            package,
            0,
            owner,
            limits,
            false,
            relationship_index,
            target_cache,
        )?;
    }
    Ok(())
}

/// Load a source-backed snapshot for one slide.
pub fn load_snapshot(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
) -> Result<Snapshot> {
    limits.validate()?;
    let relationship_index = RelationshipIndex::build(package, limits)?;
    validate_relationship_sources(package)?;
    let mut target_cache = TargetCache::new();
    load_snapshot_with_index(
        package,
        slide_index,
        slide,
        limits,
        &relationship_index,
        &mut target_cache,
    )
}

/// Load several slide snapshots while reusing one borrowed package graph.
///
/// The public `Presentation` facade calls this for an ordinary multi-slide
/// inventory.  Relationship closure validation remains per selected slide,
/// while package-wide graph indexing is performed once for the batch.
pub(crate) fn load_snapshots<'a, I>(
    package: &OpcPackage,
    slides: I,
    limits: &Limits,
) -> Result<Vec<Snapshot>>
where
    I: IntoIterator<Item = (usize, &'a dyn Part)>,
{
    limits.validate()?;
    let relationship_index = RelationshipIndex::build(package, limits)?;
    validate_relationship_sources(package)?;
    let mut snapshots = Vec::new();
    let iterator = slides.into_iter();
    if let Some(upper) = iterator.size_hint().1 {
        snapshots
            .try_reserve_exact(upper)
            .map_err(|source| Error::Allocation {
                resource: "ink-action slide snapshots",
                source,
            })?;
    }
    let mut target_cache = TargetCache::new();
    for (slide_index, slide) in iterator {
        snapshots.push(load_snapshot_with_index(
            package,
            slide_index,
            slide,
            limits,
            &relationship_index,
            &mut target_cache,
        )?);
    }
    Ok(snapshots)
}

fn load_snapshot_with_index(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &Limits,
    relationship_index: &RelationshipIndex<'_>,
    target_cache: &mut TargetCache,
) -> Result<Snapshot> {
    let source_readset = capture_source_readset(package, slide.partname())?;
    let (anchors, raw_anchors) = load_slide_catalog_with_index(
        package,
        slide_index,
        slide,
        limits,
        true,
        relationship_index,
        target_cache,
    )?;
    let mut owner_relationships = Vec::new();
    owner_relationships
        .try_reserve_exact(slide.rels().len())
        .map_err(|source| Error::Allocation {
            resource: "ink-action owner relationships",
            source,
        })?;
    for relationship in slide.rels().iter() {
        push_reference(
            &mut owner_relationships,
            slide.partname(),
            relationship,
            package.read_limits().max_relationships_per_part(),
        )?;
    }
    sort_references(&mut owner_relationships);
    let package_relationships = package_relationships(package.rels().iter())?;
    Ok(Snapshot::from_parts(
        slide_index,
        slide.partname().clone(),
        slide.blob_arc(),
        anchors,
        raw_anchors,
        owner_relationships,
        source_readset.content_types,
        source_readset.owner_relationships,
        source_readset.package_relationships,
        package_relationships,
        package.read_limits(),
        *limits,
    ))
}

fn capture_source_readset(package: &OpcPackage, owner: &PackURI) -> Result<SourceReadset> {
    let read_limits = package.read_limits();
    let root = PackURI::new("/").map_err(OpcError::InvalidPackUri)?;
    Ok(SourceReadset {
        content_types: package.source_content_types_with_limits(read_limits)?,
        owner_relationships: package.source_relationships_with_limits(owner, read_limits)?,
        package_relationships: package.source_relationships_with_limits(&root, read_limits)?,
    })
}

/// Validate every currently published `.rels` member under the package's
/// ingress policy before the owner builds an inventory.  `source_*` capture
/// validates one member's exact provenance and parser bounds, but its
/// per-call aggregate counters cannot account for unrelated members.  The
/// owner therefore performs one bounded package-wide pass here, including
/// relationship members which are unrelated to the selected slide.
fn validate_relationship_sources(package: &OpcPackage) -> Result<()> {
    let limits = package.read_limits();
    let root = PackURI::new("/").map_err(OpcError::InvalidPackUri)?;
    let mut relationship_parts = 0usize;
    let mut total_xml_bytes = 0usize;
    let mut total_xml_events = 0usize;
    let mut total_relationships = 0usize;

    let root_source = package.source_relationships_with_limits(&root, limits)?;
    validate_relationship_source(
        &root_source,
        limits,
        &mut relationship_parts,
        &mut total_xml_bytes,
        &mut total_xml_events,
        &mut total_relationships,
    )?;

    for part in package.iter_parts() {
        let source = package.source_relationships_with_limits(part.partname(), limits)?;
        validate_relationship_source(
            &source,
            limits,
            &mut relationship_parts,
            &mut total_xml_bytes,
            &mut total_xml_events,
            &mut total_relationships,
        )?;
    }
    Ok(())
}

fn validate_relationship_source(
    source: &OwnedRelationships,
    limits: litchi_opc::ReadLimits,
    relationship_parts: &mut usize,
    total_xml_bytes: &mut usize,
    total_xml_events: &mut usize,
    total_relationships: &mut usize,
) -> Result<()> {
    if !source.member_present() {
        return Ok(());
    }
    *relationship_parts = relationship_parts
        .checked_add(1)
        .ok_or_else(|| invalid("relationship part count overflows"))?;
    check_relationship_limit(
        ReadResource::RelationshipParts,
        *relationship_parts,
        limits.max_relationship_parts(),
    )?;

    *total_xml_bytes = total_xml_bytes
        .checked_add(source.bytes().len())
        .ok_or_else(|| invalid("relationship XML byte count overflows"))?;
    check_relationship_limit(
        ReadResource::TotalRelationshipXmlBytes,
        *total_xml_bytes,
        limits.max_total_relationship_xml_bytes(),
    )?;

    let (events, relationships) = relationship_xml_metrics(source.bytes())?;
    *total_xml_events = total_xml_events
        .checked_add(events)
        .ok_or_else(|| invalid("relationship XML event count overflows"))?;
    check_relationship_limit(
        ReadResource::TotalRelationshipXmlEvents,
        *total_xml_events,
        limits.max_total_relationship_xml_events(),
    )?;
    *total_relationships = total_relationships
        .checked_add(relationships)
        .ok_or_else(|| invalid("relationship count overflows"))?;
    check_relationship_limit(
        ReadResource::TotalRelationships,
        *total_relationships,
        limits.max_total_relationships(),
    )?;
    Ok(())
}

fn relationship_xml_metrics(xml: &[u8]) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut relationships = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("relationship XML event count overflows"))?;
        match reader.read_event()? {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"Relationship" =>
            {
                relationships = relationships
                    .checked_add(1)
                    .ok_or_else(|| invalid("relationship count overflows"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((events, relationships))
}

fn check_relationship_limit(resource: ReadResource, actual: usize, maximum: usize) -> Result<()> {
    if actual > maximum {
        return Err(Error::Opc(OpcError::ReadLimit {
            resource,
            actual: actual as u64,
            maximum: maximum as u64,
        }));
    }
    Ok(())
}

fn capture_owner_relationship_source(package: &OpcPackage, owner: &PackURI) -> Result<()> {
    let read_limits = package.read_limits();
    package.source_relationships_with_limits(owner, read_limits)?;
    Ok(())
}

fn package_relationships<'a>(
    relationships: impl Iterator<Item = &'a litchi_opc::Relationship>,
) -> Result<Vec<PackageRelationship>> {
    let mut values = Vec::new();
    if let Some(upper) = relationships.size_hint().1 {
        values
            .try_reserve_exact(upper)
            .map_err(|source| Error::Allocation {
                resource: "ink-action package relationships",
                source,
            })?;
    }
    for relationship in relationships {
        values.push(PackageRelationship {
            relationship_id: relationship.r_id().to_owned(),
            relationship_type: relationship.reltype().to_owned(),
            target_ref: relationship.target_ref().to_owned(),
            target_mode: relationship.target_mode(),
        });
    }
    values.sort_by(|left, right| {
        left.relationship_id
            .cmp(&right.relationship_id)
            .then_with(|| left.relationship_type.cmp(&right.relationship_type))
            .then_with(|| left.target_ref.cmp(&right.target_ref))
            .then_with(|| {
                target_mode_rank(left.target_mode).cmp(&target_mode_rank(right.target_mode))
            })
    });
    Ok(values)
}

/// Apply a source-checked existing-target patch atomically.
pub fn apply_patch(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    let before = patch.before();
    let slide = package.get_part(&before.slide_part_name)?;
    let current = load_snapshot(package, before.slide_index, slide, &before.limits())?;
    if !current.same_source(before) {
        return Err(Error::StaleSource);
    }
    if patch.is_empty() {
        return Ok(current);
    }
    if package.is_signed() || package.requires_signature_edit_policy() {
        return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
    }

    preflight_patch_limits(package, patch)?;
    let mut staged = package.clone();
    install_patch(&mut staged, patch)?;
    let slide = staged.get_part(&before.slide_part_name)?;
    let resulting = load_snapshot(&staged, before.slide_index, slide, &before.limits())?;
    if !resulting.same_source(patch.after()) {
        return Err(invalid(
            "published ink-action graph differs from the commit",
        ));
    }
    *package = staged;
    Ok(resulting)
}

/// Apply a committed existing-target transaction atomically.
#[inline]
pub fn apply_commit(package: &mut OpcPackage, commit: Commit) -> Result<Snapshot> {
    apply_patch(package, commit.patch())
}

fn install_patch(package: &mut OpcPackage, patch: &Patch) -> Result<()> {
    let before = patch.before();
    let after = patch.after();
    let owner_changed = package.get_part(&before.slide_part_name)?.blob() != after.source_xml();
    if owner_changed {
        // Slide snapshots retain the package part's `Arc<Vec<u8>>`; built-in
        // OPC parts adopt that allocation through `set_blob_shared`.
        // `preflight_patch_limits` has already admitted its exact size.
        package
            .get_part_mut(&before.slide_part_name)?
            .set_blob_shared(Arc::clone(&after.source_xml));
    }

    let mut seen = HashSet::<PackURI>::new();
    for anchor in after.anchors() {
        if !seen.insert(anchor.target_part_name().clone()) {
            continue;
        }
        let (target_content_type, target_bytes_equal) = {
            let part = package.get_part(anchor.target_part_name())?;
            (
                part.content_type().to_owned(),
                part.blob() == anchor.target_bytes(),
            )
        };
        if target_content_type != anchor.content_type() {
            return Err(Error::ContentType {
                expected: anchor.content_type().to_owned(),
                actual: target_content_type,
            });
        }
        if !target_bytes_equal {
            // The OPC `Part` API owns `Arc<Vec<u8>>`, while the action cache
            // intentionally shares the profile's `Arc<[u8]>`.  Convert only
            // at this publication boundary, after the staged per-part and
            // aggregate read limits have admitted the exact size.
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(anchor.target_bytes().len())
                .map_err(|source| Error::Allocation {
                    resource: "ink-action published target bytes",
                    source,
                })?;
            bytes.extend_from_slice(anchor.target_bytes());
            package
                .get_part_mut(anchor.target_part_name())?
                .set_blob(bytes);
        }
    }
    Ok(())
}

fn preflight_patch_limits(package: &OpcPackage, patch: &Patch) -> Result<()> {
    let read_limits = package.read_limits();
    let part_count = package.part_count();
    if part_count > read_limits.max_parts() {
        return Err(Error::Opc(OpcError::ReadLimit {
            resource: ReadResource::Parts,
            actual: part_count as u64,
            maximum: read_limits.max_parts() as u64,
        }));
    }
    let before = patch.before();
    let after = patch.after();
    let mut targets = HashMap::<PackURI, Arc<[u8]>>::new();
    targets
        .try_reserve(after.anchors().len())
        .map_err(|source| Error::Allocation {
            resource: "ink-action patch target index",
            source,
        })?;
    for anchor in after.anchors() {
        targets
            .entry(anchor.target_part_name().clone())
            .or_insert_with(|| Arc::clone(&anchor.target_bytes));
    }
    let mut total = 0u64;
    for part in package.iter_parts() {
        let size = if part.partname() == &before.slide_part_name {
            after.source_xml().len()
        } else if let Some(bytes) = targets.get(part.partname()) {
            bytes.len()
        } else {
            // Only a part the patch leaves in place is measured by its
            // current payload, so only such a part is decoded (ADR 0030).
            package.get_part(part.partname())?.blob().len()
        };
        let actual =
            u64::try_from(size).map_err(|_| invalid("ink-action staged part bytes exceed u64"))?;
        if actual > read_limits.max_part_bytes() {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                actual,
                maximum: read_limits.max_part_bytes(),
            }));
        }
        total = total
            .checked_add(actual)
            .ok_or_else(|| invalid("ink-action staged package bytes overflow"))?;
        if total > read_limits.max_total_part_bytes() {
            return Err(Error::Opc(OpcError::ReadLimit {
                resource: ReadResource::TotalPartBytes,
                actual: total,
                maximum: read_limits.max_total_part_bytes(),
            }));
        }
    }
    Ok(())
}

fn sort_references(references: &mut [InboundReference]) {
    references.sort_by(|left, right| {
        left.source_part
            .as_str()
            .cmp(right.source_part.as_str())
            .then_with(|| left.relationship_id.cmp(&right.relationship_id))
            .then_with(|| left.relationship_type.cmp(&right.relationship_type))
            .then_with(|| left.target_ref.cmp(&right.target_ref))
            .then_with(|| {
                target_mode_rank(left.target_mode).cmp(&target_mode_rank(right.target_mode))
            })
    });
}

const fn target_mode_rank(mode: TargetMode) -> u8 {
    match mode {
        TargetMode::Internal => 0,
        TargetMode::External => 1,
    }
}

fn push_reference(
    output: &mut Vec<InboundReference>,
    source: &PackURI,
    relationship: &litchi_opc::Relationship,
    maximum: usize,
) -> Result<()> {
    if output.len() >= maximum {
        return Err(limit("ink-action relationship edges", maximum));
    }
    output.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "ink-action relationship edges",
        source,
    })?;
    for value in [
        relationship.r_id(),
        relationship.reltype(),
        relationship.target_ref(),
    ] {
        if value.len() > MAX_RELATIONSHIP_FIELD_BYTES {
            return Err(limit(
                "ink-action relationship metadata bytes",
                MAX_RELATIONSHIP_FIELD_BYTES,
            ));
        }
    }
    output.push(InboundReference {
        source_part: source.clone(),
        relationship_id: relationship.r_id().to_owned(),
        relationship_type: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
        target_mode: relationship.target_mode(),
    });
    Ok(())
}

/// Expose the bounded owner constants for matrix/evidence tests.
#[must_use]
pub const fn default_limits() -> Limits {
    Limits::DEFAULT
}

/// The generic action target policy currently accepts only explicitly typed
/// `text/xml`; it intentionally does not infer an action MIME from InkML.
#[must_use]
pub const fn generic_content_type() -> &'static str {
    REQUIRED_CONTENT_TYPE
}
