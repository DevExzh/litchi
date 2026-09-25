//! Source-bound eager relationship edits and exact restoration.

use super::{
    OpcPackage, RelationshipBinding, SourceRelationshipsXml,
    check_canonical_relationship_attribute_limits, check_canonical_relationship_structure,
    check_relationship_capture_limits,
};
use crate::rel::Relationships;
use crate::source_backed::escaped_xml_attribute_len;
use crate::xml_attributes::BytesStartExt as _;
use crate::{
    OpcError, OwnedElementEdit, OwnedElementUpdate, OwnedXmlPart, PackURI, ReadLimits,
    ReadResource, Result, TargetMode,
};
use litchi_core::xml::ReaderOrigin;
use quick_xml::{events::Event, reader::NsReader};
use std::collections::HashSet;
use std::ops::Range;
use std::sync::Arc;

/// Immutable relationship XML bound to its package or part owner.
///
/// Capture with [`OpcPackage::source_relationships`]. Clones share source bytes;
/// an original token can restore its exact XML after a changed save/reopen.
/// `source_relationships_with_limits` applies the caller's bounded read
/// profile; the compatibility wrapper uses the default profile. Part-member
/// presence is retained, including explicit empty members. The package root
/// always has a relationship member in authored publication.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OwnedRelationships {
    owner: PackURI,
    xml: OwnedXmlPart,
    member_present: bool,
}

/// Output-free metrics for the current relationship collection of one owner.
///
/// This is descriptive admission data, not a source token or a materialize
/// authority. It is bound to the owner's current in-memory relationship
/// collection and must be recomputed before a later capture if that
/// collection changes. Callers may use the exact values to precharge an
/// operation aggregate, then call [`OpcPackage::source_relationships_with_limits`]
/// to perform the fresh bounded capture and parser checks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RelationshipSourcePlan {
    final_len: usize,
    event_count: usize,
    relationship_count: usize,
    member_present: bool,
}

impl RelationshipSourcePlan {
    /// Exact retained or canonical relationship-member byte length.
    #[must_use]
    pub const fn final_len(self) -> usize {
        self.final_len
    }

    /// Exact relationship-ingress quick-XML event count, including EOF.
    /// Relationship ingress uses `trim_text(true)`; raw source-publication
    /// validation applies its separate per-member event ceiling.
    #[must_use]
    pub const fn event_count(self) -> usize {
        self.event_count
    }

    /// Exact typed relationship-element count in the current graph.
    #[must_use]
    pub const fn relationship_count(self) -> usize {
        self.relationship_count
    }

    /// Whether this owner publishes a relationship member. An explicit empty
    /// source member is `true`; an owner with no source member and no edges is
    /// `false`.
    #[must_use]
    pub const fn member_present(self) -> bool {
        self.member_present
    }
}

/// One borrowed relationship edge to add during a source-bound edit plan.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RelationshipEdit<'a> {
    /// Relationship identifier written to the `Id` attribute.
    pub id: &'a str,
    /// Relationship type URI.
    pub reltype: &'a str,
    /// Internal relative target or external target URI.
    pub target: &'a str,
    /// Whether the target is internal or external.
    pub mode: TargetMode,
}

/// A borrowed dry-run for a newly authored relationship member.
///
/// This plan deliberately has no source token or owner provenance.  It is for
/// graph operations that create a relationship owner during the current
/// transaction, where no source `.rels` member exists to bind to.  The
/// canonical XML length, relationship count, and event count are computed
/// before any XML buffer is allocated.  Call [`Self::materialize`] only after
/// the caller has admitted those exact metrics against its aggregate limits.
#[derive(Debug)]
pub struct CanonicalRelationshipsPlan<'source> {
    relationships: &'source Relationships,
    final_len: usize,
    relationship_count: usize,
    event_count: usize,
}

impl<'source> CanonicalRelationshipsPlan<'source> {
    /// Exact canonical relationship-member byte length.
    #[must_use]
    pub const fn final_len(&self) -> usize {
        self.final_len
    }

    /// Exact number of relationship edges in the canonical member.
    #[must_use]
    pub const fn relationship_count(&self) -> usize {
        self.relationship_count
    }

    /// Exact quick-xml event count, including declaration, root, children,
    /// root close, and EOF.
    #[must_use]
    pub const fn event_count(&self) -> usize {
        self.event_count
    }

    /// Materialize the canonical XML member after rechecking the supplied
    /// limits.  This performs the sole final XML allocation and returns raw
    /// bytes without inventing source or owner provenance.
    pub fn materialize(self, limits: ReadLimits) -> Result<Vec<u8>> {
        check_canonical_relationship_plan_limits(
            self.relationships,
            self.final_len,
            self.event_count,
            limits,
        )?;
        let bytes = self.relationships.try_to_xml_bytes()?;
        if bytes.len() != self.final_len {
            return Err(invalid(
                "canonical relationship plan length differs from output",
            ));
        }
        Ok(bytes)
    }
}

impl Relationships {
    /// Plan a canonical relationship member without source XML or provenance.
    ///
    /// This is the neutral bridge for newly authored graph owners.  It shares
    /// the canonical serializer, escaping, validation, and caller `ReadLimits`
    /// checks used by OPC package capture, while keeping the final byte buffer
    /// deferred until [`CanonicalRelationshipsPlan::materialize`]. Source-bound
    /// capture adds its owned XML fixed-size ceiling before materialization.
    pub fn plan_canonical<'source>(
        &'source self,
        limits: ReadLimits,
    ) -> Result<CanonicalRelationshipsPlan<'source>> {
        // Validate all authored fields before the length pass can lead to a
        // later canonical XML allocation. The serializer escapes XML syntax,
        // but it cannot make XML 1.0-forbidden Unicode scalar values valid.
        validate_canonical_relationship_values(self, limits)?;
        let final_len = crate::source_backed::canonical_relationship_xml_len(self)?;
        let event_count = self
            .len()
            .checked_add(4)
            .ok_or_else(|| invalid("canonical relationship event count overflows"))?;
        check_canonical_relationship_plan_limits(self, final_len, event_count, limits)?;
        Ok(CanonicalRelationshipsPlan {
            relationships: self,
            final_len,
            relationship_count: self.len(),
            event_count,
        })
    }
}

fn validate_canonical_relationship_values(
    relationships: &Relationships,
    limits: ReadLimits,
) -> Result<()> {
    for relationship in relationships.iter() {
        if relationship.r_id().is_empty()
            || relationship.reltype().is_empty()
            || relationship.target_ref().is_empty()
        {
            return Err(invalid("relationship fields must not be empty"));
        }
        if !crate::pkgreader::is_xml_id(relationship.r_id()) {
            return Err(invalid("relationship Id is not an XML ID"));
        }
        if relationship.reltype().chars().any(char::is_whitespace)
            || relationship.reltype().chars().any(char::is_control)
            || relationship.target_ref().chars().any(char::is_control)
        {
            return Err(invalid(
                "relationship Type or Target is not a valid URI reference",
            ));
        }
        if [
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        ]
        .into_iter()
        .any(|value| {
            value
                .chars()
                .any(|character| !crate::xml_splice::xml10_character(character))
        }) {
            return Err(invalid(
                "relationship fields contain a forbidden XML 1.0 character",
            ));
        }
        limits.check(
            ReadResource::RelationshipTargetBytes,
            relationship.target_ref().len() as u64,
            limits.max_relationship_target_bytes() as u64,
        )?;
    }
    Ok(())
}

fn check_canonical_relationship_plan_limits(
    relationships: &Relationships,
    final_len: usize,
    event_count: usize,
    limits: ReadLimits,
) -> Result<()> {
    check_relationship_capture_limits(relationships, limits)?;
    check_canonical_relationship_attribute_limits(relationships, limits)?;
    check_canonical_relationship_structure(relationships, final_len, limits)?;
    limits.check(
        ReadResource::RelationshipXmlBytes,
        final_len as u64,
        limits.max_relationship_xml_bytes() as u64,
    )?;
    limits.check(
        ReadResource::PartBytes,
        final_len as u64,
        limits.max_part_bytes(),
    )?;
    limits.check(
        ReadResource::ArchiveEntryBytes,
        final_len as u64,
        limits.max_archive_entry_bytes(),
    )?;
    limits.check(
        ReadResource::XmlEvents,
        event_count as u64,
        limits.max_xml_events() as u64,
    )?;
    Ok(())
}

#[derive(Clone, Debug)]
struct OwnedRelationshipEdit {
    id: String,
    reltype: String,
    target: String,
    mode: TargetMode,
}

#[derive(Debug)]
struct RelationshipSourceScan {
    root_name: Vec<u8>,
    relationship_name: Vec<u8>,
    root_tag: Range<usize>,
    root_close: Option<usize>,
    root_empty: bool,
    removal_ranges: Vec<Range<usize>>,
    removed_bytes: usize,
    removed_events: usize,
    relationship_count: usize,
    event_count: usize,
}

/// A source-bound relationship edit admission.
///
/// The plan scans the source member and computes exact final XML bytes,
/// relationship count, event count, and member presence without constructing
/// the candidate XML.  `materialize` performs one bounded output allocation
/// and reparses the result before returning a new source token.
#[derive(Debug)]
pub struct RelationshipsEditPlan<'source> {
    source: &'source OwnedRelationships,
    additions: Vec<OwnedRelationshipEdit>,
    scan: RelationshipSourceScan,
    final_len: usize,
    relationship_count: usize,
    event_count: usize,
    member_present: bool,
}

impl<'source> RelationshipsEditPlan<'source> {
    /// Exact final relationship-member byte length.
    #[must_use]
    pub const fn final_len(&self) -> usize {
        self.final_len
    }

    /// Exact final number of relationship edges.
    #[must_use]
    pub const fn relationship_count(&self) -> usize {
        self.relationship_count
    }

    /// Exact final quick-xml event count, including `Event::Eof`.
    #[must_use]
    pub const fn event_count(&self) -> usize {
        self.event_count
    }

    /// Whether the final package publishes this relationship member.
    #[must_use]
    pub const fn member_present(&self) -> bool {
        self.member_present
    }

    /// Source token used by the plan for exact no-op and provenance checks.
    #[must_use]
    pub fn source(&self) -> &'source OwnedRelationships {
        self.source
    }

    /// Materialize one final relationship XML buffer after caller aggregate
    /// limits have admitted the exact metrics.
    pub fn materialize(self, limits: ReadLimits) -> Result<OwnedRelationships> {
        limits.check(
            ReadResource::RelationshipXmlBytes,
            self.final_len as u64,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            self.final_len as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            self.final_len as u64,
            limits.max_archive_entry_bytes(),
        )?;
        limits.check(
            ReadResource::XmlEvents,
            self.event_count as u64,
            limits.max_xml_events() as u64,
        )?;
        let source = self.source;
        if self.additions.is_empty() && self.scan.removal_ranges.is_empty() {
            return Ok(source.clone());
        }

        let mut output = Vec::new();
        output
            .try_reserve_exact(self.final_len)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship edit output",
                source,
            })?;
        let insertion = if self.scan.root_empty {
            self.scan
                .root_tag
                .end
                .checked_sub(2)
                .ok_or_else(|| invalid("empty relationships root is incomplete"))?
        } else {
            self.scan
                .root_close
                .ok_or_else(|| invalid("relationships root is unclosed"))?
        };
        let mut cursor = 0usize;
        let mut inserted = false;
        for range in &self.scan.removal_ranges {
            if !inserted && !self.additions.is_empty() && insertion <= range.start {
                output.extend_from_slice(&source.bytes()[cursor..insertion]);
                if self.scan.root_empty {
                    output.push(b'>');
                }
                append_owned_relationship_chunks(
                    &mut output,
                    &self.scan.relationship_name,
                    &self.additions,
                )?;
                if self.scan.root_empty {
                    output.extend_from_slice(b"</");
                    output.extend_from_slice(&self.scan.root_name);
                    output.push(b'>');
                    cursor = self.scan.root_tag.end;
                } else {
                    cursor = insertion;
                }
                inserted = true;
            }
            output.extend_from_slice(&source.bytes()[cursor..range.start]);
            cursor = range.end;
        }
        if !inserted && !self.additions.is_empty() {
            output.extend_from_slice(&source.bytes()[cursor..insertion]);
            if self.scan.root_empty {
                output.push(b'>');
            }
            append_owned_relationship_chunks(
                &mut output,
                &self.scan.relationship_name,
                &self.additions,
            )?;
            if self.scan.root_empty {
                output.extend_from_slice(b"</");
                output.extend_from_slice(&self.scan.root_name);
                output.push(b'>');
                cursor = self.scan.root_tag.end;
            } else {
                cursor = insertion;
            }
        }
        output.extend_from_slice(&source.bytes()[cursor..]);
        if output.len() != self.final_len {
            return Err(invalid("relationship edit length differs from plan"));
        }
        let xml = OwnedXmlPart::capture_with_limits(
            source.xml.name.clone(),
            source.xml.content_type.clone(),
            Arc::new(output),
            limits,
        )?;
        crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            xml.bytes(),
            &source.owner,
            limits,
        )?;
        Ok(OwnedRelationships {
            owner: source.owner.clone(),
            xml,
            member_present: self.member_present,
        })
    }
}

impl OwnedRelationships {
    /// Plan source-preserving relationship additions and removals in one pass.
    ///
    /// Existing comments, declarations, namespace aliases, attribute order,
    /// whitespace, and unrelated relationship elements remain byte-for-byte
    /// intact.  Removal selectors that do not match are idempotent no-ops.
    pub fn plan_edit<'source>(
        &'source self,
        additions: &[RelationshipEdit<'_>],
        removals: &[&str],
        limits: ReadLimits,
    ) -> Result<RelationshipsEditPlan<'source>> {
        limits.check(
            ReadResource::RelationshipXmlBytes,
            self.bytes().len() as u64,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            self.bytes().len() as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            self.bytes().len() as u64,
            limits.max_archive_entry_bytes(),
        )?;
        if additions.len() > limits.max_relationships_per_part() {
            return Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipsPerPart,
                actual: additions.len() as u64,
                maximum: limits.max_relationships_per_part() as u64,
            });
        }
        let mut raw_addition_bytes = 0usize;
        let mut addition_ids = HashSet::<&str>::new();
        addition_ids
            .try_reserve(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship edit addition IDs",
                source,
            })?;
        for addition in additions {
            if addition.id.is_empty() || addition.reltype.is_empty() || addition.target.is_empty() {
                return Err(invalid("relationship fields must not be empty"));
            }
            if !crate::pkgreader::is_xml_id(addition.id) {
                return Err(invalid("relationship Id is not an XML ID"));
            }
            if addition.reltype.chars().any(char::is_whitespace)
                || addition.reltype.chars().any(char::is_control)
                || addition.target.chars().any(char::is_control)
            {
                return Err(invalid(
                    "relationship Type or Target is not a valid URI reference",
                ));
            }
            let escaped_id_len = escaped_xml_attribute_len(addition.id)?;
            let escaped_reltype_len = escaped_xml_attribute_len(addition.reltype)?;
            let escaped_target_len = escaped_xml_attribute_len(addition.target)?;
            limits.check(
                ReadResource::XmlAttributeBytes,
                escaped_id_len as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            limits.check(
                ReadResource::XmlAttributeBytes,
                escaped_reltype_len as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            limits.check(
                ReadResource::XmlAttributeBytes,
                escaped_target_len as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            limits.check(
                ReadResource::RelationshipTargetBytes,
                addition.target.len() as u64,
                limits.max_relationship_target_bytes() as u64,
            )?;
            if !addition_ids.insert(addition.id) {
                return Err(invalid("relationship ID already exists"));
            }
            raw_addition_bytes = raw_addition_bytes
                .checked_add(escaped_id_len)
                .and_then(|size| size.checked_add(escaped_reltype_len))
                .and_then(|size| size.checked_add(escaped_target_len))
                .ok_or_else(|| invalid("relationship edit selector bytes overflow"))?;
        }
        for id in removals {
            limits.check(
                ReadResource::XmlAttributeBytes,
                id.len() as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
            raw_addition_bytes = raw_addition_bytes
                .checked_add(id.len())
                .ok_or_else(|| invalid("relationship edit selector bytes overflow"))?;
        }
        limits.check(
            ReadResource::RelationshipXmlBytes,
            raw_addition_bytes as u64,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        let parsed = crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            self.bytes(),
            &self.owner,
            limits,
        )?;
        let maximum_relationships = limits.max_relationships_per_part();
        let mut selected = HashSet::<&str>::new();
        selected
            .try_reserve(removals.len().min(maximum_relationships))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship edit removal IDs",
                source,
            })?;
        for id in removals {
            if selected.contains(id) {
                return Err(invalid("relationship removal IDs contain a duplicate"));
            }
            if selected.len() >= maximum_relationships {
                return Err(invalid(
                    "relationship removal batch exceeds the package limit",
                ));
            }
            selected.insert(*id);
        }
        let mut owned_additions = Vec::new();
        owned_additions
            .try_reserve_exact(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship edit additions",
                source,
            })?;
        for addition in additions {
            if parsed.get(addition.id).is_some() && !selected.contains(addition.id) {
                return Err(invalid("relationship ID already exists"));
            }
            owned_additions.push(OwnedRelationshipEdit {
                id: clone_string_bounded(addition.id, "OPC relationship edit IDs")?,
                reltype: clone_string_bounded(addition.reltype, "OPC relationship edit types")?,
                target: clone_string_bounded(addition.target, "OPC relationship edit targets")?,
                mode: addition.mode,
            });
        }
        let scan = scan_relationship_source(self.bytes(), &selected, limits)?;
        let removed_count = scan.removal_ranges.len();
        let relationship_count = scan
            .relationship_count
            .checked_sub(removed_count)
            .and_then(|count| count.checked_add(owned_additions.len()))
            .ok_or_else(|| invalid("relationship count overflows"))?;
        if relationship_count > maximum_relationships {
            return Err(invalid("relationship count exceeds the package limit"));
        }
        let fragment_len =
            owned_relationship_fragment_len(&scan.relationship_name, &owned_additions)?;
        let expansion = usize::from(scan.root_empty && !owned_additions.is_empty())
            .checked_mul(
                scan.root_name
                    .len()
                    .checked_add(2)
                    .ok_or_else(|| invalid("relationship root expansion overflows"))?,
            )
            .ok_or_else(|| invalid("relationship root expansion overflows"))?;
        let final_len = self
            .bytes()
            .len()
            .checked_sub(scan.removed_bytes)
            .and_then(|value| value.checked_add(fragment_len))
            .and_then(|value| value.checked_add(expansion))
            .ok_or_else(|| invalid("relationship XML output size overflows"))?;
        let final_events = scan
            .event_count
            .checked_sub(scan.removed_events)
            .and_then(|value| value.checked_add(owned_additions.len()))
            .and_then(|value| {
                value.checked_add(usize::from(scan.root_empty && !owned_additions.is_empty()))
            })
            .ok_or_else(|| invalid("relationship XML event count overflows"))?;
        limits.check(
            ReadResource::RelationshipXmlBytes,
            final_len as u64,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            ReadResource::PartBytes,
            final_len as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            ReadResource::ArchiveEntryBytes,
            final_len as u64,
            limits.max_archive_entry_bytes(),
        )?;
        limits.check(
            ReadResource::XmlEvents,
            final_events as u64,
            limits.max_xml_events() as u64,
        )?;
        let member_present = self.member_present || !owned_additions.is_empty();
        Ok(RelationshipsEditPlan {
            source: self,
            additions: owned_additions,
            scan,
            final_len,
            relationship_count,
            event_count: final_events,
            member_present,
        })
    }
}

fn clone_string_bounded(value: &str, resource: &'static str) -> Result<String> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation { resource, source })?;
    owned.push_str(value);
    Ok(owned)
}

fn owned_relationship_fragment_len(
    relationship_name: &[u8],
    additions: &[OwnedRelationshipEdit],
) -> Result<usize> {
    let mut length = 0usize;
    for addition in additions {
        append_relationship_chunks(
            relationship_name,
            &[(
                addition.reltype.as_str(),
                addition.target.as_str(),
                addition.id.as_str(),
                addition.mode,
            )],
            |chunk| {
                length = length
                    .checked_add(chunk.len())
                    .ok_or_else(|| invalid("relationship fragment size overflows"))?;
                Ok(())
            },
        )?;
    }
    Ok(length)
}

fn append_owned_relationship_chunks(
    output: &mut Vec<u8>,
    relationship_name: &[u8],
    additions: &[OwnedRelationshipEdit],
) -> Result<()> {
    for addition in additions {
        append_relationship_chunks(
            relationship_name,
            &[(
                addition.reltype.as_str(),
                addition.target.as_str(),
                addition.id.as_str(),
                addition.mode,
            )],
            |chunk| {
                output.extend_from_slice(chunk);
                Ok(())
            },
        )?;
    }
    Ok(())
}

fn scan_relationship_source(
    source: &[u8],
    selected: &HashSet<&str>,
    limits: ReadLimits,
) -> Result<RelationshipSourceScan> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let origin = ReaderOrigin::of(source);
    let mut depth = 0usize;
    let mut root_name = None;
    let mut root_tag = None;
    let mut root_close = None;
    let mut root_empty = false;
    let mut removal_ranges = Vec::new();
    removal_ranges
        .try_reserve(selected.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC relationship edit removal ranges",
            source,
        })?;
    let mut removed_bytes = 0usize;
    let mut removed_events = 0usize;
    let mut relationship_count = 0usize;
    let mut event_count = 0usize;
    let mut open_selected: Option<(usize, usize, usize)> = None;
    loop {
        event_count = event_count
            .checked_add(1)
            .ok_or_else(|| invalid("relationship XML event count overflows"))?;
        limits.check(
            ReadResource::XmlEvents,
            event_count as u64,
            limits.max_xml_events() as u64,
        )?;
        let start = relationship_offset(&reader, origin)?;
        let event = reader.read_event()?;
        let end = relationship_offset(&reader, origin)?;
        if end < start || end > source.len() {
            return Err(invalid("relationship XML event range is invalid"));
        }
        match event {
            Event::Start(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Relationships" {
                    return Err(invalid("relationships root must be Relationships"));
                }
                root_name = Some(element.name().as_ref().to_vec());
                root_tag = Some(start..end);
                depth = 1;
            },
            Event::Empty(element) if depth == 0 => {
                if element.local_name().as_ref() != b"Relationships" {
                    return Err(invalid("relationships root must be Relationships"));
                }
                root_name = Some(element.name().as_ref().to_vec());
                root_tag = Some(start..end);
                root_empty = true;
            },
            Event::Start(element) if depth == 1 => {
                if element.local_name().as_ref() == b"Relationship" {
                    relationship_count = relationship_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship count overflows"))?;
                    let id = relationship_id(&element, reader.decoder())?;
                    if selected.contains(id.as_str()) {
                        open_selected = Some((start, end, event_count));
                    }
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("relationship XML depth overflows"))?;
            },
            Event::Empty(element) if depth == 1 => {
                if element.local_name().as_ref() == b"Relationship" {
                    relationship_count = relationship_count
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship count overflows"))?;
                    let id = relationship_id(&element, reader.decoder())?;
                    if selected.contains(id.as_str()) {
                        let range = start..end;
                        removed_bytes = removed_bytes
                            .checked_add(range.len())
                            .ok_or_else(|| invalid("relationship removal bytes overflow"))?;
                        removed_events = removed_events
                            .checked_add(1)
                            .ok_or_else(|| invalid("relationship removal events overflow"))?;
                        removal_ranges.push(range);
                    }
                }
            },
            Event::End(_) => {
                if depth == 1 {
                    root_close = Some(start);
                }
                if depth == 2 {
                    if let Some((open_start, _open_end, open_event_count)) = open_selected.take() {
                        let range = open_start..end;
                        removed_bytes = removed_bytes
                            .checked_add(range.len())
                            .ok_or_else(|| invalid("relationship removal bytes overflow"))?;
                        removed_events = removed_events
                            .checked_add(event_count - open_event_count + 1)
                            .ok_or_else(|| invalid("relationship removal events overflow"))?;
                        removal_ranges.push(range);
                    }
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced relationship XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || root_name.is_none() || root_tag.is_none() {
        return Err(invalid("relationships root is missing or unclosed"));
    }
    removal_ranges.sort_unstable_by_key(|range| range.start);
    if removal_ranges
        .windows(2)
        .any(|pair| pair[0].end > pair[1].start)
    {
        return Err(invalid("relationship removal ranges overlap"));
    }
    let root_name = root_name.expect("checked above");
    let mut relationship_name = root_name
        .iter()
        .position(|byte| *byte == b':')
        .map_or_else(Vec::new, |colon| root_name[..=colon].to_vec());
    relationship_name.extend_from_slice(b"Relationship");
    Ok(RelationshipSourceScan {
        root_name,
        relationship_name,
        root_tag: root_tag.expect("checked above"),
        root_close,
        root_empty,
        removal_ranges,
        removed_bytes,
        removed_events,
        relationship_count,
        event_count,
    })
}

fn relationship_id(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<String> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if attribute.key.as_ref() == b"Id" {
            return attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map(|value| value.into_owned())
                .map_err(|error| invalid(error.to_string()));
        }
    }
    Err(invalid("relationship is missing Id"))
}

impl OwnedRelationships {
    /// Source part URI, or `/` for package relationships.
    #[must_use]
    pub fn owner(&self) -> &PackURI {
        &self.owner
    }

    /// Exact relationship XML; an absent part member has a canonical empty view.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.xml.bytes()
    }

    pub(crate) fn bytes_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml.bytes)
    }

    /// Whether publication includes this relationship member.
    #[must_use]
    pub fn member_present(&self) -> bool {
        self.member_present
    }

    /// Remove one edge while preserving surrounding XML byte for byte.
    ///
    /// A missing ID returns an unchanged token. The result must fit the caller
    /// limit and the default OPC relationship XML limit. Removal keeps an
    /// existing member, even when the collection becomes empty. This edits only
    /// the edge; the format owner must validate and edit its dependency closure.
    pub fn without_relationship(&self, id: &str, max_output_bytes: usize) -> Result<Self> {
        self.without_relationships(&[id], max_output_bytes)
    }

    /// Remove a bounded batch of edges by collecting matches in one source
    /// scan and applying one batched XML update while preserving every
    /// unrelated source byte. The generic [`OwnedXmlPart`] update performs
    /// its own structural validation scan before the single output
    /// allocation.
    ///
    /// Missing IDs and duplicate selectors are idempotent. The borrowed
    /// selectors are deduplicated before source ranges are collected, and the
    /// default 100,000 relationship-per-part ceiling bounds the unique
    /// selector index; the generic XML batch edit ceiling of 65,536 matched
    /// removals applies in addition. The result must fit both
    /// `max_output_bytes` and the default OPC relationship XML ceiling.
    /// Removal keeps an existing member, including when the collection becomes
    /// empty.
    pub fn without_relationships(&self, ids: &[&str], max_output_bytes: usize) -> Result<Self> {
        let maximum = max_output_bytes.min(ReadLimits::default().max_relationship_xml_bytes());
        if ids.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }

        let maximum_selectors = ReadLimits::default().max_relationships_per_part();
        let mut selected_ids: HashSet<&str> = HashSet::new();
        selected_ids
            .try_reserve(ids.len().min(maximum_selectors))
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship removal IDs",
                source,
            })?;
        for id in ids {
            if !selected_ids.contains(id) && selected_ids.len() >= maximum_selectors {
                return Err(invalid(
                    "relationship removal batch exceeds the package limit",
                ));
            }
            selected_ids.insert(*id);
        }

        let mut updates = Vec::new();
        updates
            .try_reserve_exact(selected_ids.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship removal updates",
                source,
            })?;
        let mut reader = NsReader::from_reader(self.bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let origin = ReaderOrigin::of(self.bytes());
        let mut depth = 0usize;
        loop {
            let start = relationship_offset(&reader, origin)?;
            let event = reader.read_event()?;
            let end = relationship_offset(&reader, origin)?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    if depth == 1 && element.local_name().as_ref() == b"Relationship" {
                        let mut remove = false;
                        for attribute in element.checked_attributes() {
                            let attribute =
                                attribute.map_err(|error| invalid(error.to_string()))?;
                            if attribute.key.as_ref() == b"Id" {
                                let value = attribute.decoded_and_normalized_value(
                                    quick_xml::XmlVersion::Implicit1_0,
                                    reader.decoder(),
                                )?;
                                if selected_ids.contains(value.as_ref()) {
                                    remove = true;
                                    break;
                                }
                            }
                        }
                        if remove {
                            updates.push(OwnedElementUpdate {
                                start_tag: start..end,
                                edit: OwnedElementEdit::Remove,
                            });
                        }
                    }
                    if !empty {
                        depth += 1;
                    }
                },
                Event::End(_) => {
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unbalanced relationship XML"))?;
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if updates.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }
        let xml = self.xml.update_elements(&updates, maximum)?;
        crate::pkgreader::PackageReader::parse_owned_relationships(xml.bytes(), &self.owner)?;
        Ok(Self {
            owner: self.owner.clone(),
            xml,
            member_present: self.member_present,
        })
    }

    /// Append one validated internal or external relationship while retaining
    /// every unrelated source byte in the owner `.rels` member.
    pub fn with_relationship(
        &self,
        reltype: &str,
        target: &str,
        id: &str,
        mode: TargetMode,
        max_output_bytes: usize,
    ) -> Result<Self> {
        self.with_relationships(&[(reltype, target, id, mode)], max_output_bytes)
    }

    /// Append a bounded batch in one source scan and one output allocation.
    /// This avoids repeatedly copying a growing owner `.rels` member when one
    /// host operation creates several fresh Ink/image targets.
    pub fn with_relationships(
        &self,
        additions: &[(&str, &str, &str, TargetMode)],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let maximum = max_output_bytes.min(ReadLimits::default().max_relationship_xml_bytes());
        if additions.is_empty() {
            if self.bytes().len() > maximum {
                return Err(invalid("relationship XML exceeds output limit"));
            }
            return Ok(self.clone());
        }
        let parsed =
            crate::pkgreader::PackageReader::parse_owned_relationships(self.bytes(), &self.owner)?;
        let maximum_relationships = ReadLimits::default().max_relationships_per_part();
        let total_relationships = parsed
            .len()
            .checked_add(additions.len())
            .ok_or_else(|| invalid("relationship count overflows"))?;
        if total_relationships > maximum_relationships {
            return Err(invalid("relationship count exceeds the package limit"));
        }
        let mut ids = HashSet::new();
        ids.try_reserve(additions.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship batch IDs",
                source,
            })?;
        for (reltype, target, id, _) in additions {
            if id.is_empty() || reltype.is_empty() || target.is_empty() {
                return Err(invalid("relationship fields must not be empty"));
            }
            if parsed.get(id).is_some() || !ids.insert(id) {
                return Err(invalid("relationship ID already exists"));
            }
        }
        let mut reader = NsReader::from_reader(self.bytes());
        reader.config_mut().trim_text(false);
        reader.config_mut().check_end_names = true;
        let origin = ReaderOrigin::of(self.bytes());
        let mut depth = 0usize;
        let mut root_tag = None;
        let mut root_name = None;
        let mut root_close = None;
        let mut root_empty = false;
        loop {
            let start = relationship_offset(&reader, origin)?;
            let event = reader.read_event()?;
            let end = relationship_offset(&reader, origin)?;
            match event {
                Event::Start(element) if depth == 0 => {
                    if element.local_name().as_ref() != b"Relationships" {
                        return Err(invalid("relationships root must be Relationships"));
                    }
                    root_tag = Some(start..end);
                    root_name = Some(element.name().as_ref().to_vec());
                    depth = 1;
                },
                Event::Empty(element) if depth == 0 => {
                    if element.local_name().as_ref() != b"Relationships" {
                        return Err(invalid("relationships root must be Relationships"));
                    }
                    root_tag = Some(start..end);
                    root_name = Some(element.name().as_ref().to_vec());
                    root_empty = true;
                    break;
                },
                Event::Start(_) => {
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship XML depth overflows"))?;
                },
                Event::End(_) => {
                    if depth == 1 {
                        root_close = Some(start..end);
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unbalanced relationship XML"))?;
                    if depth == 0 {
                        break;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        let root_tag = root_tag.ok_or_else(|| invalid("relationships root is missing"))?;
        if !root_empty && root_close.is_none() {
            return Err(invalid("relationships root is unclosed"));
        }
        let root_name = root_name.ok_or_else(|| invalid("relationships root name is missing"))?;
        let relationship_name = root_name.iter().position(|byte| *byte == b':').map_or_else(
            || b"Relationship".to_vec(),
            |colon| {
                let mut name = root_name[..colon].to_vec();
                name.extend_from_slice(b":Relationship");
                name
            },
        );
        let mut fragment_len = 0usize;
        append_relationship_chunks(&relationship_name, additions, |chunk| {
            fragment_len = fragment_len
                .checked_add(chunk.len())
                .ok_or_else(|| invalid("relationship batch XML size overflows"))?;
            Ok(())
        })?;
        let expansion = if root_empty {
            root_name
                .len()
                .checked_add(2)
                .ok_or_else(|| invalid("relationship root expansion overflows"))?
        } else {
            0
        };
        let size = self
            .bytes()
            .len()
            .checked_add(fragment_len)
            .and_then(|size| size.checked_add(expansion))
            .ok_or_else(|| invalid("relationship XML output size overflows"))?;
        if size > maximum {
            return Err(invalid("relationship XML exceeds output limit"));
        }
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(size)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship batch output",
                source,
            })?;
        let insertion = if root_empty {
            root_tag
                .end
                .checked_sub(2)
                .ok_or_else(|| invalid("empty relationship root is incomplete"))?
        } else {
            root_close
                .as_ref()
                .ok_or_else(|| invalid("relationships root is unclosed"))?
                .start
        };
        bytes.extend_from_slice(&self.bytes()[..insertion]);
        if root_empty {
            bytes.push(b'>');
        }
        append_relationship_chunks(&relationship_name, additions, |chunk| {
            bytes.extend_from_slice(chunk);
            Ok(())
        })?;
        if root_empty {
            bytes.extend_from_slice(b"</");
            bytes.extend_from_slice(&root_name);
            bytes.push(b'>');
            bytes.extend_from_slice(&self.bytes()[root_tag.end..]);
        } else {
            bytes.extend_from_slice(&self.bytes()[insertion..]);
        }
        debug_assert_eq!(bytes.len(), size);
        let xml = OwnedXmlPart::capture(
            self.xml.name.clone(),
            self.xml.content_type.clone(),
            Arc::new(bytes),
        )?;
        crate::pkgreader::PackageReader::parse_owned_relationships(xml.bytes(), &self.owner)?;
        Ok(Self {
            owner: self.owner.clone(),
            xml,
            member_present: true,
        })
    }
}

impl OpcPackage {
    /// Plan source-bound relationship capture without retaining XML bytes.
    ///
    /// The returned metrics describe the owner's current relationship
    /// collection and its currently bound source XML, when the semantic
    /// binding still matches.  A changed collection falls back to the
    /// output-free canonical serializer plan.  The plan is admission data,
    /// not authority to skip a later capture: callers must precharge the
    /// operation aggregate and then invoke
    /// [`OpcPackage::source_relationships_with_limits`] again so the current
    /// collection is freshly checked and materialized.
    pub fn plan_source_relationships_with_limits(
        &self,
        owner: &PackURI,
        limits: ReadLimits,
    ) -> Result<RelationshipSourcePlan> {
        let (owner_ref, relationships) = if owner.as_str() == "/" {
            (owner, self.rels())
        } else {
            let part = self.get_part(owner)?;
            (part.partname(), part.rels())
        };
        OwnedXmlPart::check_derived_capture_member_name(owner_ref, limits)?;
        let relationship_uri = owner_ref.rels_uri().map_err(OpcError::InvalidPackUri)?;
        check_relationship_capture_limits(relationships, limits)?;

        let source = self
            .source_relationships_xml
            .get(owner_ref)
            .filter(|source| source.binding.matches(relationships));
        let member_present = owner_ref.as_str() == "/"
            || !relationships.is_empty()
            || self.source_relationships_member_present(owner_ref);
        let (final_len, event_count, relationship_count, member_present) =
            if let Some(source) = source {
                limits.check(
                    ReadResource::RelationshipXmlBytes,
                    source.bytes.len() as u64,
                    limits.max_relationship_xml_bytes() as u64,
                )?;
                limits.check(
                    ReadResource::TotalRelationshipXmlBytes,
                    source.bytes.len() as u64,
                    limits.max_total_relationship_xml_bytes() as u64,
                )?;
                limits.check(
                    ReadResource::PartBytes,
                    source.bytes.len() as u64,
                    limits.max_part_bytes(),
                )?;
                OwnedXmlPart::preflight_capture_with_limits(
                    &relationship_uri,
                    crate::constants::content_type::OPC_RELATIONSHIPS,
                    source.bytes.as_slice(),
                    limits,
                )?;
                let (event_count, parsed_relationships) =
                    relationship_xml_metrics(source.bytes.as_slice(), limits)?;
                if parsed_relationships != relationships.len() {
                    return Err(invalid(
                        "relationship source plan disagrees with the in-memory graph",
                    ));
                }
                (
                    source.bytes.len(),
                    event_count,
                    relationships.len(),
                    member_present,
                )
            } else {
                let canonical = relationships.plan_canonical(limits)?;
                OwnedXmlPart::check_capture_size(&relationship_uri, canonical.final_len(), limits)?;
                (
                    canonical.final_len(),
                    canonical.event_count(),
                    canonical.relationship_count(),
                    member_present,
                )
            };

        // `preflight_capture_with_limits` above admits the raw source-publication
        // profile (including trim_text(false) events, depth, and all source
        // attributes). The metrics below intentionally use trim_text(true),
        // matching relationship ingress and the aggregate relationship ledger.
        // Keep both views: the first protects fresh OwnedXmlPart capture and
        // this one supplies exact retained relationship aggregate accounting.
        limits.check(
            ReadResource::XmlEvents,
            event_count as u64,
            limits.max_xml_events() as u64,
        )?;
        limits.check(
            ReadResource::TotalRelationshipXmlBytes,
            final_len as u64,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        limits.check(
            ReadResource::TotalRelationshipXmlEvents,
            event_count as u64,
            limits.max_total_relationship_xml_events() as u64,
        )?;
        Ok(RelationshipSourcePlan {
            final_len,
            event_count,
            relationship_count,
            member_present,
        })
    }

    /// Capture source-bound relationship XML without exposing mutable provenance.
    /// Newly authored or changed collections receive a validated canonical view.
    pub fn source_relationships(&self, owner: &PackURI) -> Result<OwnedRelationships> {
        self.source_relationships_with_limits(owner, ReadLimits::default())
    }

    /// Capture source-bound relationship XML under an explicit bounded read
    /// policy. Relationship counts and field lengths are checked before
    /// semantic binding clones; retained XML and canonical output are checked
    /// before parsing or allocation.
    pub fn source_relationships_with_limits(
        &self,
        owner: &PackURI,
        limits: ReadLimits,
    ) -> Result<OwnedRelationships> {
        let (owner_ref, relationships) = if owner.as_str() == "/" {
            (owner, self.rels())
        } else {
            let part = self.get_part(owner)?;
            (part.partname(), part.rels())
        };
        OwnedXmlPart::check_derived_capture_member_name(owner_ref, limits)?;
        let relationship_uri = owner_ref.rels_uri().map_err(OpcError::InvalidPackUri)?;
        check_relationship_capture_limits(relationships, limits)?;
        let source = self
            .source_relationships_xml
            .get(owner_ref)
            .filter(|source| source.binding.matches(relationships));
        let bytes = if let Some(source) = source {
            limits.check(
                ReadResource::RelationshipXmlBytes,
                source.bytes.len() as u64,
                limits.max_relationship_xml_bytes() as u64,
            )?;
            limits.check(
                ReadResource::TotalRelationshipXmlBytes,
                source.bytes.len() as u64,
                limits.max_total_relationship_xml_bytes() as u64,
            )?;
            limits.check(
                ReadResource::PartBytes,
                source.bytes.len() as u64,
                limits.max_part_bytes(),
            )?;
            OwnedXmlPart::check_capture_size(&relationship_uri, source.bytes.len(), limits)?;
            Arc::clone(&source.bytes)
        } else {
            let canonical = relationships.plan_canonical(limits)?;
            let length = canonical.final_len();
            limits.check(
                ReadResource::TotalRelationshipXmlBytes,
                length as u64,
                limits.max_total_relationship_xml_bytes() as u64,
            )?;
            OwnedXmlPart::check_capture_size(&relationship_uri, length, limits)?;
            Arc::new(canonical.materialize(limits)?)
        };
        crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            &bytes, owner_ref, limits,
        )?;
        let member_present = owner_ref.as_str() == "/"
            || !relationships.is_empty()
            || self.source_relationships_member_present(owner_ref);
        let owner = owner_ref.clone();
        let xml = OwnedXmlPart::capture_with_limits(
            relationship_uri,
            crate::constants::content_type::OPC_RELATIONSHIPS.into(),
            bytes,
            limits,
        )?;
        Ok(OwnedRelationships {
            owner,
            xml,
            member_present,
        })
    }

    /// Replace an exact expected relationship snapshot with a source-derived token.
    ///
    /// Owner, XML bytes and member presence must still match. Stale snapshots,
    /// different owners and signed changes fail before mutation. The caller is
    /// responsible for the format-level dependency closure of changed edges.
    pub fn try_replace_relationships(
        &mut self,
        expected: &OwnedRelationships,
        replacement: &OwnedRelationships,
    ) -> Result<bool> {
        self.try_replace_relationships_with_limits(expected, replacement, ReadLimits::default())
    }

    /// Replace an exact expected relationship snapshot under an explicit
    /// bounded read policy.
    ///
    /// The current owner XML is checked before replacement parsing or package
    /// mutation. A changed signed package is refused before retaining parsed
    /// replacement fields, and the replacement is parsed under the same
    /// policy used for the current source check.
    pub fn try_replace_relationships_with_limits(
        &mut self,
        expected: &OwnedRelationships,
        replacement: &OwnedRelationships,
        limits: ReadLimits,
    ) -> Result<bool> {
        if expected.owner != replacement.owner {
            return Err(invalid("relationship replacement has a different owner"));
        }
        let current = self.source_relationships_with_limits(&expected.owner, limits)?;
        if current != *expected {
            return Err(invalid("stale relationship replacement"));
        }
        if expected == replacement {
            return Ok(false);
        }
        if self.is_signed() || self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let relationships = crate::pkgreader::PackageReader::parse_owned_relationships_with_limits(
            replacement.bytes(),
            &replacement.owner,
            limits,
        )?;
        if !replacement.member_present && !relationships.is_empty() {
            return Err(invalid("absent relationship member has edges"));
        }
        if replacement.member_present {
            // The token becomes the owner's retained `.rels` source, which the
            // writer publishes without its audit (change 0665). Its bytes came
            // from the caller, so audit them before any mutation (ADR 0006).
            let member = replacement
                .owner
                .rels_uri()
                .map_err(OpcError::InvalidPackUri)?;
            super::verified_metadata_token(
                &member,
                Arc::clone(&replacement.xml.bytes),
                limits.max_relationship_xml_bytes(),
            )?;
        }
        let binding = RelationshipBinding::from_relationships(&relationships)?;
        self.source_relationships_xml
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "owned relationship provenance",
                source,
            })?;
        let source = Arc::new(SourceRelationshipsXml {
            bytes: Arc::clone(&replacement.xml.bytes),
            binding,
        });
        if replacement.owner.as_str() == "/" {
            self.revoke_exact_source();
            self.signature_graph_tracked = true;
            self.rels = relationships;
        } else {
            *self.get_part_mut(&replacement.owner)?.rels_mut() = relationships;
        }
        if replacement.member_present {
            self.source_relationships_xml
                .insert(replacement.owner.clone(), source);
        } else {
            self.source_relationships_xml.remove(&replacement.owner);
        }
        Ok(true)
    }
}

/// Count and emit exactly the same lexical bytes without temporary field strings.
fn append_relationship_chunks(
    name: &[u8],
    additions: &[(&str, &str, &str, TargetMode)],
    mut emit: impl FnMut(&[u8]) -> Result<()>,
) -> Result<()> {
    for (reltype, target, id, mode) in additions {
        emit(b"<")?;
        emit(name)?;
        for (prefix, value) in [
            (b" Id=\"".as_slice(), *id),
            (b" Type=\"", *reltype),
            (b" Target=\"", *target),
        ] {
            emit(prefix)?;
            let mut start = 0;
            for (offset, byte) in value.bytes().enumerate() {
                let escaped: &[u8] = match byte {
                    b'&' => b"&amp;",
                    b'<' => b"&lt;",
                    b'>' => b"&gt;",
                    b'"' => b"&quot;",
                    b'\'' => b"&apos;",
                    _ => continue,
                };
                emit(&value.as_bytes()[start..offset])?;
                emit(escaped)?;
                start = offset + 1;
            }
            emit(&value.as_bytes()[start..])?;
            emit(b"\"")?;
        }
        if *mode == TargetMode::External {
            emit(b" TargetMode=\"External\"")?;
        }
        emit(b"/>")?;
    }
    Ok(())
}

/// Count retained relationship XML events without allocating a parsed graph.
///
/// This intentionally uses the same whitespace and end-tag configuration as
/// OPC relationship ingress.  The source bytes have already passed semantic
/// ingress validation; the typed element count is retained as a consistency
/// check against the current relationship collection.
fn relationship_xml_metrics(bytes: &[u8], limits: ReadLimits) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    reader.config_mut().check_end_names = true;
    let mut events = 0usize;
    let mut relationships = 0usize;
    let mut depth = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("relationship XML event count overflows"))?;
        limits.check(
            ReadResource::XmlEvents,
            events as u64,
            limits.max_xml_events() as u64,
        )?;
        let decoder = reader.decoder();
        match reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
        {
            Event::Start(element) => {
                let next_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("relationship XML depth overflows"))?;
                limits.check(
                    ReadResource::XmlDepth,
                    next_depth as u64,
                    limits.max_xml_depth() as u64,
                )?;
                if depth == 1 && element.local_name().as_ref() == b"Relationship" {
                    check_relationship_attributes(&element, decoder, limits)?;
                    relationships = relationships
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship count overflows"))?;
                }
                depth = next_depth;
            },
            Event::Empty(element) => {
                if depth == 1 && element.local_name().as_ref() == b"Relationship" {
                    check_relationship_attributes(&element, decoder, limits)?;
                    relationships = relationships
                        .checked_add(1)
                        .ok_or_else(|| invalid("relationship count overflows"))?;
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("relationship XML depth underflows"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok((events, relationships))
}

/// Recheck the two attribute byte views enforced by relationship ingress.
/// The raw value is charged before decoding so a hostile entity spelling
/// cannot force an unbounded temporary before the caller's limit is applied.
fn check_relationship_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
    limits: ReadLimits,
) -> Result<()> {
    for attribute_result in element.checked_attributes() {
        let attribute = attribute_result.map_err(|error| invalid(error.to_string()))?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            attribute.value.as_ref().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            value.len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OpcError {
    OpcError::InvalidRelationship(message.into())
}

/// The byte offset in the relationship bytes of `reader`'s position.
///
/// `origin` is the bytes' [`ReaderOrigin`]: the reader does not count a
/// leading byte-order mark, and spans taken here splice the original bytes.
fn relationship_offset(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("relationship XML position overflows usize"))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure"
    )]
    use super::*;
    use crate::{BlobPart, PackageWriter};

    const XML: &[u8] = br#"<?xml version='1.0'?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- before --> <r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>
 <!-- middle --> <r:Relationship Id='rId2' Type='urn:test' Target='https://example.test/b' TargetMode='External'/>
 <!-- after -->
</r:Relationships>"#;

    fn source(xml: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        for (name, bytes) in [
            ("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#.as_slice()),
            ("_rels/.rels", xml),
            ("custom/item.bin", b"opaque".as_slice()),
            ("custom/_rels/item.bin.rels", xml),
        ] { writer.write_deflated(name, bytes).unwrap(); }
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn canonical_relationship_plan_defers_source_less_xml_and_honors_exact_caps() {
        let mut relationships = Relationships::new("/xl/drawings".to_owned());
        relationships
            .try_add_relationship(
                "urn:test&kind".to_owned(),
                "../media/vector&1.svg".to_owned(),
                "rId7".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let plan = relationships.plan_canonical(ReadLimits::default()).unwrap();
        assert_eq!(plan.relationship_count(), 1);
        assert_eq!(plan.event_count(), 5);
        let exact = plan.final_len();
        let limits = ReadLimits::builder()
            .max_relationship_xml_bytes(exact)
            .unwrap()
            .max_total_relationship_xml_bytes(exact)
            .unwrap()
            .max_part_bytes(exact as u64)
            .unwrap()
            .max_archive_entry_bytes(exact as u64)
            .unwrap()
            .max_xml_events(5)
            .unwrap()
            .max_total_relationship_xml_events(5)
            .unwrap()
            .build()
            .unwrap();
        let output = relationships
            .plan_canonical(limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        assert_eq!(output, relationships.to_xml().as_bytes());

        let under = ReadLimits::builder()
            .max_relationship_xml_bytes(exact - 1)
            .unwrap()
            .max_part_bytes((exact - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(relationships.plan_canonical(under).is_err());
        assert!(matches!(
            relationships
                .plan_canonical(ReadLimits::default())
                .unwrap()
                .materialize(under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                ..
            })
        ));
    }

    #[test]
    fn source_relationship_plan_reports_retained_comments_and_exact_caps() {
        let package = OpcPackage::from_vec(source(XML)).unwrap();
        let owner = PackURI::new("/").unwrap();
        let (events, relationships) = relationship_xml_metrics(XML, ReadLimits::default()).unwrap();
        assert_eq!(events, 10);
        assert_eq!(relationships, 2);
        let plan = package
            .plan_source_relationships_with_limits(&owner, ReadLimits::default())
            .unwrap();
        let retained_token = package
            .source_relationships_with_limits(&owner, ReadLimits::default())
            .unwrap();
        assert_eq!(plan.final_len(), XML.len());
        assert_eq!(retained_token.bytes(), XML);
        assert_eq!(plan.final_len(), retained_token.bytes().len());
        assert_eq!(plan.event_count(), events);
        assert_eq!(plan.relationship_count(), relationships);
        assert!(plan.member_present());
        assert!(
            std::str::from_utf8(XML)
                .unwrap()
                .contains("<!-- before -->")
        );
        assert!(
            std::str::from_utf8(XML)
                .unwrap()
                .contains("<!-- middle -->")
        );

        let exact = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_total_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_part_bytes(XML.len() as u64)
            .unwrap()
            // The source-publication validator retains whitespace text events
            // (17 total); relationship ingress trims them (10 aggregate
            // events). Both ceilings must admit this retained member.
            .max_xml_events(17)
            .unwrap()
            .max_total_relationship_xml_events(events)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            package
                .plan_source_relationships_with_limits(&owner, exact)
                .unwrap(),
            plan
        );
        assert_eq!(
            package
                .source_relationships_with_limits(&owner, exact)
                .unwrap()
                .bytes(),
            XML
        );

        let bytes_under = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&owner, bytes_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                ..
            })
        ));
        let raw_events_under = ReadLimits::builder()
            .max_xml_events(16)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&owner, raw_events_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                ..
            })
        ));
        let aggregate_bytes_under = ReadLimits::builder()
            .max_total_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&owner, aggregate_bytes_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                ..
            })
        ));
        let aggregate_events_under = ReadLimits::builder()
            .max_total_relationship_xml_events(events - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&owner, aggregate_events_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlEvents,
                ..
            })
        ));

        let depth_under = ReadLimits::builder()
            .max_xml_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&owner, depth_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                ..
            })
        ));
        assert!(matches!(
            package.source_relationships_with_limits(&owner, depth_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                ..
            })
        ));

        let escaped_target = "&#x61;".repeat(20);
        let escaped_attribute_xml = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="{escaped_target}" TargetMode="External"/></Relationships>"#
        );
        let escaped_attribute_package =
            OpcPackage::from_vec(source(escaped_attribute_xml.as_bytes())).unwrap();
        let raw_attribute_under = ReadLimits::builder()
            .max_xml_attribute_bytes(96)
            .unwrap()
            .max_relationship_target_bytes(96)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            escaped_attribute_package
                .plan_source_relationships_with_limits(&owner, raw_attribute_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlAttributeBytes,
                ..
            })
        ));
        assert!(matches!(
            escaped_attribute_package.source_relationships_with_limits(&owner, raw_attribute_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlAttributeBytes,
                ..
            })
        ));
    }

    #[test]
    fn source_relationship_plan_tracks_changed_new_empty_and_mixed_case_owners() {
        let retained = OpcPackage::from_vec(source(XML)).unwrap();
        let mixed_case = PackURI::new("/CUSTOM/ITEM.BIN").unwrap();
        let retained_plan = retained
            .plan_source_relationships_with_limits(&mixed_case, ReadLimits::default())
            .unwrap();
        assert!(retained_plan.member_present());
        assert_eq!(retained_plan.relationship_count(), 2);

        let mut changed = OpcPackage::from_vec(source(XML)).unwrap();
        changed
            .rels_mut()
            .try_add_relationship(
                "urn:changed".to_owned(),
                "custom/item.bin".to_owned(),
                "rId9".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        let root = PackURI::new("/").unwrap();
        let changed_plan = changed
            .plan_source_relationships_with_limits(&root, ReadLimits::default())
            .unwrap();
        let changed_token = changed
            .source_relationships_with_limits(&root, ReadLimits::default())
            .unwrap();
        assert!(changed_plan.member_present());
        assert_eq!(changed_plan.relationship_count(), 3);
        assert_eq!(changed_plan.final_len(), changed_token.bytes().len());
        assert_eq!(changed_plan.event_count(), 7);
        let changed_exact = ReadLimits::builder()
            .max_relationship_xml_bytes(changed_plan.final_len())
            .unwrap()
            .max_total_relationship_xml_bytes(changed_plan.final_len())
            .unwrap()
            .max_part_bytes(changed_plan.final_len() as u64)
            .unwrap()
            .max_xml_events(changed_plan.event_count())
            .unwrap()
            .max_total_relationship_xml_events(changed_plan.event_count())
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            changed
                .plan_source_relationships_with_limits(&root, changed_exact)
                .unwrap(),
            changed_plan
        );
        let changed_under = ReadLimits::builder()
            .max_relationship_xml_bytes(changed_plan.final_len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            changed.plan_source_relationships_with_limits(&root, changed_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                ..
            })
        ));

        let empty_xml =
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        let explicit_empty = OpcPackage::from_vec(source(empty_xml)).unwrap();
        let explicit_owner = PackURI::new("/custom/item.bin").unwrap();
        let explicit_plan = explicit_empty
            .plan_source_relationships_with_limits(&explicit_owner, ReadLimits::default())
            .unwrap();
        assert!(explicit_plan.member_present());
        assert_eq!(explicit_plan.relationship_count(), 0);

        let mut authored = OpcPackage::new();
        let absent_owner = PackURI::new("/custom/no-rels.bin").unwrap();
        authored.add_part(Box::new(BlobPart::new(
            absent_owner.clone(),
            "application/octet-stream".into(),
            vec![],
        )));
        let absent_plan = authored
            .plan_source_relationships_with_limits(&absent_owner, ReadLimits::default())
            .unwrap();
        let absent_token = authored
            .source_relationships_with_limits(&absent_owner, ReadLimits::default())
            .unwrap();
        assert!(!absent_plan.member_present());
        assert!(!absent_token.member_present());
        assert_eq!(absent_plan.relationship_count(), 0);
        assert_eq!(absent_plan.final_len(), absent_token.bytes().len());

        let authored_root = authored
            .plan_source_relationships_with_limits(&root, ReadLimits::default())
            .unwrap();
        let authored_root_token = authored
            .source_relationships_with_limits(&root, ReadLimits::default())
            .unwrap();
        assert!(authored_root.member_present());
        assert_eq!(authored_root.relationship_count(), 0);
        assert_eq!(authored_root.final_len(), authored_root_token.bytes().len());
    }

    #[test]
    fn relationship_removal_round_trip_preserves_context_and_restores_source_xml() {
        let source = source(XML);
        for owned in [false, true] {
            for owner in ["/", "/custom/item.bin"] {
                let owner = PackURI::new(owner).unwrap();
                let mut package = if owned {
                    OpcPackage::from_vec(source.clone()).unwrap()
                } else {
                    OpcPackage::from_bytes(&source).unwrap()
                };
                let before = package.source_relationships(&owner).unwrap();
                assert_eq!(before.bytes(), XML);
                assert!(Arc::ptr_eq(&before.xml.bytes, &before.clone().xml.bytes));
                let after = before.without_relationship("rId1", XML.len()).unwrap();
                assert_eq!(after.bytes(), String::from_utf8(XML.to_vec()).unwrap().replace("<r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>", "").as_bytes());
                assert!(
                    before
                        .without_relationship("rId1", after.bytes().len() - 1)
                        .is_err()
                );
                assert_eq!(
                    before
                        .without_relationship("rId1", after.bytes().len())
                        .unwrap(),
                    after
                );
                assert_eq!(
                    before.without_relationship("missing", XML.len()).unwrap(),
                    before
                );
                assert!(!package.try_replace_relationships(&before, &before).unwrap());
                assert!(package.try_replace_relationships(&before, &after).unwrap());
                let output = PackageWriter::to_bytes(&package).unwrap();
                let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
                assert_eq!(
                    archive
                        .read(owner.rels_uri().unwrap().membername())
                        .unwrap(),
                    after.bytes()
                );
                assert_eq!(archive.read("custom/item.bin").unwrap(), b"opaque");
                let mut reopened = OpcPackage::from_vec(output).unwrap();
                assert_eq!(reopened.source_relationships(&owner).unwrap(), after);
                let unchanged = PackageWriter::to_bytes(&reopened).unwrap();
                assert!(reopened.try_replace_relationships(&before, &after).is_err());
                assert_eq!(PackageWriter::to_bytes(&reopened).unwrap(), unchanged);
                assert!(reopened.try_replace_relationships(&after, &before).unwrap());
                let restored = PackageWriter::to_bytes(&reopened).unwrap();
                let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
                assert_eq!(
                    archive
                        .read(owner.rels_uri().unwrap().membername())
                        .unwrap(),
                    XML
                );
                assert_eq!(archive.read("custom/item.bin").unwrap(), b"opaque");
            }
        }
    }

    #[test]
    fn relationship_batch_append_copies_source_once_and_keeps_comments() {
        let bytes = source(XML);
        let mut package = OpcPackage::from_vec(bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .with_relationships(
                &[
                    ("urn:ink", "custom/item.bin", "rId7", TargetMode::Internal),
                    (
                        "urn:image",
                        "custom/image.png",
                        "rId8",
                        TargetMode::External,
                    ),
                ],
                1024 * 1024,
            )
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("<!-- before -->")
        );
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("Id=\"rId7\"")
        );
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("Id=\"rId8\"")
        );
        package.try_replace_relationships(&before, &after).unwrap();
        assert_eq!(package.source_relationships(&owner).unwrap(), after);
    }

    #[test]
    fn borrowed_relationship_plan_reports_exact_mixed_output_and_preserves_prefix() {
        let package = OpcPackage::from_bytes(&source(XML)).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let additions = [RelationshipEdit {
            id: "rId7",
            reltype: "urn:test&new",
            target: "custom/new&item.bin",
            mode: TargetMode::Internal,
        }];
        let plan = before
            .plan_edit(&additions, &["rId1"], ReadLimits::default())
            .unwrap();
        assert!(plan.final_len() < before.bytes().len() + 256);
        assert_eq!(plan.relationship_count(), 2);
        let exact = plan.final_len().max(before.bytes().len());
        let limits = ReadLimits::builder()
            .max_relationship_xml_bytes(exact)
            .unwrap()
            .max_part_bytes(exact as u64)
            .unwrap()
            .build()
            .unwrap();
        let after = before
            .plan_edit(&additions, &["rId1"], limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        let text = std::str::from_utf8(after.bytes()).unwrap();
        assert!(text.contains("<!-- before -->"));
        assert!(text.contains("<r:Relationship Id=\"rId7\""));
        assert!(!text.contains("Id='rId1'"));
        assert_eq!(after.member_present(), before.member_present());
    }

    #[test]
    fn borrowed_relationship_plan_expands_prefixed_empty_root_and_honors_exact_cap() {
        let xml = br#"<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'/>"#;
        let package = OpcPackage::from_bytes(&source(xml)).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let additions = [RelationshipEdit {
            id: "rId1",
            reltype: "urn:test",
            target: "item.bin",
            mode: TargetMode::Internal,
        }];
        let plan = before
            .plan_edit(&additions, &[], ReadLimits::default())
            .unwrap();
        let exact = plan.final_len().max(before.bytes().len());
        let limits = ReadLimits::builder()
            .max_relationship_xml_bytes(exact)
            .unwrap()
            .max_part_bytes(exact as u64)
            .unwrap()
            .build()
            .unwrap();
        let after = before
            .plan_edit(&additions, &[], limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .contains("</r:Relationships>")
        );
        assert!(
            before
                .plan_edit(&additions, &[], limits)
                .unwrap()
                .materialize(
                    ReadLimits::builder()
                        .max_relationship_xml_bytes(exact - 1)
                        .unwrap()
                        .max_part_bytes((exact - 1) as u64)
                        .unwrap()
                        .build()
                        .unwrap(),
                )
                .is_err()
        );
    }

    #[test]
    fn relationship_batch_removal_preserves_pi_comments_namespaces_and_paired_tags() {
        let xml = br#"<?xml version='1.0'?>
<?before?>
<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'>
 <!-- before --> <r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='rId1'></r:Relationship>
 <?middle?> <r:Relationship Id='rId2' Type='urn:test' Target='https://example.test/b' TargetMode='External'/>
 <!-- after -->
</r:Relationships>
<?after?>
"#;
        let package = OpcPackage::from_bytes(&source(xml)).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .without_relationships(&["rId1", "missing", "rId2", "rId2"], xml.len())
            .unwrap();

        assert!(after.member_present());
        assert!(
            !after
                .bytes()
                .windows(b"rId1".len())
                .any(|window| window == b"rId1")
        );
        assert!(
            !after
                .bytes()
                .windows(b"rId2".len())
                .any(|window| window == b"rId2")
        );
        for marker in [
            b"<?before?>".as_slice(),
            b"<?middle?>",
            b"<?after?>",
            b"<!-- before -->",
            b"<!-- after -->",
        ] {
            assert!(
                after
                    .bytes()
                    .windows(marker.len())
                    .any(|window| window == marker)
            );
        }
        let sequential = before
            .without_relationship("rId1", xml.len())
            .unwrap()
            .without_relationship("rId2", xml.len())
            .unwrap();
        assert_eq!(after.bytes(), sequential.bytes());
        assert!(
            before
                .without_relationships(&["rId1", "rId2"], after.bytes().len() - 1)
                .is_err()
        );
    }

    #[test]
    fn relationship_batch_removal_missing_duplicate_and_empty_batches_are_bounded_noops() {
        let bytes = source(XML);
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();

        let missing = before
            .without_relationships(&["missing", "missing"], XML.len())
            .unwrap();
        assert_eq!(missing, before);
        assert!(Arc::ptr_eq(&before.xml.bytes, &missing.xml.bytes));
        assert!(
            before
                .without_relationships(&["missing"], before.bytes().len() - 1)
                .is_err()
        );
        assert!(
            before
                .without_relationships(&[], before.bytes().len() - 1)
                .is_err()
        );
        let exact_empty = before
            .without_relationships(&[], before.bytes().len())
            .unwrap();
        assert_eq!(exact_empty, before);
        assert!(Arc::ptr_eq(&before.xml.bytes, &exact_empty.xml.bytes));
    }

    #[test]
    fn relationship_batch_removal_keeps_absent_member_presence_on_noop() {
        let mut package = OpcPackage::new();
        let owner = PackURI::new("/custom/no-rels.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/octet-stream".into(),
            vec![],
        )));
        let before = package.source_relationships(&owner).unwrap();
        assert!(!before.member_present());
        let after = before
            .without_relationships(&["missing"], before.bytes().len())
            .unwrap();
        assert_eq!(after, before);
        assert!(!after.member_present());
        assert!(Arc::ptr_eq(&before.xml.bytes, &after.xml.bytes));
    }

    #[test]
    fn relationship_batch_expands_empty_root_without_dropping_trailing_members() {
        let xml = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/><!--tail--><?keep?>
"#;
        let archive = source(xml);
        let package = OpcPackage::from_bytes(&archive).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .with_relationships(
                &[(
                    crate::constants::relationship_type::OFFICE_DOCUMENT,
                    "word/document.xml",
                    "rId1",
                    TargetMode::Internal,
                )],
                4096,
            )
            .unwrap();
        assert!(
            std::str::from_utf8(after.bytes())
                .unwrap()
                .ends_with("<!--tail--><?keep?>\n")
        );
    }

    #[test]
    fn empty_relationship_batch_enforces_requested_output_limit() {
        let bytes = source(XML);
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/").unwrap();
        let before = package.source_relationships(&owner).unwrap();

        assert!(before.with_relationships(&[], 0).is_err());
        assert!(
            before
                .with_relationships(&[], before.bytes().len() - 1)
                .is_err()
        );
        assert_eq!(
            before
                .with_relationships(&[], before.bytes().len())
                .unwrap(),
            before
        );
    }

    #[test]
    fn relationship_append_accepts_exact_final_size_for_paired_and_empty_roots() {
        let sources: &[&[u8]] = &[
            XML,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/><!--tail-->"#,
            br#"<p:Relationships xmlns:p="http://schemas.openxmlformats.org/package/2006/relationships" /><?tail?>"#,
        ];
        let additions = [
            (
                "urn:image",
                "custom/image.png",
                "rId9",
                TargetMode::Internal,
            ),
            (
                "urn:external",
                "https://example.test/α?a=1&b=2",
                "rId10",
                TargetMode::External,
            ),
        ];
        for xml in sources {
            let package = OpcPackage::from_bytes(&source(xml)).unwrap();
            let before = package
                .source_relationships(&PackURI::new("/").unwrap())
                .unwrap();
            let expected = before.with_relationships(&additions, 4096).unwrap();
            let exact = before
                .with_relationships(&additions, expected.bytes().len())
                .unwrap();
            assert_eq!(exact.bytes(), expected.bytes());
            assert!(
                before
                    .with_relationships(&additions, expected.bytes().len() - 1)
                    .is_err()
            );
            assert_eq!(before.bytes(), *xml);
        }
    }

    #[test]
    fn relationship_chunks_escape_each_field_without_changing_utf8() {
        let additions = [("urn:α&", "a<\"'>", "id", TargetMode::External)];
        let mut bytes = Vec::new();
        append_relationship_chunks(b"p:Relationship", &additions, |chunk| {
            bytes.extend_from_slice(chunk);
            Ok(())
        })
        .unwrap();
        assert_eq!(bytes,
            "<p:Relationship Id=\"id\" Type=\"urn:α&amp;\" Target=\"a&lt;&quot;&apos;&gt;\" TargetMode=\"External\"/>".as_bytes());
    }

    #[test]
    fn relationship_tokens_reject_lexically_stale_sources_and_different_owners() {
        let bytes = source(XML);
        let original = OpcPackage::from_bytes(&bytes).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = original.source_relationships(&owner).unwrap();
        let after = before.without_relationship("rId2", XML.len()).unwrap();
        let changed = String::from_utf8(XML.to_vec())
            .unwrap()
            .replace("<!-- middle -->", "<!-- other -->");
        let changed_bytes = source(changed.as_bytes());
        let mut package = OpcPackage::from_vec(changed_bytes).unwrap();
        let saved = PackageWriter::to_bytes(&package).unwrap();
        assert!(package.try_replace_relationships(&before, &after).is_err());
        let root = package
            .source_relationships(&PackURI::new("/").unwrap())
            .unwrap();
        assert!(package.try_replace_relationships(&before, &root).is_err());
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), saved);
    }

    #[test]
    fn relationship_tokens_preserve_empty_member_presence_and_refuse_signed_changes() {
        let bytes = source(XML);
        let mut package = OpcPackage::from_vec(bytes).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let empty = before
            .without_relationship("rId1", XML.len())
            .unwrap()
            .without_relationship("rId2", XML.len())
            .unwrap();
        assert!(package.try_replace_relationships(&before, &empty).unwrap());
        let output = PackageWriter::to_bytes(&package).unwrap();
        assert_eq!(
            soapberry_zip::office::ArchiveReader::new(&output)
                .unwrap()
                .read(owner.rels_uri().unwrap().membername())
                .unwrap(),
            empty.bytes()
        );
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            "application/vnd.openxmlformats-package.digital-signature-origin".into(),
            vec![],
        )));
        assert!(!package.try_replace_relationships(&empty, &empty).unwrap());
        assert!(matches!(
            package.try_replace_relationships(&empty, &before),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert_eq!(package.source_relationships(&owner).unwrap(), empty);
    }

    #[test]
    fn bounded_relationship_replacement_checks_token_before_mutation() {
        let mut package = OpcPackage::from_vec(source(XML)).unwrap();
        let owner = PackURI::new("/custom/item.bin").unwrap();
        let before = package.source_relationships(&owner).unwrap();
        let after = before.without_relationship("rId1", XML.len()).unwrap();
        assert!(package.try_replace_relationships(&before, &after).unwrap());
        let current = package.source_relationships(&owner).unwrap();
        let before_bytes = PackageWriter::to_bytes(&package).unwrap();

        let under = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .max_total_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.try_replace_relationships_with_limits(&current, &before, under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                actual,
                maximum,
            }) if actual == XML.len() as u64 && maximum == (XML.len() - 1) as u64
        ));
        assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before_bytes);

        let exact = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_total_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_part_bytes(XML.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .try_replace_relationships_with_limits(&current, &before, exact)
                .unwrap()
        );
        assert_eq!(
            package
                .source_relationships_with_limits(&owner, exact)
                .unwrap(),
            before
        );
    }

    #[test]
    fn bounded_relationship_capture_checks_retained_and_authored_quotas() {
        let archive = source(XML);
        let package = OpcPackage::from_bytes(&archive).unwrap();
        let root = PackURI::new("/").unwrap();
        let exact = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len())
            .unwrap()
            .max_part_bytes(XML.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            package
                .source_relationships_with_limits(&root, exact)
                .unwrap()
                .bytes(),
            XML
        );

        let root_member = root.rels_uri().unwrap().membername().len();
        let root_name_exact = ReadLimits::builder()
            .max_archive_member_name_bytes(root_member as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&root, root_name_exact)
                .is_ok()
        );
        let root_name_under = ReadLimits::builder()
            .max_archive_member_name_bytes((root_member - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&root, root_name_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMemberNameBytes,
                actual,
                maximum,
            }) if actual == root_member as u64 && maximum == (root_member - 1) as u64
        ));

        let custom_owner = PackURI::new("/custom/item.bin").unwrap();
        let custom_member = custom_owner.rels_uri().unwrap().membername().len();
        let custom_name_exact = ReadLimits::builder()
            .max_archive_member_name_bytes(custom_member as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&custom_owner, custom_name_exact)
                .is_ok()
        );
        let custom_name_under = ReadLimits::builder()
            .max_archive_member_name_bytes((custom_member - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&custom_owner, custom_name_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveMemberNameBytes,
                actual,
                maximum,
            }) if actual == custom_member as u64 && maximum == (custom_member - 1) as u64
        ));

        let under = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_relationships_with_limits(&root, under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipXmlBytes,
                ..
            })
        ));

        let over = ReadLimits::builder()
            .max_relationship_xml_bytes(XML.len() + 1)
            .unwrap()
            .max_part_bytes((XML.len() + 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .source_relationships_with_limits(&root, over)
                .is_ok()
        );

        let mut authored = OpcPackage::new();
        authored
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/a&b".to_owned(),
                "rId1".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let authored_token = authored.source_relationships(&root).unwrap();
        let authored_exact = ReadLimits::builder()
            .max_relationship_xml_bytes(authored_token.bytes().len())
            .unwrap()
            .max_part_bytes(authored_token.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            authored
                .source_relationships_with_limits(&root, authored_exact)
                .unwrap()
                .bytes(),
            authored_token.bytes()
        );

        // Canonical source capture must admit the ZIP-entry ceiling before
        // asking the serializer for its output buffer. The neutral plan and
        // the fresh capture must reject the same under-limit profile.
        let archive_entry_under = ReadLimits::builder()
            .max_archive_entry_bytes(authored_token.bytes().len() as u64 - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.plan_source_relationships_with_limits(&root, archive_entry_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveEntryBytes,
                ..
            })
        ));
        assert!(matches!(
            authored.source_relationships_with_limits(&root, archive_entry_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::ArchiveEntryBytes,
                ..
            })
        ));

        // Authored relationship fields are checked against the shared XML
        // 1.0 character predicate before canonical length/materialization.
        let mut invalid_xml = OpcPackage::new();
        invalid_xml
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "target\u{fffe}".to_owned(),
                "rId1".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        assert!(matches!(
            invalid_xml.plan_source_relationships_with_limits(&root, ReadLimits::default()),
            Err(OpcError::InvalidRelationship(_))
        ));
        assert!(matches!(
            invalid_xml.source_relationships_with_limits(&root, ReadLimits::default()),
            Err(OpcError::InvalidRelationship(_))
        ));

        let event_under = ReadLimits::builder()
            .max_xml_events(4)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, event_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlEvents,
                actual: 5,
                maximum: 4,
            })
        ));
        let depth_under = ReadLimits::builder()
            .max_xml_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, depth_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::XmlDepth,
                actual: 2,
                maximum: 1,
            })
        ));
        let aggregate_bytes_under = ReadLimits::builder()
            .max_total_relationship_xml_bytes(authored_token.bytes().len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_relationships_with_limits(&root, aggregate_bytes_under),
            Err(OpcError::ReadLimit {
                resource: ReadResource::TotalRelationshipXmlBytes,
                ..
            })
        ));

        let mut too_many = authored.clone();
        too_many
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/c".to_owned(),
                "rId2".to_owned(),
                TargetMode::External,
            )
            .unwrap();
        let count_limit = ReadLimits::builder()
            .max_relationships_per_part(1)
            .unwrap()
            .max_total_relationships(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            too_many.source_relationships_with_limits(&root, count_limit),
            Err(OpcError::ReadLimit {
                resource: ReadResource::RelationshipsPerPart,
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn source_canonical_capture_rejects_the_fixed_owned_xml_ceiling_before_materialization() {
        let target_bytes = (32 * 1024 * 1024) + 1;
        let output_limit = target_bytes + 4096;
        let limits = ReadLimits::builder()
            .max_relationship_xml_bytes(output_limit)
            .unwrap()
            .max_total_relationship_xml_bytes(output_limit)
            .unwrap()
            .max_part_bytes(output_limit as u64)
            .unwrap()
            .max_archive_entry_bytes(output_limit as u64)
            .unwrap()
            .max_xml_attribute_bytes(output_limit)
            .unwrap()
            .max_relationship_target_bytes(target_bytes)
            .unwrap()
            .build()
            .unwrap();
        let mut package = OpcPackage::new();
        package
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "x".repeat(target_bytes),
                "rId1".to_owned(),
                TargetMode::External,
            )
            .unwrap();

        // The neutral plan is output-free and may be admitted by the caller's
        // raised limits. Do not materialize its deliberately oversized XML.
        let canonical = package.rels().plan_canonical(limits).unwrap();
        assert!(canonical.final_len() > 32 * 1024 * 1024);

        let root = PackURI::new("/").unwrap();
        assert!(matches!(
            package.plan_source_relationships_with_limits(&root, limits),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
        assert!(matches!(
            package.source_relationships_with_limits(&root, limits),
            Err(OpcError::SourceBackedOverlayUnavailable { .. })
        ));
    }
}
