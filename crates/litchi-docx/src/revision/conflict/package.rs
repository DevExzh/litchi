//! Story discovery and failure-atomic package publication.

use super::model::{Binding, Source};
use super::transaction::{Commit, Patch};
use super::{Limits, Snapshot};
use crate::{Error, Package, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::{HashSet, VecDeque};
use std::sync::Arc;

const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/";
const TOPOLOGY_MAGIC: &[u8] = b"litchi.docx.conflict.topology.v1\0";

/// One reachable WordprocessingML story and its conflict snapshot.
#[derive(Clone, Debug)]
pub struct Story {
    part: PackURI,
    content_type: String,
    snapshot: Snapshot,
}

impl Story {
    /// Canonical OPC part name that owns this story.
    #[must_use]
    pub const fn part(&self) -> &PackURI {
        &self.part
    }

    /// Declared content type of the story part.
    #[must_use]
    pub fn content_type(&self) -> &str {
        &self.content_type
    }

    /// Bounded semantic conflict snapshot for this story.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Consume the wrapper and retain the story-bound snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

impl Package {
    /// Read the conflict revisions in the resolved main-document story.
    pub fn conflicts(&self) -> Result<Snapshot> {
        self.conflicts_with_limits(Limits::default())
    }

    /// Read the main-story conflicts with explicit semantic resource limits.
    pub fn conflicts_with_limits(&self, limits: Limits) -> Result<Snapshot> {
        let limits = limits.validate()?;
        self.ensure_conflict_opc_current("conflicts_with_limits")?;
        let captured = capture(self.opc_package(), limits, false)?;
        let story = captured
            .stories
            .iter()
            .find(|story| story.part == captured.main)
            .ok_or_else(|| invalid("resolved main document is missing from the story inventory"))?;
        snapshot_single(story, Arc::clone(&captured.topology), limits)
    }

    /// Read every reachable WordprocessingML story independently.
    ///
    /// Stories are returned with the main document first and remaining parts
    /// ordered by canonical part name. Conflict ranges never pair across two
    /// entries in this inventory.
    pub fn conflict_stories(&self) -> Result<Vec<Story>> {
        self.conflict_stories_with_limits(Limits::default())
    }

    /// Read every reachable story with explicit semantic resource limits.
    pub fn conflict_stories_with_limits(&self, limits: Limits) -> Result<Vec<Story>> {
        let limits = limits.validate()?;
        self.ensure_conflict_opc_current("conflict_stories_with_limits")?;
        let captured = capture(self.opc_package(), limits, true)?;
        let mut stories = Vec::new();
        stories
            .try_reserve_exact(captured.stories.len())
            .map_err(|source| Error::Allocation {
                resource: "DOCX conflict story inventory",
                source,
            })?;
        let mut aggregate = Aggregate::default();
        for located in &captured.stories {
            let parsed = snapshot_accounted(
                located,
                Arc::clone(&captured.topology),
                limits,
                &mut aggregate,
                None,
            )?;
            stories.push(Story {
                part: located.part.clone(),
                content_type: located.content_type.clone(),
                snapshot: parsed,
            });
        }
        Ok(stories)
    }

    /// Publish a prepared story-bound conflict transaction.
    ///
    /// The commit remains reusable when publication fails.
    pub fn apply_conflicts(&mut self, commit: &Commit) -> Result<Snapshot> {
        self.apply_conflict_patch(commit.patch())
    }

    /// Publish a reversible story-bound conflict patch.
    ///
    /// Exact no-ops bypass raw OPC editing, retaining the original blob `Arc`,
    /// signatures, and managed-writer state. Real changes clone-stage the OPC
    /// graph through an internal semantic publication path, replace only the
    /// bound story blob,
    /// and reparse it before the candidate is published.
    pub fn apply_conflict_patch(&mut self, patch: &Patch) -> Result<Snapshot> {
        let claim = patch.claim_publication()?;
        self.ensure_conflict_opc_current("apply_conflict_patch")?;
        let binding = patch
            .story_binding()
            .ok_or_else(|| invalid("conflict patch is not bound to a package story"))?
            .clone();
        let source_limits = patch.source_limits();
        let target_limits = patch.target_limits();
        let before = patch.before_source();
        let after = patch.after_source();
        let captured = capture(self.opc_package(), source_limits, true)?;
        let current = bound_story(&captured, &binding)?;
        if current.source.as_slice() != before.as_ref() {
            return Err(invalid("conflict patch source story is stale"));
        }
        let current_selection =
            parse_selected_with_totals(&captured, current.part.as_str(), source_limits, None)?;
        if patch.is_noop() {
            claim.finalize();
            return Ok(current_selection.snapshot);
        }

        // Validate the complete target before the package or its signatures
        // can change. This is deliberately independent of XML serialization.
        preflight_source_len(after.len(), target_limits)?;
        preflight_replacement_total(&captured, current.source.len(), after.len(), target_limits)?;
        let mut replacement_aggregate = current_selection
            .totals
            .without(current_selection.selected)?;
        let expected = snapshot_accounted(
            current,
            Arc::clone(&captured.topology),
            target_limits,
            &mut replacement_aggregate,
            Some(&after),
        )?;
        if expected.source() != after.as_ref() {
            return Err(invalid("conflict patch target did not round-trip exactly"));
        }

        let target = current.part.clone();
        let replacement = source_blob(after)?;
        let expected_topology = Arc::clone(&captured.topology);
        let expected_published = expected.clone();
        let published = self.edit_semantic_opc("apply_conflict_patch", move |candidate| {
            let staged = capture(candidate, source_limits, true)?;
            if staged.topology.as_ref() != expected_topology.as_ref() {
                return Err(invalid("conflict story topology is stale"));
            }
            let source = bound_story(&staged, &binding)?;
            if source.part != target || source.source.as_slice() != before.as_ref() {
                return Err(invalid("conflict patch source story is stale"));
            }
            preflight_replacement_total(
                &staged,
                source.source.len(),
                replacement.len(),
                target_limits,
            )?;
            candidate
                .get_part_mut(&target)?
                .set_blob_shared(Arc::clone(&replacement));
            Ok(expected_published)
        })?;
        claim.finalize();
        Ok(published)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum Role {
    Header,
    Footer,
    Footnotes,
    Endnotes,
    Comments,
    Glossary,
}

impl Role {
    const fn expected_content_type(self) -> &'static str {
        match self {
            Self::Header => ct::WML_HEADER,
            Self::Footer => ct::WML_FOOTER,
            Self::Footnotes => ct::WML_FOOTNOTES,
            Self::Endnotes => ct::WML_ENDNOTES,
            Self::Comments => ct::WML_COMMENTS,
            Self::Glossary => ct::WML_DOCUMENT_GLOSSARY,
        }
    }

    const fn singleton_index(self) -> Option<usize> {
        match self {
            Self::Header | Self::Footer => None,
            Self::Footnotes => Some(0),
            Self::Endnotes => Some(1),
            Self::Comments => Some(2),
            Self::Glossary => Some(3),
        }
    }

    const fn code(self) -> u8 {
        match self {
            Self::Header => 1,
            Self::Footer => 2,
            Self::Footnotes => 3,
            Self::Endnotes => 4,
            Self::Comments => 5,
            Self::Glossary => 6,
        }
    }
}

#[derive(Debug)]
struct Located {
    part: PackURI,
    content_type: String,
    source: Arc<Vec<u8>>,
}

#[derive(Debug)]
struct Edge {
    source: PackURI,
    id: String,
    reltype: String,
    target_ref: String,
    target: PackURI,
    role: Role,
}

#[derive(Debug)]
struct RootEdge {
    id: String,
    reltype: String,
    target_ref: String,
}

#[derive(Debug)]
struct Capture {
    main: PackURI,
    stories: Vec<Located>,
    topology: Arc<[u8]>,
    total_story_bytes: usize,
}

fn capture(package: &OpcPackage, limits: Limits, enforce_story_bytes: bool) -> Result<Capture> {
    let main_part = package.main_document_part()?;
    validate_main_content_type(main_part.content_type())?;
    let main = main_part.partname().clone();
    let root = root_edge(package)?;

    let capacity = package.part_count().min(limits.max_stories);
    let mut visited = HashSet::new();
    visited
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX conflict story visited set",
            source,
        })?;
    let mut queue = VecDeque::new();
    queue
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX conflict story queue",
            source,
        })?;
    let mut stories = Vec::new();
    stories
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "DOCX conflict story parts",
            source,
        })?;
    let mut edges = Vec::new();
    let mut total_story_bytes = 0usize;

    visited.insert(main.clone());
    queue.push_back(main.clone());
    while let Some(name) = queue.pop_front() {
        if stories.len() == limits.max_stories {
            return Err(invalid(format!(
                "DOCX conflict story count exceeds {}",
                limits.max_stories
            )));
        }
        let part = package.get_part(&name)?;
        if enforce_story_bytes {
            preflight_source_len(part.blob().len(), limits)?;
            total_story_bytes = checked_total(
                total_story_bytes,
                part.blob().len(),
                limits.max_total_story_bytes,
                "story byte",
            )?;
        } else {
            total_story_bytes = total_story_bytes
                .checked_add(part.blob().len())
                .ok_or_else(|| invalid("DOCX conflict aggregate story byte count overflow"))?;
        }
        stories.push(Located {
            part: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            source: part.blob_arc(),
        });

        let mut relationships = Vec::new();
        let mut story_relationships = 0usize;
        for relationship in part.rels().iter() {
            let Some(role) = relationship_role(relationship.reltype()) else {
                continue;
            };
            if story_relationships == limits.max_relationships_per_story {
                return Err(invalid(format!(
                    "story '{}' relationship count exceeds {}",
                    part.partname(),
                    limits.max_relationships_per_story
                )));
            }
            if edges
                .len()
                .checked_add(relationships.len())
                .is_none_or(|count| count == limits.max_total_relationships)
            {
                return Err(invalid(format!(
                    "DOCX conflict story relationship count exceeds {}",
                    limits.max_total_relationships
                )));
            }
            relationships
                .try_reserve(1)
                .map_err(|source| Error::Allocation {
                    resource: "DOCX conflict story relationships",
                    source,
                })?;
            relationships.push((role, relationship));
            story_relationships += 1;
        }
        relationships.sort_unstable_by(|left, right| {
            left.1
                .r_id()
                .cmp(right.1.r_id())
                .then_with(|| left.1.reltype().cmp(right.1.reltype()))
                .then_with(|| left.1.target_ref().cmp(right.1.target_ref()))
        });

        let mut singleton = [false; 4];
        for (role, relationship) in relationships {
            if let Some(index) = role.singleton_index()
                && std::mem::replace(&mut singleton[index], true)
            {
                return Err(invalid(format!(
                    "story '{}' has multiple {:?} relationships",
                    part.partname(),
                    role
                )));
            }
            if relationship.is_external() {
                return Err(invalid(format!(
                    "story '{}' has an external {:?} relationship",
                    part.partname(),
                    role
                )));
            }
            let requested = relationship.target_partname()?;
            let target_part = package.get_part(&requested).map_err(|_| {
                invalid(format!(
                    "story '{}' {:?} relationship targets missing part '{}'",
                    part.partname(),
                    role,
                    requested
                ))
            })?;
            if target_part.content_type() != role.expected_content_type() {
                return Err(invalid(format!(
                    "story '{}' {:?} relationship targets content type '{}'",
                    part.partname(),
                    role,
                    target_part.content_type()
                )));
            }
            edges.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "DOCX conflict story topology edges",
                source,
            })?;
            let target = target_part.partname().clone();
            edges.push(Edge {
                source: part.partname().clone(),
                id: relationship.r_id().to_owned(),
                reltype: relationship.reltype().to_owned(),
                target_ref: relationship.target_ref().to_owned(),
                target: target.clone(),
                role,
            });
            if !visited.contains(&target) {
                if visited.len() == limits.max_stories {
                    return Err(invalid(format!(
                        "DOCX conflict story count exceeds {}",
                        limits.max_stories
                    )));
                }
                visited.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX conflict story visited set",
                    source,
                })?;
                queue.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "DOCX conflict story queue",
                    source,
                })?;
                visited.insert(target.clone());
                queue.push_back(target);
            }
        }
    }

    stories.sort_unstable_by(|left, right| left.part.as_str().cmp(right.part.as_str()));
    let main_index = stories
        .iter()
        .position(|story| story.part == main)
        .ok_or_else(|| invalid("resolved main document disappeared during story capture"))?;
    stories.swap(0, main_index);
    edges.sort_unstable_by(|left, right| {
        left.source
            .as_str()
            .cmp(right.source.as_str())
            .then_with(|| left.id.cmp(&right.id))
            .then_with(|| left.reltype.cmp(&right.reltype))
            .then_with(|| left.target_ref.cmp(&right.target_ref))
    });
    let topology = topology(&root, &stories, &edges, limits.max_topology_bytes)?;
    Ok(Capture {
        main,
        stories,
        topology,
        total_story_bytes,
    })
}

fn root_edge(package: &OpcPackage) -> Result<RootEdge> {
    let mut relationships = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = relationships
        .next()
        .ok_or_else(|| invalid("main-document relationship is missing"))?;
    if relationships.next().is_some() {
        return Err(invalid("package has multiple main-document relationships"));
    }
    if relationship.is_external() {
        return Err(invalid("main-document relationship cannot be external"));
    }
    Ok(RootEdge {
        id: relationship.r_id().to_owned(),
        reltype: relationship.reltype().to_owned(),
        target_ref: relationship.target_ref().to_owned(),
    })
}

fn validate_main_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MAIN
            | ct::WML_TEMPLATE_MAIN
            | ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "main document has unsupported content type '{content_type}'"
        )))
    }
}

fn relationship_role(value: &str) -> Option<Role> {
    let suffix = value
        .strip_prefix(TRANSITIONAL_RELATIONSHIPS)
        .or_else(|| value.strip_prefix(STRICT_RELATIONSHIPS))?;
    match suffix {
        "header" => Some(Role::Header),
        "footer" => Some(Role::Footer),
        "footnotes" => Some(Role::Footnotes),
        "endnotes" => Some(Role::Endnotes),
        "comments" => Some(Role::Comments),
        "glossaryDocument" => Some(Role::Glossary),
        _ => None,
    }
}

fn topology(
    root: &RootEdge,
    stories: &[Located],
    edges: &[Edge],
    limit: usize,
) -> Result<Arc<[u8]>> {
    let size = topology_size(root, stories, edges)?;
    if size > limit {
        return Err(invalid(format!(
            "DOCX conflict story topology exceeds {limit} bytes"
        )));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(size)
        .map_err(|source| Error::Allocation {
            resource: "DOCX conflict story topology",
            source,
        })?;
    bytes.extend_from_slice(TOPOLOGY_MAGIC);
    put_field(&mut bytes, root.id.as_bytes());
    put_field(&mut bytes, root.reltype.as_bytes());
    put_field(&mut bytes, root.target_ref.as_bytes());
    put_number(&mut bytes, stories.len());
    for story in stories {
        put_field(&mut bytes, story.part.as_str().as_bytes());
        put_field(&mut bytes, story.content_type.as_bytes());
    }
    put_number(&mut bytes, edges.len());
    for edge in edges {
        put_field(&mut bytes, edge.source.as_str().as_bytes());
        put_field(&mut bytes, edge.id.as_bytes());
        put_field(&mut bytes, edge.reltype.as_bytes());
        put_field(&mut bytes, edge.target_ref.as_bytes());
        put_field(&mut bytes, edge.target.as_str().as_bytes());
        bytes.push(edge.role.code());
    }
    debug_assert_eq!(bytes.len(), size);
    Ok(Arc::from(bytes.into_boxed_slice()))
}

fn topology_size(root: &RootEdge, stories: &[Located], edges: &[Edge]) -> Result<usize> {
    let mut size = TOPOLOGY_MAGIC.len();
    for value in [
        root.id.as_bytes(),
        root.reltype.as_bytes(),
        root.target_ref.as_bytes(),
    ] {
        size = add_field_size(size, value.len())?;
    }
    size = add_size(size, 8)?;
    for story in stories {
        size = add_field_size(size, story.part.as_str().len())?;
        size = add_field_size(size, story.content_type.len())?;
    }
    size = add_size(size, 8)?;
    for edge in edges {
        for length in [
            edge.source.as_str().len(),
            edge.id.len(),
            edge.reltype.len(),
            edge.target_ref.len(),
            edge.target.as_str().len(),
        ] {
            size = add_field_size(size, length)?;
        }
        size = add_size(size, 1)?;
    }
    Ok(size)
}

fn add_field_size(size: usize, length: usize) -> Result<usize> {
    add_size(add_size(size, 8)?, length)
}

fn add_size(size: usize, added: usize) -> Result<usize> {
    size.checked_add(added)
        .ok_or_else(|| invalid("DOCX conflict topology size overflow"))
}

fn put_number(output: &mut Vec<u8>, value: usize) {
    output.extend_from_slice(&(value as u64).to_le_bytes());
}

fn put_field(output: &mut Vec<u8>, value: &[u8]) {
    put_number(output, value.len());
    output.extend_from_slice(value);
}

#[derive(Clone, Copy, Default)]
struct Aggregate {
    story_bytes: usize,
    conflicts: usize,
    ranges: usize,
    metadata_bytes: usize,
    text_segments: usize,
}

struct Selection {
    snapshot: Snapshot,
    totals: Aggregate,
    selected: Aggregate,
}

fn snapshot_single(story: &Located, topology: Arc<[u8]>, limits: Limits) -> Result<Snapshot> {
    preflight_source_len(story.source.len(), limits)?;
    Snapshot::from_source_with_parse_and_retained_limits(
        Source::Blob(Arc::clone(&story.source)),
        limits,
        limits,
    )
    .map(|snapshot| {
        snapshot.with_binding(Binding::new(
            story.part.as_str().to_owned(),
            story.content_type.clone(),
            topology,
        ))
    })
}

fn parse_selected_with_totals(
    captured: &Capture,
    selected: &str,
    limits: Limits,
    replacement: Option<&Source>,
) -> Result<Selection> {
    let mut aggregate = Aggregate::default();
    let mut selected_snapshot = None;
    let mut selected_contribution = None;
    for story in &captured.stories {
        let source = if story.part.as_str() == selected {
            replacement
        } else {
            None
        };
        let parsed = snapshot_accounted(
            story,
            Arc::clone(&captured.topology),
            limits,
            &mut aggregate,
            source,
        )?;
        if story.part.as_str() == selected {
            selected_contribution = Some(contribution(&parsed, parsed.source().len())?);
            selected_snapshot = Some(parsed);
        }
    }
    Ok(Selection {
        snapshot: selected_snapshot.ok_or_else(|| invalid("selected conflict story is missing"))?,
        totals: aggregate,
        selected: selected_contribution
            .ok_or_else(|| invalid("selected conflict story contribution is missing"))?,
    })
}

fn snapshot_accounted(
    story: &Located,
    topology: Arc<[u8]>,
    limits: Limits,
    aggregate: &mut Aggregate,
    replacement: Option<&Source>,
) -> Result<Snapshot> {
    let length = replacement.map_or_else(|| story.source.len(), Source::len);
    let mut effective = limits;
    effective.max_source_bytes = effective.max_source_bytes.min(remaining(
        limits.max_total_story_bytes,
        aggregate.story_bytes,
        "story bytes",
    )?);
    effective.max_conflicts = effective.max_conflicts.min(remaining(
        limits.max_total_conflicts,
        aggregate.conflicts,
        "conflict elements",
    )?);
    effective.max_ranges = effective.max_ranges.min(remaining(
        limits.max_total_ranges,
        aggregate.ranges,
        "conflict ranges",
    )?);
    effective.max_metadata_bytes = effective.max_metadata_bytes.min(remaining(
        limits.max_total_metadata_bytes,
        aggregate.metadata_bytes,
        "metadata bytes",
    )?);
    effective.max_text_segments = effective.max_text_segments.min(remaining(
        limits.max_total_text_segments,
        aggregate.text_segments,
        "text segments",
    )?);
    effective = effective.validate()?;
    preflight_source_len(length, effective)?;
    let source = replacement
        .cloned()
        .unwrap_or_else(|| Source::Blob(Arc::clone(&story.source)));
    let parsed = Snapshot::from_source_with_parse_and_retained_limits(source, effective, limits)?
        .with_binding(Binding::new(
            story.part.as_str().to_owned(),
            story.content_type.clone(),
            topology,
        ));

    aggregate.add(contribution(&parsed, length)?, limits)?;
    Ok(parsed)
}

impl Aggregate {
    fn add(&mut self, value: Self, limits: Limits) -> Result<()> {
        self.story_bytes = checked_total(
            self.story_bytes,
            value.story_bytes,
            limits.max_total_story_bytes,
            "story byte",
        )?;
        self.conflicts = checked_total(
            self.conflicts,
            value.conflicts,
            limits.max_total_conflicts,
            "conflict element",
        )?;
        self.ranges = checked_total(
            self.ranges,
            value.ranges,
            limits.max_total_ranges,
            "conflict range",
        )?;
        self.metadata_bytes = checked_total(
            self.metadata_bytes,
            value.metadata_bytes,
            limits.max_total_metadata_bytes,
            "metadata byte",
        )?;
        self.text_segments = checked_total(
            self.text_segments,
            value.text_segments,
            limits.max_total_text_segments,
            "text segment",
        )?;
        Ok(())
    }

    fn without(self, value: Self) -> Result<Self> {
        Ok(Self {
            story_bytes: subtract(self.story_bytes, value.story_bytes, "story bytes")?,
            conflicts: subtract(self.conflicts, value.conflicts, "conflict elements")?,
            ranges: subtract(self.ranges, value.ranges, "conflict ranges")?,
            metadata_bytes: subtract(self.metadata_bytes, value.metadata_bytes, "metadata bytes")?,
            text_segments: subtract(self.text_segments, value.text_segments, "text segments")?,
        })
    }
}

fn contribution(snapshot: &Snapshot, story_bytes: usize) -> Result<Aggregate> {
    let inventory = snapshot.inventory();
    let mut text_segments = 0usize;
    for conflict in &inventory.conflicts {
        text_segments = text_segments
            .checked_add(conflict.text_spans().len())
            .ok_or_else(|| invalid("DOCX conflict aggregate text segment count overflow"))?;
    }
    Ok(Aggregate {
        story_bytes,
        conflicts: inventory.conflicts.len(),
        ranges: inventory.ranges.len(),
        metadata_bytes: metadata_bytes(inventory)?,
        text_segments,
    })
}

fn subtract(total: usize, removed: usize, resource: &'static str) -> Result<usize> {
    total
        .checked_sub(removed)
        .ok_or_else(|| invalid(format!("DOCX conflict aggregate {resource} underflow")))
}

fn remaining(limit: usize, used: usize, resource: &'static str) -> Result<usize> {
    limit.checked_sub(used).ok_or_else(|| {
        invalid(format!(
            "DOCX conflict aggregate {resource} limit is exhausted"
        ))
    })
}

fn metadata_bytes(inventory: &super::Inventory) -> Result<usize> {
    let mut bytes = 0usize;
    for metadata in inventory
        .conflicts
        .iter()
        .map(|conflict| &conflict.metadata)
        .chain(inventory.ranges.iter().map(|range| &range.metadata))
    {
        bytes = bytes
            .checked_add(metadata.author.len())
            .and_then(|value| value.checked_add(metadata.date.as_deref().map_or(0, str::len)))
            .ok_or_else(|| invalid("DOCX conflict aggregate metadata size overflow"))?;
    }
    Ok(bytes)
}

fn preflight_source_len(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_source_bytes {
        return Err(invalid(format!(
            "DOCX conflict story has {length} bytes, exceeding {}",
            limits.max_source_bytes
        )));
    }
    Ok(())
}

fn preflight_replacement_total(
    captured: &Capture,
    before: usize,
    after: usize,
    limits: Limits,
) -> Result<()> {
    let total = captured
        .total_story_bytes
        .checked_sub(before)
        .and_then(|value| value.checked_add(after))
        .ok_or_else(|| invalid("DOCX conflict aggregate story size overflow"))?;
    if total > limits.max_total_story_bytes {
        return Err(invalid(format!(
            "DOCX conflict aggregate story bytes exceed {}",
            limits.max_total_story_bytes
        )));
    }
    Ok(())
}

fn source_blob(source: Source) -> Result<Arc<Vec<u8>>> {
    match source {
        Source::Blob(source) => Ok(source),
        Source::Detached(_) => Err(invalid(
            "package-bound conflict patch target is not OPC-native",
        )),
    }
}

fn checked_total(
    total: usize,
    added: usize,
    limit: usize,
    resource: &'static str,
) -> Result<usize> {
    let total = total
        .checked_add(added)
        .ok_or_else(|| invalid(format!("DOCX conflict aggregate {resource} count overflow")))?;
    if total > limit {
        return Err(invalid(format!(
            "DOCX conflict aggregate {resource} count exceeds {limit}"
        )));
    }
    Ok(total)
}

fn bound_story<'a>(captured: &'a Capture, binding: &Binding) -> Result<&'a Located> {
    if captured.topology.as_ref() != binding.topology().as_ref() {
        return Err(invalid("conflict story topology is stale"));
    }
    let story = captured
        .stories
        .iter()
        .find(|story| story.part.as_str() == binding.part())
        .ok_or_else(|| invalid("bound conflict story part is missing"))?;
    if story.content_type != binding.content_type() {
        return Err(invalid("bound conflict story content type is stale"));
    }
    Ok(story)
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
