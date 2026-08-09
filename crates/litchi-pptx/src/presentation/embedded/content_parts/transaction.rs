//! Detached, source-checked transactions for one slide-owned content-part graph.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI, TargetMode};

use super::codec;
use super::model::{Anchor, ContentPart, Payload, Relationship, Target};
use super::validation;
use crate::{Error, Result};

/// Stable fingerprint of the complete detached slide-owned source graph.
pub type Revision = u64;

/// Relationship metadata captured from one owning part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct RelationshipState {
    pub(crate) id: String,
    pub(crate) relationship_type: String,
    pub(crate) target_ref: String,
    pub(crate) target_mode: TargetMode,
}

/// An immutable semantic view of one slide's content-part graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) slide_index: usize,
    pub(crate) slide_part_name: PackURI,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) parts: Arc<[ContentPart]>,
    pub(crate) slide_relationships: Arc<[RelationshipState]>,
    pub(crate) payloads: Arc<[Payload]>,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load a source-checked snapshot from one `PresentationML` slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    #[inline]
    pub fn load(
        package: &OpcPackage,
        slide_index: usize,
        slide: &dyn litchi_opc::Part,
        limits: &mut super::package::Limits,
    ) -> Result<Self> {
        super::package::load_snapshot(package, slide_index, slide, limits)
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "snapshot constructor mirrors the persisted part fields"
    )]
    pub(crate) fn from_parts(
        slide_index: usize,
        slide_part_name: PackURI,
        source_xml: Arc<Vec<u8>>,
        mut parts: Vec<ContentPart>,
        slide_relationships: Vec<RelationshipState>,
    ) -> Result<Self> {
        if source_xml.len() > super::super::MAX_XML_BYTES {
            return Err(validation::limit_xml());
        }
        for (index, part) in parts.iter_mut().enumerate() {
            part.slide_index = slide_index;
            part.slide_part_name = slide_part_name.clone();
            part.index = index;
        }
        validation::validate_parts(&parts)?;
        let slide_relationships = sorted_relationships(slide_relationships)?;
        for relationship in &slide_relationships {
            validation::validate_field(
                &relationship.relationship_type,
                "content-part slide relationship type",
            )?;
            validation::validate_field(
                &relationship.target_ref,
                "content-part slide relationship target",
            )?;
            validation::validate_id(&relationship.id, "content-part slide relationship ID")?;
        }
        let payloads = unique_payloads(&parts)?;
        let revision = fingerprint(
            source_xml.as_slice(),
            &parts,
            &slide_relationships,
            &payloads,
        );
        Ok(Self {
            slide_index,
            slide_part_name,
            source_xml,
            parts: Arc::from(parts),
            slide_relationships: Arc::from(slide_relationships),
            payloads: Arc::from(payloads),
            revision,
        })
    }

    /// Zero-based presentation position of the owning slide.
    #[inline]
    #[must_use]
    pub const fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Absolute OPC part name of the owning slide.
    #[inline]
    #[must_use]
    pub fn slide_part_name(&self) -> &PackURI {
        &self.slide_part_name
    }

    /// Content parts in active `PresentationML` order.
    #[inline]
    #[must_use]
    pub fn parts(&self) -> &[ContentPart] {
        &self.parts
    }

    /// Alias emphasizing the slide-owned semantic inventory.
    #[inline]
    #[must_use]
    pub fn content_parts(&self) -> &[ContentPart] {
        self.parts()
    }

    /// Exact source slide XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Compact source revision used for optimistic stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start a detached transaction over this slide graph.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            working: self.parts.as_ref().to_vec(),
            origins: (0..self.parts.len()).map(Some).collect(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.slide_index == other.slide_index
            && self.slide_part_name == other.slide_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.parts.as_ref() == other.parts.as_ref()
            && self.slide_relationships.as_ref() == other.slide_relationships.as_ref()
            && self.payloads.as_ref() == other.payloads.as_ref()
            && self.revision == other.revision
    }
}

/// Failure-atomic editor over typed anchor, relationship, and payload values.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    working: Vec<ContentPart>,
    origins: Vec<Option<usize>>,
}

impl Transaction {
    /// Borrow the immutable source snapshot used by this edit.
    #[inline]
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the currently staged content-part inventory.
    #[inline]
    #[must_use]
    pub fn parts(&self) -> &[ContentPart] {
        &self.working
    }

    /// Alias emphasizing the semantic inventory.
    #[inline]
    #[must_use]
    pub fn content_parts(&self) -> &[ContentPart] {
        self.parts()
    }

    /// Whether the staged graph differs from the source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.working.len() != self.source.parts.len()
            || self
                .working
                .iter()
                .zip(&self.origins)
                .any(|(part, origin)| match origin {
                    Some(index) => self.source.parts.get(*index) != Some(part),
                    None => true,
                })
    }

    /// Replace one anchor and relationship graph atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(
        &mut self,
        index: usize,
        anchor: Anchor,
        relationship: Relationship,
    ) -> Result<bool> {
        let mut candidate = self.part(index)?.clone();
        candidate.anchor = anchor;
        candidate.relationship = relationship;
        self.validate_candidate(&candidate, Some(index))?;
        if candidate == self.working[index] {
            return Ok(false);
        }
        self.working[index] = candidate;
        Ok(true)
    }

    /// Update only the anchor's relationship ID while preserving raw markup.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_relationship_id(&mut self, index: usize, value: impl Into<String>) -> Result<bool> {
        let value = value.into();
        validation::validate_id(&value, "content-part relationship ID")?;
        let mut candidate = self.part(index)?.clone();
        if candidate.relationship_id() == value {
            return Ok(false);
        }
        candidate.anchor.xml =
            codec::rewrite_anchor_relationship_id(candidate.anchor.xml(), &value)?;
        candidate.anchor.relationship_id.clone_from(&value);
        candidate.relationship.id = value;
        self.validate_candidate(&candidate, Some(index))?;
        self.working[index] = candidate;
        Ok(true)
    }

    /// Update the relationship type without rewriting the owning slide XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_relationship_type(
        &mut self,
        index: usize,
        value: impl Into<String>,
    ) -> Result<bool> {
        let value = value.into();
        validation::validate_field(&value, "content-part relationship type")?;
        let mut candidate = self.part(index)?.clone();
        if candidate.relationship.relationship_type == value {
            return Ok(false);
        }
        candidate.relationship.relationship_type = value;
        self.validate_candidate(&candidate, Some(index))?;
        self.working[index] = candidate;
        Ok(true)
    }

    /// Update the relationship target reference while keeping its target kind
    /// and validating that internal references still name the opaque payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_target_reference(&mut self, index: usize, value: impl Into<String>) -> Result<bool> {
        let value = value.into();
        validation::validate_field(&value, "content-part relationship target")?;
        if value.is_empty() {
            return Err(invalid("content-part relationship target cannot be empty"));
        }
        let mut candidate = self.part(index)?.clone();
        if candidate.relationship.target_ref == value {
            return Ok(false);
        }
        match candidate.relationship.target_mode {
            TargetMode::Internal => {
                let resolved = litchi_opc::Relationship::new_with_mode(
                    candidate.relationship.id.clone(),
                    candidate.relationship.relationship_type.clone(),
                    value.clone(),
                    candidate.slide_part_name.base_uri().to_owned(),
                    TargetMode::Internal,
                )
                .target_partname()?;
                let payload = candidate
                    .payload()
                    .ok_or_else(|| invalid("internal content-part payload is missing"))?;
                if !resolved.is_equivalent_to(payload.part_name()) {
                    return Err(invalid(
                        "content-part target reference does not resolve to its payload",
                    ));
                }
            },
            TargetMode::External => {
                candidate.relationship.target = Target::external(value.clone());
            },
        }
        candidate.relationship.target_ref = value;
        self.validate_candidate(&candidate, Some(index))?;
        self.working[index] = candidate;
        Ok(true)
    }

    /// Replace the inert target and update its OPC target mode and reference.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_target(&mut self, index: usize, target: Target) -> Result<bool> {
        let mut candidate = self.part(index)?.clone();
        let old_part = candidate
            .payload()
            .map(|payload| payload.part_name().clone());
        match &target {
            Target::Internal(payload) => {
                candidate.relationship.target_mode = TargetMode::Internal;
                if old_part.as_ref() != Some(payload.part_name()) {
                    candidate.relationship.target_ref = payload
                        .part_name()
                        .relative_ref(candidate.slide_part_name.base_uri());
                }
            },
            Target::External { target_ref } => {
                candidate.relationship.target_mode = TargetMode::External;
                candidate.relationship.target_ref.clone_from(target_ref);
            },
        }
        if candidate.target() == &target {
            return Ok(false);
        }
        candidate.relationship.target = target;
        self.validate_candidate(&candidate, Some(index))?;
        self.working[index] = candidate;
        Ok(true)
    }

    /// Replace internal payload bytes and content type while retaining its
    /// part name, relationship ID, and opaque outbound relationship metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn replace_payload(
        &mut self,
        index: usize,
        content_type: impl Into<Arc<str>>,
        bytes: impl Into<Vec<u8>>,
    ) -> Result<bool> {
        let mut candidate = self.part(index)?.clone();
        let payload = candidate
            .payload()
            .ok_or_else(|| invalid("content-part target is external"))?
            .clone();
        let replacement = Payload {
            part_name: payload.part_name,
            content_type: content_type.into(),
            bytes: Arc::<[u8]>::from(bytes.into()),
            relationships: payload.relationships,
        };
        if candidate.payload() == Some(&replacement) {
            return Ok(false);
        }
        candidate.relationship.target = Target::Internal(replacement);
        candidate.relationship.target_mode = TargetMode::Internal;
        self.validate_candidate(&candidate, Some(index))?;
        self.working[index] = candidate;
        Ok(true)
    }

    /// Add one detached content part at an active slide position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert(
        &mut self,
        index: usize,
        anchor: Anchor,
        relationship: Relationship,
    ) -> Result<()> {
        if index > self.working.len() {
            return Err(Error::IndexOutOfBounds {
                index,
                len: self.working.len(),
            });
        }
        let part = self.make_part(index, anchor, relationship)?;
        self.working.insert(index, part);
        self.origins.insert(index, None);
        self.renumber();
        self.validate_all()
    }

    /// Append one detached content part and return its active index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn push(&mut self, anchor: Anchor, relationship: Relationship) -> Result<usize> {
        let index = self.working.len();
        self.insert(index, anchor, relationship)?;
        Ok(index)
    }

    /// Remove one content part and return its detached semantic value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, index: usize) -> Result<ContentPart> {
        if index >= self.working.len() {
            return Err(Error::IndexOutOfBounds {
                index,
                len: self.working.len(),
            });
        }
        let part = self.working.remove(index);
        self.origins.remove(index);
        self.renumber();
        self.validate_all()?;
        Ok(part)
    }

    /// Commit the detached graph into a reversible source patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        self.validate_all()?;
        validate_relationship_collisions(&self.source, &self.working)?;
        let source_xml = build_source(&self.source, &self.working, &self.origins)?;
        let slide_relationships = build_relationships(&self.source, &self.working, &source_xml)?;
        let after = Snapshot::from_parts(
            self.source.slide_index,
            self.source.slide_part_name.clone(),
            Arc::new(source_xml),
            self.working,
            slide_relationships,
        )?;
        Ok(Commit {
            patch: Patch {
                before: self.source,
                after: after.clone(),
            },
            snapshot: after,
        })
    }

    fn part(&self, index: usize) -> Result<&ContentPart> {
        self.working.get(index).ok_or(Error::IndexOutOfBounds {
            index,
            len: self.working.len(),
        })
    }

    fn make_part(
        &self,
        index: usize,
        anchor: Anchor,
        relationship: Relationship,
    ) -> Result<ContentPart> {
        let part = ContentPart {
            slide_index: self.source.slide_index,
            slide_part_name: self.source.slide_part_name.clone(),
            index,
            anchor,
            relationship,
        };
        validation::validate_parts(std::slice::from_ref(&part))?;
        Ok(part)
    }

    fn validate_candidate(&self, candidate: &ContentPart, index: Option<usize>) -> Result<()> {
        let mut parts = self.working.clone();
        if let Some(index) = index {
            parts[index] = candidate.clone();
        }
        validation::validate_parts(&parts)
    }

    fn validate_all(&self) -> Result<()> {
        validation::validate_parts(&self.working)
    }

    fn renumber(&mut self) {
        for (index, part) in self.working.iter_mut().enumerate() {
            part.index = index;
            part.slide_index = self.source.slide_index;
            part.slide_part_name = self.source.slide_part_name.clone();
        }
    }
}

/// A committed detached transaction.
#[derive(Debug, Clone)]
pub struct Commit {
    patch: Patch,
    snapshot: Snapshot,
}

impl Commit {
    /// Snapshot expected after publication.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source patch represented by this commit.
    #[inline]
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether this commit changes any source graph value.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.patch.is_empty()
    }

    /// Alias used by callers that prefer a verb-style result.
    #[inline]
    #[must_use]
    pub fn changed(&self) -> bool {
        self.is_changed()
    }

    /// Consume the commit and return its patch.
    #[inline]
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A source-checked, reversible publication plan.
#[derive(Debug, Clone)]
pub struct Patch {
    pub(crate) before: Snapshot,
    pub(crate) after: Snapshot,
}

impl Patch {
    /// Snapshot required before this patch can be applied.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Snapshot expected after this patch is applied.
    #[inline]
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether applying this patch is an exact no-op.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the inverse patch for immediate restoration.
    #[inline]
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch through the package's atomic publication boundary.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(package, self)
    }
}

fn unique_payloads(parts: &[ContentPart]) -> Result<Vec<Payload>> {
    let mut values = HashMap::<PackURI, Payload>::new();
    let mut total = 0usize;
    for part in parts {
        let Some(payload) = part.payload() else {
            continue;
        };
        if let Some(existing) = values.get(payload.part_name()) {
            if existing != payload {
                return Err(invalid(
                    "content-part payload declarations conflict for one part",
                ));
            }
        } else {
            total = total
                .checked_add(payload.bytes().len())
                .ok_or_else(validation::limit_total_payloads)?;
            if total > validation::MAX_TOTAL_PAYLOAD_BYTES {
                return Err(validation::limit_total_payloads());
            }
            values.insert(payload.part_name().clone(), payload.clone());
        }
    }
    let mut values: Vec<_> = values.into_values().collect();
    values
        .sort_unstable_by(|left, right| left.part_name().as_str().cmp(right.part_name().as_str()));
    Ok(values)
}

fn sorted_relationships(mut values: Vec<RelationshipState>) -> Result<Vec<RelationshipState>> {
    values.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    for pair in values.windows(2) {
        if pair[0].id == pair[1].id {
            return Err(invalid("duplicate content-part slide relationship ID"));
        }
    }
    Ok(values)
}

pub(crate) fn relationship_states<'a>(
    relationships: impl Iterator<Item = &'a litchi_opc::Relationship>,
) -> Result<Vec<RelationshipState>> {
    relationships
        .map(|relationship| {
            validation::validate_id(relationship.r_id(), "content-part slide relationship ID")?;
            validation::validate_field(
                relationship.reltype(),
                "content-part slide relationship type",
            )?;
            validation::validate_field(
                relationship.target_ref(),
                "content-part slide relationship target",
            )?;
            Ok(RelationshipState {
                id: relationship.r_id().to_owned(),
                relationship_type: relationship.reltype().to_owned(),
                target_ref: relationship.target_ref().to_owned(),
                target_mode: relationship.target_mode(),
            })
        })
        .collect()
}

fn validate_relationship_collisions(source: &Snapshot, parts: &[ContentPart]) -> Result<()> {
    let owned: HashSet<&str> = source
        .parts
        .iter()
        .map(ContentPart::relationship_id)
        .collect();
    let relationships: HashMap<&str, &RelationshipState> = source
        .slide_relationships
        .iter()
        .map(|value| (value.id.as_str(), value))
        .collect();
    for part in parts {
        if owned.contains(part.relationship_id()) {
            continue;
        }
        if relationships.contains_key(part.relationship_id()) {
            return Err(invalid(format!(
                "content-part relationship ID '{}' conflicts with an unrelated slide relationship",
                part.relationship_id()
            )));
        }
    }
    Ok(())
}

fn build_relationships(
    source: &Snapshot,
    parts: &[ContentPart],
    source_xml: &[u8],
) -> Result<Vec<RelationshipState>> {
    let mut values: HashMap<String, RelationshipState> = source
        .slide_relationships
        .iter()
        .cloned()
        .map(|value| (value.id.clone(), value))
        .collect();
    let original_ids: HashSet<&str> = source
        .parts
        .iter()
        .map(ContentPart::relationship_id)
        .collect();
    let active_ids: HashSet<&str> = parts.iter().map(ContentPart::relationship_id).collect();
    for id in original_ids {
        if !active_ids.contains(id) && !codec::contains_relationship_reference(source_xml, id) {
            values.remove(id);
        }
    }
    for part in parts {
        values.insert(
            part.relationship.id.clone(),
            RelationshipState {
                id: part.relationship.id.clone(),
                relationship_type: part.relationship.relationship_type.clone(),
                target_ref: part.relationship.target_ref.clone(),
                target_mode: part.relationship.target_mode,
            },
        );
    }
    let mut values: Vec<_> = values.into_values().collect();
    values.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(values)
}

fn build_source(
    source: &Snapshot,
    parts: &[ContentPart],
    origins: &[Option<usize>],
) -> Result<Vec<u8>> {
    let locations = codec::locate_content_parts(source.source_xml())?;
    let mut replacements = Vec::<(std::ops::Range<usize>, Vec<u8>)>::new();
    let mut inserts = HashMap::<usize, Vec<u8>>::new();
    for (origin, original) in source.parts.iter().enumerate() {
        let matching = locations
            .iter()
            .filter(|location| location.relationship_id == original.relationship_id());
        let replacement = origins
            .iter()
            .position(|value| *value == Some(origin))
            .map(|index| &parts[index]);
        for location in matching {
            if replacement.is_none() {
                replacements.push((location.range.clone(), Vec::new()));
                continue;
            }
            let replacement =
                replacement.ok_or_else(|| invalid("content-part replacement state is missing"))?;
            if replacement.anchor.xml == original.anchor.xml {
                continue;
            }
            let id_only = codec::rewrite_anchor_relationship_id(
                original.anchor.xml(),
                replacement.relationship_id(),
            )
            .is_ok_and(|value| value == replacement.anchor.xml);
            if id_only {
                replacements.push((
                    location.relationship_span.clone(),
                    escape_attribute(replacement.relationship_id()).into_bytes(),
                ));
            } else {
                replacements.push((location.range.clone(), replacement.anchor.xml.clone()));
            }
        }
    }
    for (index, (part, origin)) in parts.iter().zip(origins).enumerate() {
        if origin.is_some() {
            continue;
        }
        let offset = origins
            .iter()
            .enumerate()
            .skip(index + 1)
            .find_map(|(_, value)| {
                let origin = (*value)?;
                locations
                    .iter()
                    .find(|location| {
                        location.relationship_id == source.parts[origin].relationship_id()
                    })
                    .map(|location| location.range.start)
            })
            .unwrap_or(codec::shape_tree_insertion(source.source_xml())?);
        inserts
            .entry(offset)
            .or_default()
            .extend_from_slice(part.anchor.xml());
    }
    for (offset, value) in inserts {
        if let Some((_, replacement)) = replacements
            .iter_mut()
            .find(|(range, _)| range.start == offset && range.end > range.start)
        {
            let mut combined = value;
            combined.extend_from_slice(replacement);
            *replacement = combined;
        } else {
            replacements.push((offset..offset, value));
        }
    }
    codec::replace_spans(source.source_xml(), replacements)
}

fn fingerprint(
    source_xml: &[u8],
    parts: &[ContentPart],
    relationships: &[RelationshipState],
    payloads: &[Payload],
) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    feed_bytes(&mut hash, source_xml);
    for part in parts {
        feed_text(&mut hash, part.relationship_id());
        feed_bytes(&mut hash, part.anchor.xml());
        feed_text(&mut hash, part.relationship().relationship_type());
        feed_text(&mut hash, part.relationship().target_ref());
    }
    for relationship in relationships {
        feed_text(&mut hash, &relationship.id);
        feed_text(&mut hash, &relationship.relationship_type);
        feed_text(&mut hash, &relationship.target_ref);
        feed_mode(&mut hash, relationship.target_mode);
    }
    for payload in payloads {
        feed_text(&mut hash, payload.part_name().as_str());
        feed_text(&mut hash, payload.content_type());
        feed_bytes(&mut hash, payload.bytes());
        for relationship in payload.relationships() {
            feed_text(&mut hash, relationship.id());
            feed_text(&mut hash, relationship.relationship_type());
            feed_text(&mut hash, relationship.target_ref());
            feed_mode(&mut hash, relationship.target_mode());
        }
    }
    hash
}

fn feed_text(hash: &mut u64, value: &str) {
    feed_bytes(hash, value.as_bytes());
}

fn feed_bytes(hash: &mut u64, bytes: &[u8]) {
    for byte in bytes {
        *hash ^= u64::from(*byte);
        *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
}

fn feed_mode(hash: &mut u64, mode: TargetMode) {
    *hash ^= match mode {
        TargetMode::Internal => 0,
        TargetMode::External => 1,
    };
    *hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
}

fn escape_attribute(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
